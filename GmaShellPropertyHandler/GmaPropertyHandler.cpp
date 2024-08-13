// Adapted from Microsoft sample PlaylistPropertyHandler/PlaylistPropertyHandler.cpp

#include "dll.h"
#include "Helpers.h"
#include "RegisterExtension.h"
#include "PropertyStoreHelpers.h"
#include "cJSON.h"
#include <shobjidl.h>
#include <shlwapi.h>
#include <propvarutil.h>
#include <propkey.h>
#include <intsafe.h>

class DECLSPEC_UUID("9DBD2C50-62AD-11D0-B806-00C04FD706EC") PropertyThumbnailHandler;

const ULONGLONG c_ullGmaHeaderSizeLimit = 8192ull; // Safety catch to abort reading GMA header past this number of bytes
// Afaik there is no data in the GMA header that specifies its length, and there is no public specification I can find that specifies a length on the 3 fields it contains
// The only way to know we've reached the end of the header is to find the null byte that follows and delineates each of the 3 consecutive fields, which could be unnaturally deep in the file
// So we will abort our reading past this number of bytes as a fallback, for bizarre GMAs and non-GMAs that might otherwise fool us into reading say 3GB of null bytes
// 8192 is an arbitrary number that should be plenty long for the majority of GMAs. GMAs using the recently added "ignore" field in the json chunk with an absurd number of entries might breach this limit.

struct GmaHeaderInfoExtract
{
	PWSTR pwszName;
	PWSTR pwszAuthor;
	PWSTR pwszDescription;
	PWSTR pwszType;
	PWSTR* awszTags; // array of strings, one string for each tag
	DWORD cTags;
};

struct GmaInfo
{
	GmaHeaderInfoExtract HeaderExtract;
	BOOL HeaderUsesJsonChunkInDescription;
	PWSTR HeaderConcatForSearchContents;
};


// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    PropertyHandler
//
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CGmaPropertyHandler :
    public IPropertyStore,
    public IInitializeWithStream,
    public IPropertyStoreCapabilities,
    public IObjectProvider
{
public:
    CGmaPropertyHandler() : _cRef(1), _pStream(NULL), _pCache(NULL), _GmaInfo()
    {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void ** ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CGmaPropertyHandler, IPropertyStore),             // required
            QITABENT(CGmaPropertyHandler, IInitializeWithStream),      // required
            QITABENT(CGmaPropertyHandler, IPropertyStoreCapabilities), // optional
            QITABENT(CGmaPropertyHandler, IObjectProvider),            // optional
            {0, 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = InterlockedDecrement(&_cRef);
        if (cRef == 0)
        {
            delete this;
        }
        return cRef;
    }

    // IPropertyStore
    IFACEMETHODIMP GetCount(DWORD *pcProps);
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY *pkey);
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar);
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);
    IFACEMETHODIMP Commit();

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode);

    // IPropertyStoreCapabilities
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key)
    {
		UNREFERENCED_PARAMETER(key);
		return S_FALSE; // This is a readonly property handler. We don't support arbitrary modification of the GMA header fields.
    }
	
    // by implementing IObjectProvider clients of this property handler can
    // get to this object and access a handler specific programming model.
    // those clients use IObjectProvider::QueryObject(<object ID defined by the handler>, ...)
    // this would be useful for clients that need access to special features of
    // a particular file format but want to mostly use the IPropertyStore abstraction
    IFACEMETHODIMP QueryObject(REFGUID guidObject, REFIID riid, void **ppv)
    {
        *ppv = NULL;
        HRESULT hr = E_NOTIMPL;
        // provide access to this object directly, but in this case only
        // on IUnknown. provide access to handler specific features
        // by creating a handler specific interface that can be retrieved here
        if (guidObject == __uuidof(CGmaPropertyHandler))
        {
            hr = QueryInterface(riid, ppv);
        }
        return hr;
    }

private:
    void _ReleaseResources()
    {
        SafeRelease(&_pStream);
        SafeRelease(&_pCache);
		_ReleaseGmaHeaderInfoExtractAllocs();
    }

    ~CGmaPropertyHandler()
    {
        _ReleaseResources();
        DllRelease();
    }

    long _cRef;
    IStream *_pStream; // Stream from Initialize()
    IPropertyStoreCache *_pCache; // Storage unit for the final property values which Windows retrieves
	
	GmaInfo _GmaInfo; // Relevant information of the GMA file

	void _SafeReleaseWString(PWSTR* ppws)
	{
		if (*ppws != NULL)
		{
			delete[] *ppws;
			*ppws = NULL;
		}
	}

	void _ReleaseGmaHeaderInfoExtractAllocs()
	{
		_SafeReleaseWString(&_GmaInfo.HeaderExtract.pwszName);
		_SafeReleaseWString(&_GmaInfo.HeaderExtract.pwszAuthor);
		_SafeReleaseWString(&_GmaInfo.HeaderExtract.pwszDescription);
		_SafeReleaseWString(&_GmaInfo.HeaderExtract.pwszType);
		if (_GmaInfo.HeaderExtract.awszTags != NULL)
		{
			for (DWORD i = 0; i < _GmaInfo.HeaderExtract.cTags; i++)
				_SafeReleaseWString(&_GmaInfo.HeaderExtract.awszTags[i]);
			delete[] _GmaInfo.HeaderExtract.awszTags;
			_GmaInfo.HeaderExtract.awszTags = NULL;
			_GmaInfo.HeaderExtract.cTags = 0ul;
		}

		_SafeReleaseWString(&_GmaInfo.HeaderConcatForSearchContents);
	}


	// --------------------------------------------------
	//   Main interface
	// --------------------------------------------------
	
	//
	// GMA information retrieval
	//

    HRESULT _ReadRelevantGmaData();
	HRESULT _ReadGmaHeaderField(ULONGLONG* ullStreamSize, ULONGLONG* ullStreamPos, BYTE** pcBuf);
	HRESULT _ReadAndParseStringGmaHeaderField(ULONGLONG* ullStreamSize, ULONGLONG* ullStreamPos, PSTR* pszOutField, PWSTR* pszWString);
	HRESULT _PostProcessGmaData();

	//
	// Transport to Property value store
	//

	HRESULT _SendGmaDataToPropertyStore();
	HRESULT _SetWstrPropertyValueInPropertyCache(PWSTR pwszPropValue, PROPERTYKEY propKey);
	HRESULT _SetWstrArrayPropertyValueInPropertyCache(LPWSTR* apwstrPropValue, ULONG cPropValue, PROPERTYKEY propKey);
	bool _SearchContentStringConcat(PWSTR pwszSearchContents, PWSTR pwszInsertString, const WCHAR wcSuffixChar, size_t* pullInsertPos);
	
};


// ____________________________________________________________________________________________________
// 
//     COM interfaces
// ____________________________________________________________________________________________________
//

HRESULT CGmaPropertyHandler_CreateInstance(REFIID riid, void **ppv)
{
    *ppv = NULL;
    CGmaPropertyHandler *pirm = new (std::nothrow) CGmaPropertyHandler();
    HRESULT hr = pirm ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = pirm->QueryInterface(riid, ppv);
        pirm->Release();
    }
    return hr;
}

// --------------------------------------------------
//   IPropertyStore
// --------------------------------------------------

HRESULT CGmaPropertyHandler::GetCount(DWORD *pcProps)
{
    *pcProps = 0;
    return _pCache ? _pCache->GetCount(pcProps) : E_UNEXPECTED;
}

HRESULT CGmaPropertyHandler::GetAt(DWORD iProp, PROPERTYKEY *pkey)
{
    *pkey = PKEY_Null;
    return _pCache ? _pCache->GetAt(iProp, pkey) : E_UNEXPECTED;
}

HRESULT CGmaPropertyHandler::GetValue(REFPROPERTYKEY key, PROPVARIANT *pPropVar)
{
    PropVariantInit(pPropVar);
    return _pCache ? _pCache->GetValue(key, pPropVar) : E_UNEXPECTED;
}

// SetValue just updates the internal value cache
HRESULT CGmaPropertyHandler::SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar)
{
	UNREFERENCED_PARAMETER(key);
	UNREFERENCED_PARAMETER(propVar);
	return STG_E_ACCESSDENIED; // This is a readonly property handler
}

HRESULT CGmaPropertyHandler::Commit()
{
    return STG_E_ACCESSDENIED; // This is a readonly property handler
}


// --------------------------------------------------
//   IInitializeWithStream
// --------------------------------------------------

HRESULT CGmaPropertyHandler::Initialize(IStream *pStream, DWORD grfMode)
{
	UNREFERENCED_PARAMETER(grfMode);

    HRESULT hr = E_UNEXPECTED;
	
    if (!_pStream)
    {
		// Save a reference to the stream
        hr = pStream->QueryInterface(&_pStream);
        if (FAILED(hr))
			return hr;

		// Create the property cache for storing our property values
		hr = PSCreateMemoryPropertyStore(IID_PPV_ARGS(&_pCache));
		if (FAILED(hr))
		{
			_ReleaseResources();
			return hr;
		}

		// Read the property data we want from the GMA file
		hr = _ReadRelevantGmaData();
		if (FAILED(hr))
		{
			_ReleaseResources();
			return hr;
		}

		// Clean up the data
		hr = _PostProcessGmaData();
		if (FAILED(hr))
		{
			_ReleaseResources();
			return hr;
		}

		// Send the read-in data to our property cache for Windows to read
		hr = _SendGmaDataToPropertyStore();
		if (FAILED(hr))
		{
			_ReleaseResources();
			return hr;
		}
    }
    return hr;
}


// ____________________________________________________________________________________________________
// 
//     GMA data transport to Property value store
// ____________________________________________________________________________________________________
//

/*
	Mapping of GMA properties to the predefined Windows Properties

	---	Example GMA Data ---------------------------        ---- Windows Property ------------

	`name` header field								   ->	System.Title

	`author` header field							   ->	System.Author

	JSON blob in `description` header field
	{
		"description": "Description",				   ->	System.Link.Description  (bizarrely, there is no generic or document-based description property, so this is the closest)
		"type": "weapon",							   ->	System.Category  (Type is already used for extension type, and Category is the closest semantic match)
		"tags": [									   ->	System.Keywords  (aka tags)
			"fun",
			"roleplay"
		]
	}

*/

HRESULT CGmaPropertyHandler::_SendGmaDataToPropertyStore()
{
	HRESULT hr = E_UNEXPECTED;
	
	//
	// Main properties
	//
	
	// Name (string)
	hr = _SetWstrPropertyValueInPropertyCache(_GmaInfo.HeaderExtract.pwszName, PKEY_Title); // method handles fallback to empty string for NULL
	if (FAILED(hr))
		return hr;
	
	// Author (multi-value string)
	PWSTR apwszDescription[1] = { _GmaInfo.HeaderExtract.pwszAuthor };
	hr = _SetWstrArrayPropertyValueInPropertyCache(_GmaInfo.HeaderExtract.pwszAuthor ? apwszDescription : NULL, 1ul, PKEY_Author); // ditto
	if (FAILED(hr))
		return hr;
	
	// Description (string)
	hr = _SetWstrPropertyValueInPropertyCache(_GmaInfo.HeaderExtract.pwszDescription, PKEY_Link_Description);
	if (FAILED(hr))
		return hr;
	
	// Type (multi-value string)
	PWSTR apwszType[1] = { _GmaInfo.HeaderExtract.pwszType ? _GmaInfo.HeaderExtract.pwszType : L"" };
	hr = _SetWstrArrayPropertyValueInPropertyCache(_GmaInfo.HeaderExtract.pwszAuthor ? apwszType : NULL, 1ul, PKEY_Category);
	if (FAILED(hr))
		return hr;
	
	// Tags (multi-value string)
	hr = _SetWstrArrayPropertyValueInPropertyCache(_GmaInfo.HeaderExtract.cTags > 0ul ? _GmaInfo.HeaderExtract.awszTags : NULL, _GmaInfo.HeaderExtract.cTags, PKEY_Keywords);
	if (FAILED(hr))
		return hr;

	//
	// "Contents" string for the search indexer
	//

	// We will concatenate all strings together, using newlines to separate each property and whitespace to separate each item within multi-string properties
	ULONG cSearchContents = 0ul;
	if (_GmaInfo.HeaderExtract.pwszName != NULL)
		cSearchContents += lstrlenW(_GmaInfo.HeaderExtract.pwszName) + 1ul; // +1 for trailing \n
	
	if (_GmaInfo.HeaderExtract.pwszAuthor != NULL)
		cSearchContents += lstrlenW(_GmaInfo.HeaderExtract.pwszAuthor) + 1ul; // +1 for trailing \n
	
	if (_GmaInfo.HeaderExtract.pwszDescription != NULL)
		cSearchContents += lstrlenW(_GmaInfo.HeaderExtract.pwszDescription) + 1ul; // +1 for trailing \n
	
	if (_GmaInfo.HeaderExtract.pwszType != NULL)
		cSearchContents += lstrlenW(_GmaInfo.HeaderExtract.pwszType) + 1ul; // +1 for trailing \n

	if (_GmaInfo.HeaderExtract.cTags > 0ul && _GmaInfo.HeaderExtract.awszTags != NULL)
	{
		for (ULONG i = 0ul; i < _GmaInfo.HeaderExtract.cTags; i++)
			cSearchContents += lstrlenW(_GmaInfo.HeaderExtract.awszTags[i]) + 1ul; // +1 for trailing space
		cSearchContents += 1; // trailing \n
	}

	_GmaInfo.HeaderConcatForSearchContents = new WCHAR[cSearchContents + 1]();
	size_t ullInsertPos = 0;
	_SearchContentStringConcat(_GmaInfo.HeaderConcatForSearchContents, _GmaInfo.HeaderExtract.pwszName, L'\n', &ullInsertPos);
	_SearchContentStringConcat(_GmaInfo.HeaderConcatForSearchContents, _GmaInfo.HeaderExtract.pwszAuthor, L'\n', &ullInsertPos);
	_SearchContentStringConcat(_GmaInfo.HeaderConcatForSearchContents, _GmaInfo.HeaderExtract.pwszDescription, L'\n', &ullInsertPos);
	_SearchContentStringConcat(_GmaInfo.HeaderConcatForSearchContents, _GmaInfo.HeaderExtract.pwszType, L'\n', &ullInsertPos);
	if (_GmaInfo.HeaderExtract.cTags > 0ul && _GmaInfo.HeaderExtract.awszTags != NULL)
	{
		for (ULONG i = 0ul; i < _GmaInfo.HeaderExtract.cTags; i++)
		{
			WCHAR wcSuffixChar = (i < _GmaInfo.HeaderExtract.cTags - 1ul) ? L' ' : L'\n';
			_SearchContentStringConcat(_GmaInfo.HeaderConcatForSearchContents, _GmaInfo.HeaderExtract.awszTags[i], wcSuffixChar, &ullInsertPos);
		}
	}

	hr = _SetWstrPropertyValueInPropertyCache(_GmaInfo.HeaderConcatForSearchContents, PKEY_Search_Contents);
	if (FAILED(hr))
		return hr;

	// All properties set
	return hr;
}

HRESULT CGmaPropertyHandler::_SetWstrPropertyValueInPropertyCache(PWSTR pwszPropValue, PROPERTYKEY propKey)
{
	HRESULT hr = E_UNEXPECTED;

	if (_pCache == NULL)
		return hr;

	PROPVARIANT propvar = {};
	InitPropVariantFromString( pwszPropValue ? (LPCWSTR)pwszPropValue : L"", &propvar); // Fallback to dummy empty string, for when pwszPropValue is null

	hr = PSCoerceToCanonicalValue(propKey, &propvar);
	if (FAILED(hr))
		return hr;

	if (pwszPropValue != NULL) // If property was extracted from the GMA, set it in the prop store normally
		hr = _pCache->SetValueAndState(propKey, &propvar, PSC_NORMAL);
	else // If the GMA was missing it, store the dummy empty value and report it as absent
		hr = _pCache->SetValueAndState(propKey, &propvar, PSC_NOTINSOURCE);
	
	PropVariantClear(&propvar);

	return hr;
}

HRESULT CGmaPropertyHandler::_SetWstrArrayPropertyValueInPropertyCache(LPWSTR* apwstrPropValue, ULONG cPropValue, PROPERTYKEY propKey)
{
	HRESULT hr = E_UNEXPECTED;

	if (_pCache == NULL)
		return hr;

	PROPVARIANT propvar = {};
	if (apwstrPropValue != NULL)
	{
		InitPropVariantFromStringVector((LPCWSTR*)apwstrPropValue, cPropValue, &propvar);
	}
	else
	{
		LPCWSTR dummy[1] = { L"" };
		InitPropVariantFromStringVector(dummy, cPropValue, &propvar);
	}

	hr = PSCoerceToCanonicalValue(propKey, &propvar);
	if (FAILED(hr))
		return hr;

	if (apwstrPropValue != NULL) // If property was extracted from the GMA, set it in the prop store normally
		hr = _pCache->SetValueAndState(propKey, &propvar, PSC_NORMAL);
	else // If the GMA was missing it, store the dummy empty value and report it as absent
		hr = _pCache->SetValueAndState(propKey, &propvar, PSC_NOTINSOURCE);
	
	PropVariantClear(&propvar);

	return hr;
}

bool CGmaPropertyHandler::_SearchContentStringConcat(PWSTR pwszSearchContents, PWSTR pwszInsertString, const WCHAR wcSuffixChar, size_t* pullInsertPos)
{
	if (pwszInsertString != NULL)
	{
		size_t cInsertString = lstrlenW(pwszInsertString);
		wcsncpy_s( &(pwszSearchContents[*pullInsertPos]), cInsertString + 1, pwszInsertString, cInsertString  );
		*pullInsertPos += cInsertString;

		pwszSearchContents[*pullInsertPos] = wcSuffixChar;
		*pullInsertPos += 1ul;

		return true;
	}

	return false;
}


// ____________________________________________________________________________________________________
// 
//     GMA information retrieval
// ____________________________________________________________________________________________________
//
	
/*
	Header examples

	159321088.gma
		- Old(er) gma header format, back when `description` actually had a description in it, before it was turned into a json chunk
		- `author` field is always* constant "author" string

		0000h  47 4D 41 44 03 1E 18 19 08 01 00 10 01 0C 11 DF 51 00 00 00 00 00 74 74  GMAD...........ßQ.....tt 
		0018h  74 5F 6D 69 6E 65 63 72 61 66 74 5F 62 35 00 4E 6F 20 63 72 65 64 69 74  t_minecraft_b5.No credit 
		0030h  20 74 6F 20 6D 65 20 69 20 6A 75 73 74 20 75 70 6C 6F 61 64 65 64 20 61   to me i just uploaded a 
		0048h  6C 6C 20 63 72 65 64 69 74 20 74 6F 20 20 20 20 20 28 3D 43 47 3D 29 20  ll credit to     (=CG=)  
		0060h  46 69 6E 6E 69 65 73 70 2E 20 41 20 54 72 6F 75 62 6C 65 20 49 6E 20 54  Finniesp. A Trouble In T 
		0078h  65 72 72 6F 72 69 73 74 20 54 6F 77 6E 20 4D 43 20 62 65 74 61 20 6D 61  errorist Town MC beta ma 
		0090h  70 2E 20 49 66 20 74 68 65 20 6F 77 6E 65 72 20 77 61 6E 74 73 20 69 74  p. If the owner wants it 
		00A8h  20 6F 66 66 20 74 68 65 20 77 6F 72 6B 73 68 6F 70 20 68 65 20 63 61 6E   off the workshop he can 
		00C0h  20 6A 75 73 74 20 6C 65 61 76 65 20 61 20 63 6F 6D 6D 65 6E 74 2E 0A 20   just leave a comment..  
		00D8h  20 20 20 0A 00 61 75 74 68 6F 72 00 01 00 00 00 01 00 00 00 6D 61 70 73     ..author.........maps 
		00F0h  2F 74 74 74 5F 6D 69 6E 65 63 72 61 66 74 5F 62 35 2E 62 73 70 00 9A 6C  /ttt_minecraft_b5.bsp.šl 
		0108h  53 01 00 00 00 00 9B 50 51 93 00 00 00 00 56 42 53 50 14 00 00 00 04 48  S.....›PQ“....VBSP.....H 

	124648402.gma
		- New(er) gma header format, with json chunk in `description`
		- `author` field is now always* a constant "Author Name" string

		0000h  47 4D 41 44 03 00 00 00 00 00 00 00 00 E4 78 0B 5D 00 00 00 00 00 4F 72  GMAD.........äx.].....Or 
		0018h  62 69 74 61 6C 20 46 72 69 65 6E 64 73 68 69 70 20 43 61 6E 6E 6F 6E 00  bital Friendship Cannon. 
		0030h  7B 0A 09 22 64 65 73 63 72 69 70 74 69 6F 6E 22 3A 20 22 44 65 73 63 72  {.."description": "Descr 
		0048h  69 70 74 69 6F 6E 22 2C 0A 09 22 74 79 70 65 22 3A 20 22 77 65 61 70 6F  iption",.."type": "weapo 
		0060h  6E 22 2C 0A 09 22 74 61 67 73 22 3A 20 5B 0A 09 09 22 66 75 6E 22 2C 0A  n",.."tags": [..."fun",. 
		0078h  09 09 22 63 61 72 74 6F 6F 6E 22 0A 09 5D 0A 7D 00 41 75 74 68 6F 72 20  .."cartoon"..].}.Author  
		0090h  4E 61 6D 65 00 01 00 00 00 01 00 00 00 6C 75 61 2F 77 65 61 70 6F 6E 73  Name.........lua/weapons 
		00A8h  2F 6F 72 62 69 74 61 6C 5F 66 72 69 65 6E 64 73 68 69 70 5F 63 61 6E 6E  /orbital_friendship_cann 
		00C0h  6F 6E 2E 6C 75 61 00 CB 3C 00 00 00 00 00 00 37 E2 3B 90 02 00 00 00 6D  on.lua.Ë<......7â;
	
	*Some gmas have actual strings in `author`. gmad.exe does not support this, so these gmas were modified or produced through alternate means.

	The header fields are delineated by null bytes.
	There are always exactly header three fields:
	- Name
	- Author
	- Description
*/

// Load the property data we want from the current file stream to the GMA file
HRESULT CGmaPropertyHandler::_ReadRelevantGmaData()
{
	HRESULT hr = E_UNEXPECTED;


	//
	// Seek and stat the stream
	//

	LARGE_INTEGER liSeek;
	liSeek.QuadPart = 0;
	hr = _pStream->Seek(liSeek, STREAM_SEEK_SET, NULL);
	if (FAILED(hr))
		return hr;

	STATSTG statstg;
	hr = _pStream->Stat(&statstg, STATFLAG_NONAME);
	if (FAILED(hr))
		return hr;

	ULONGLONG ullStreamSize = statstg.cbSize.QuadPart;
	ULONGLONG ullStreamPos = 0ull;


	//
	// Verify 4 byte magic
	//

	if (ullStreamSize < 4ull)
		return E_ABORT;

	BYTE cBufMagic[5];
	ULONG ulBytesRead = 0;
	hr = _pStream->Read(&cBufMagic, 4, &ulBytesRead);
	if (FAILED(hr))
		return hr;
	if (ulBytesRead < 4ul)
		return E_UNEXPECTED;
	
	cBufMagic[4] = 0;

	ullStreamPos += ulBytesRead;

	if (strcmp((char*)cBufMagic, "GMAD") != 0) // First four bytes are "GMAD" in ascii
		return E_ABORT;


	//
	// Jump to relevant data
	//

	// Skip 13 unknown bytes that start with nulls then end with something
	ullStreamPos += 13ull;
	if (ullStreamPos >= ullStreamSize)
		return E_ABORT;

	liSeek.QuadPart = 13ll;
	hr = _pStream->Seek(liSeek, STREAM_SEEK_CUR, NULL);
	if (FAILED(hr))
		return hr;

	// Skip nulls until we reach the first non-null byte
	// This will be the first character of the addon name
	BYTE cBuf1 = 0;
	while (cBuf1 == 0 && ullStreamPos < ullStreamSize && ullStreamPos < c_ullGmaHeaderSizeLimit)
	{
		hr = _pStream->Read(&cBuf1, 1ul, &ulBytesRead);
		if (FAILED(hr))
			return hr;
		if (ulBytesRead == 0)
			return E_UNEXPECTED;

		ullStreamPos += ulBytesRead;
	}
	
	if (ullStreamPos >= ullStreamSize)
		return E_UNEXPECTED;
	if (ullStreamPos >= c_ullGmaHeaderSizeLimit)
		return E_ABORT;
	
	// Back up to field start
	liSeek.QuadPart = -1ll;
	hr = _pStream->Seek(liSeek, STREAM_SEEK_CUR, NULL);
	if (FAILED(hr))
		return hr;
	ullStreamPos -= 1ull;


	//
	// Parse `name` header field
	//

	// Read in
	PWSTR pwszAddonName;
	hr = _ReadAndParseStringGmaHeaderField(&ullStreamSize, &ullStreamPos, NULL, &pwszAddonName);
	if (FAILED(hr))
		return hr;

	_GmaInfo.HeaderExtract.pwszName = pwszAddonName;


	//
	// Parse `description` field
	//
	
	// Read in
	PSTR pszAddonDescription; // needed for cJSON, in the case of json chunk
	PWSTR pwszAddonDescription;
	hr = _ReadAndParseStringGmaHeaderField(&ullStreamSize, &ullStreamPos, &pszAddonDescription, &pwszAddonDescription);
	if (FAILED(hr))
	{
		_ReleaseGmaHeaderInfoExtractAllocs();
		return hr;
	}
	
	// Detect json chunk in this field and handle it
	// Afaik there is no indicator in the GMA header that states whether the bytes in `description` are an old(er)-format normal string or a new(er)-format json chunk
	// And I do not know how ugc/gmod does the json or not detection
	// So we will simply check for a leading '{' and trailing '}' and try to parse it as json if both are found

	BOOL bFoundLeadingBrace = false;
	BOOL bFoundTrailingBrace = false;
	
	// Search for leading brace
	size_t cAddonDescription;
	IntToSizeT(lstrlenW(pwszAddonDescription), &cAddonDescription);
	for (size_t i = 0; i < cAddonDescription; i++)
	{
		WCHAR wcCur = pwszAddonDescription[i];

		if (iswspace(wcCur))
			continue;

		if (wcCur == L'{')
		{
			bFoundLeadingBrace = true;
			break;
		}
	}
	
	// Search for trailing brace
	if (bFoundLeadingBrace)
	{
		for (size_t i = cAddonDescription; i >= 0; i--)
		{
			WCHAR wcCur = pwszAddonDescription[i];

			if (iswspace(wcCur))
				continue;

			if (wcCur == L'}')
			{
				bFoundTrailingBrace = true;
				break;
			}
		}
	}

	// Try to parse json object from header field
	cJSON* pDescriptionJson = NULL;

	if (bFoundLeadingBrace && bFoundTrailingBrace)
	{
		// JSON is a Unicode standard, yet cJSON expects a UTF8 encoded string when parsing json bytes
		pDescriptionJson = cJSON_Parse(pszAddonDescription);

		// Return will be null if parsing failed
		// cJSON_GetErrorPtr() will have details on why
		// Could be malformed json, could be data that is not json
		// gmad.exe presumably validates addon.json on gma creation, so the former is unlikely, but still possible
		// In any regard, if the parsing fails, we will treat the body of the `description` field as basic text, i.e. the old(er) format before the json chunk introduction
	}

	if (pDescriptionJson == NULL)
	{
		// Treat as basic text
		_GmaInfo.HeaderUsesJsonChunkInDescription = false;
		_GmaInfo.HeaderExtract.pwszDescription = pwszAddonDescription;
	}
	else
	{
		_GmaInfo.HeaderUsesJsonChunkInDescription = true;

		// Extract "description" from json chunk
		cJSON* pjnDescription = cJSON_GetObjectItem(pDescriptionJson, "description");
		if (cJSON_IsString(pjnDescription) && (pjnDescription->valuestring != NULL))
		{
			PWSTR pwszJsonDescription;
			hr = ConvertMultiByteStringToWide( pjnDescription->valuestring, (int)strlen(pjnDescription->valuestring), &pwszJsonDescription, CP_UTF8 );
			if (FAILED(hr))
			{
				cJSON_Delete(pDescriptionJson);
				_ReleaseGmaHeaderInfoExtractAllocs();
				return hr;
			}

			_GmaInfo.HeaderExtract.pwszDescription = pwszJsonDescription;
		}

		// Extract "type" from json chunk
		cJSON* pjnType = cJSON_GetObjectItem(pDescriptionJson, "type");
		if (cJSON_IsString(pjnType) && (pjnType->valuestring != NULL))
		{
			PWSTR pwszJsonType;
			hr = ConvertMultiByteStringToWide( pjnType->valuestring, (int)strlen(pjnType->valuestring), &pwszJsonType, CP_UTF8 );
			if (FAILED(hr))
			{
				cJSON_Delete(pDescriptionJson);
				_ReleaseGmaHeaderInfoExtractAllocs();
				return hr;
			}

			_GmaInfo.HeaderExtract.pwszType = pwszJsonType;
		}

		// Extact the "tags" list from json chunk
		cJSON* pjnTags = cJSON_GetObjectItem(pDescriptionJson, "tags");
		if (cJSON_IsArray(pjnTags))
		{
			int cTags = cJSON_GetArraySize(pjnTags);
			if (cTags > 0ul)
			{
				PWSTR* pawszTags = new PWSTR[cTags];

				for (int i = 0; i < cTags; i++)
				{
					cJSON* pjnTagsTag = cJSON_GetArrayItem(pjnTags, i);
					if (cJSON_IsString(pjnTagsTag) && (pjnTagsTag->valuestring != NULL))
					{
						PWSTR pwszJsonTag;
						hr = ConvertMultiByteStringToWide( pjnTagsTag->valuestring, (int)strlen(pjnTagsTag->valuestring), &pwszJsonTag, CP_UTF8 );
						if (FAILED(hr))
						{
							for (int j = 0; j < i; j++)
								delete pawszTags[j];
							delete[] pawszTags;
							cJSON_Delete(pDescriptionJson);
							_ReleaseGmaHeaderInfoExtractAllocs();
							return hr;
						}

						pawszTags[i] = pwszJsonTag;
					}
				}

				_GmaInfo.HeaderExtract.awszTags = pawszTags;
				IntToDWord(cTags, &_GmaInfo.HeaderExtract.cTags);
			}
		}
	}


	//
	// Parse `author` header field
	//

	PWSTR pwszAddonAuthor;
	hr = _ReadAndParseStringGmaHeaderField(&ullStreamSize, &ullStreamPos, NULL, &pwszAddonAuthor);
	if (FAILED(hr))
		return hr;

	_GmaInfo.HeaderExtract.pwszAuthor = pwszAddonAuthor;


	// All required data has been read from the GMA file
	hr = S_OK;
    return hr;
}

HRESULT CGmaPropertyHandler::_ReadGmaHeaderField(ULONGLONG* ullStreamSize, ULONGLONG* ullStreamPos, BYTE** pcBuf)
{
	HRESULT hr = E_UNEXPECTED;

	ULONGLONG ullFieldStartPos = *ullStreamPos;

	//
	// Seek to the next null byte, which marks the end of this field and the start of the next one
	//

	ULONG ulBytesRead;
	BYTE cBuf1 = 1;
	while (cBuf1 != 0 && *ullStreamPos < *ullStreamSize && *ullStreamPos < c_ullGmaHeaderSizeLimit)
	{
		hr = _pStream->Read(&cBuf1, 1ul, &ulBytesRead);
		if (FAILED(hr))
			return hr;
		if (ulBytesRead == 0)
			return E_UNEXPECTED;

		*ullStreamPos += ulBytesRead;
	}

	if (*ullStreamPos >= *ullStreamSize)
		return E_UNEXPECTED;
	if (*ullStreamPos >= c_ullGmaHeaderSizeLimit)
		return E_ABORT;

	//
	// Copy the field, including its terminating null byte
	//

	ULONGLONG ullFieldLength = *ullStreamPos - ullFieldStartPos;

	LARGE_INTEGER liRewindAmount;
	LONGLONG llFieldLength;
	hr = ULongLongToLongLong(ullFieldLength, &llFieldLength);
	if (FAILED(hr))
		return hr;
	liRewindAmount.QuadPart = llFieldLength * -1ll;

	ULARGE_INTEGER uliRewindResultPos;
	hr = _pStream->Seek(liRewindAmount, STREAM_SEEK_CUR, &uliRewindResultPos);
	if (FAILED(hr))
		return hr;
	*ullStreamPos = uliRewindResultPos.QuadPart;

	ULONG ulReadLength;
	hr = ULongLongToULong(ullFieldLength, &ulReadLength);
	if (FAILED(hr))
		return hr;

	BYTE* pcBufFull = new BYTE[ulReadLength];
	hr = _pStream->Read(pcBufFull, ulReadLength, &ulBytesRead);
	if (FAILED(hr))
		return hr;

	*pcBuf = pcBufFull;

	return S_OK;
}

HRESULT CGmaPropertyHandler::_ReadAndParseStringGmaHeaderField(ULONGLONG* ullStreamSize, ULONGLONG* ullStreamPos, PSTR* pszOutField, PWSTR* pwszOutField)
{
	HRESULT hr = E_UNEXPECTED;

	//
	// Read field
	//

	BYTE* pcFieldBytes;
	hr = _ReadGmaHeaderField(ullStreamSize, ullStreamPos, &pcFieldBytes);
	if (FAILED(hr))
		return hr;

	//
	// Convert raw bytes to unicode strings
	//
	
	// Note that the very limited design of the gma header format (using \0 to delineate strings) implies that these strings cannot possibly be in any multi-byte encoding
	// In all likelihood these are UTF8 strings

	PWSTR pwszFieldBytes = NULL;
	hr = ConvertMultiByteStringToWide( (char*)pcFieldBytes, (int)strlen((char*)pcFieldBytes), &pwszFieldBytes, CP_UTF8 );
	if (FAILED(hr))
	{
		delete[] pwszFieldBytes;
		return hr;
	}

	if (pszOutField != NULL)
		*pszOutField = (PSTR)pcFieldBytes;

	if (pwszOutField != NULL)
		*pwszOutField = pwszFieldBytes;

	return S_OK;
}

// Various adjustments to the data read in from the GMA
HRESULT CGmaPropertyHandler::_PostProcessGmaData()
{
	HRESULT hr = E_UNEXPECTED;

	//
	// Remove stubs
	//

	// Some fields are unused and populated with meaningless stub values. We're going to remove those, so the associated Property is blank instead.

	// Author
	// In the old(er) GMA format, the `author` header field is always "author"
	// In the new(er) format, the "author" element in the json chunk is always "Author Name"
	if (_GmaInfo.HeaderExtract.pwszAuthor != NULL)
	{
		if (StrCmpW(_GmaInfo.HeaderExtract.pwszAuthor, _GmaInfo.HeaderUsesJsonChunkInDescription ? L"Author Name" : L"author") == 0)
		{
			delete _GmaInfo.HeaderExtract.pwszAuthor;
			_GmaInfo.HeaderExtract.pwszAuthor = new WCHAR[1]();
		}
	}

	// Description
	// In the new(er) GMA format, the "description" element in the json chunk is always "Description"
	if (_GmaInfo.HeaderExtract.pwszDescription != NULL)
	{
		if (_GmaInfo.HeaderUsesJsonChunkInDescription && StrCmpW(_GmaInfo.HeaderExtract.pwszDescription, L"Description") == 0)
		{
			delete _GmaInfo.HeaderExtract.pwszDescription;
			_GmaInfo.HeaderExtract.pwszDescription = new WCHAR[1]();
		}
	}

	hr = S_OK;
	return hr;
}



// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//    Component registration
//
// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

HRESULT RegisterPropLists(CRegisterExtension &re, PCWSTR pszFileAssoc)
{
    const WCHAR c_szPropertiesTileInfo[] = L"prop:System.Title;System.Author;System.Size;System.DateModified";
    const WCHAR c_szPropertiesPreviewDetails[] = L"prop:System.Title;System.Author;System.Link.Description;System.Category;System.Keywords;System.DateModified";
    const WCHAR c_szPropertiesInfoTip[] = L"prop:System.Title;System.Author;System.DateModified;System.Size";
    const WCHAR c_szPropertiesFullDetails[] = L"prop:System.Title;System.PropGroup.Description;System.Author;System.Link.Description;System.Category;System.Keywords;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
    const WCHAR c_szPropertiesExtendedTileInfo[] = L"prop:System.ItemType;System.Size;System.Link.Description;System.Category;System.Keywords";

    HRESULT hr = re.RegisterProgIDValue(pszFileAssoc, L"TileInfo", c_szPropertiesTileInfo);
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterProgIDValue(pszFileAssoc, L"PreviewDetails", c_szPropertiesPreviewDetails);
        if (SUCCEEDED(hr))
        {
            hr = re.RegisterProgIDValue(pszFileAssoc, L"InfoTip", c_szPropertiesInfoTip);
            if (SUCCEEDED(hr))
            {
                hr = re.RegisterProgIDValue(pszFileAssoc, L"FullDetails", c_szPropertiesFullDetails);
                if (SUCCEEDED(hr))
                {
                    hr = re.RegisterProgIDValue(pszFileAssoc, L"ExtendedTileInfo", c_szPropertiesExtendedTileInfo);
                }
            }
        }
    }
    return hr;
}

const WCHAR c_szGmaFileExtension[] = L".gma";
const WCHAR c_szGmaProgID[] = L"GmaShellInfo.GMA.1";

HRESULT RegisterHandler()
{
    // register the property handler COM object, and set the options it uses
    const WCHAR c_szPropertyHandlerDescription[] = L"Garry's Mod GMA (.gma) Property Handler";
    CRegisterExtension re(__uuidof(CGmaPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.RegisterInProcServer(c_szPropertyHandlerDescription, L"Both");

    // Property Handler registrations use a different mechanism than the rest of the filetype association system, and do not use ProgIDs
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterPropertyHandler(c_szGmaFileExtension);
    }

    // Associate our ProgIDs with the file extensions, and write the proplists to the ProgID to minimize conflicts with other applications and facilitate easy unregistration
    if (SUCCEEDED(hr))
    {
        hr = re.RegisterExtensionWithProgID(c_szGmaFileExtension, c_szGmaProgID);
        if (SUCCEEDED(hr))
        {
            hr = RegisterPropLists(re, c_szGmaProgID);
        }
    }

    // also register the property-driven thumbnail handler on the ProgIDs
    if (SUCCEEDED(hr))
    {
        re.SetHandlerCLSID(__uuidof(PropertyThumbnailHandler));
        hr = re.RegisterThumbnailHandler(c_szGmaProgID);
    }
    return hr;
}

HRESULT UnregisterHandler()
{
    // Unregister the property handler COM object.
    CRegisterExtension re(__uuidof(CGmaPropertyHandler), HKEY_LOCAL_MACHINE);
    HRESULT hr = re.UnRegisterObject();
    if (SUCCEEDED(hr))
    {
        // Unregister the property handler the file extensions.
        hr = re.UnRegisterPropertyHandler(c_szGmaFileExtension);
    }

    if (SUCCEEDED(hr))
    {
        // Remove the whole ProgIDs since we own all of those settings.
        // Don't try to remove the file extension association since some other application may have overridden it with their own ProgID in the meantime.
        // Leaving the association to a non-existing ProgID is handled gracefully by the Shell.
        // NOTE: If the file extension is unambiguously owned by this application, the association to the ProgID could be safely removed as well,
        //       along with any other association data stored on the file extension itself.
        hr = re.UnRegisterProgID(c_szGmaProgID, c_szGmaFileExtension);
    }
    return hr;
}
