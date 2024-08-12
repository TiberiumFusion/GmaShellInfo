#include <Windows.h>
#include "Helpers.h"

HRESULT ConvertMultiByteStringToWide(PSTR psMbString, int cbsMbString, PWSTR* ppwszConverted, long lCodePage)
{
	int cbWString = MultiByteToWideChar(lCodePage, 0, psMbString, cbsMbString, NULL, 0);

	WCHAR* pwszConverted = new WCHAR[cbWString + 1];
	int iResult = MultiByteToWideChar(lCodePage, 0, psMbString, cbsMbString, pwszConverted, cbWString);
	if (iResult == 0 && cbWString > 0)
	{
		DWORD dwError = GetLastError();

		delete[] pwszConverted;

		if (dwError != 0)
			return HRESULT_FROM_WIN32(dwError);
		else
			return E_UNEXPECTED;
	}

	pwszConverted[cbWString] = 0;

	*ppwszConverted = pwszConverted;

	return S_OK;
}
