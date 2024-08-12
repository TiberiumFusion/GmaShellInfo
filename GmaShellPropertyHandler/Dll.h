// Adapted from Microsoft sample PlaylistPropertyHandler/Dll.h

// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

#define NTDDI_VERSION NTDDI_VISTA
#define _WIN32_WINNT _WIN32_WINNT_VISTA

#include <windows.h>
#include <new>  // std::nothrow

class __declspec(uuid("6fcf7e65-467e-4659-a480-64e6e8aa7887")) CGmaPropertyHandler;
HRESULT CGmaPropertyHandler_CreateInstance(REFIID riid, void **ppv);
HRESULT RegisterHandler();
HRESULT UnregisterHandler();

void DllAddRef();
void DllRelease();
