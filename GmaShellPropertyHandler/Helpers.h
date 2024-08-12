#pragma once
#include <Windows.h>

HRESULT ConvertMultiByteStringToWide(PSTR psMbString, int cbsMbString, PWSTR* ppwszConverted, long lCodePage);
