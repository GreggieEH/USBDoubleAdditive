#include "stdafx.h"
#include "MyReggie.h"


CMyReggie::CMyReggie()
{
}


CMyReggie::~CMyReggie()
{
}

BOOL CMyReggie::GetStringValue(
	LPCTSTR				szValueName,
	LPTSTR				szValue,
	UINT				nBufferSize)
{
	HKEY			hkey = this->GetNamedKey(szValueName);
	BOOL			fSuccess = FALSE;
	DWORD			dwType;
	TCHAR			szString[MAX_PATH];
	DWORD			dwSize = MAX_PATH * sizeof(TCHAR);

	szValue[0] = '\0';
	if (NULL != hkey)
	{
		if (ERROR_SUCCESS == RegGetValue(hkey, NULL, szValueName, RRF_RT_REG_SZ, &dwType, (PVOID)szString, &dwSize))
		{
			fSuccess = TRUE;
			StringCchCopy(szValue, nBufferSize, szString);
		}
		RegCloseKey(hkey);
	}
	return fSuccess;
}

void CMyReggie::SetStringValue(
	LPCTSTR				szValueName,
	LPCTSTR				szValue)
{
	HKEY			hkey = this->GetNamedKey(szValueName);
	size_t			slen;

	if (NULL != hkey)
	{
		StringCchLength(szValue, MAX_PATH, &slen);
		RegSetKeyValue(hkey, NULL, szValueName, REG_SZ, (LPCVOID)szValue, slen * sizeof(TCHAR));
		RegCloseKey(hkey);
	}
}

HKEY CMyReggie::GetNamedKey(
	LPCTSTR				szValueName)
{
	HKEY		res = NULL;
	LONG		lresult;

	lresult = RegOpenKey(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Wow6432Node\\Sciencetech\\MonoSetup", &res);
	return res;
}
