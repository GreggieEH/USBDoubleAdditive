#include "stdafx.h"
#include "DateCheck.h"
#include "MyCrypto.h"
#include "MyReggie.h"
//#include "MyCNG.h"

CDateCheck::CDateCheck()
{
}


CDateCheck::~CDateCheck()
{
}

BOOL CDateCheck::MyDateCheck()
{
//	CMyCNG			cng;					// cryptography next generation
	MyCrypto		crypt;
	CMyReggie		reggie;
	TCHAR			szString[MAX_PATH];
	TCHAR			szLastOpDate[MAX_PATH];
	TCHAR			szExpDate[MAX_PATH];
	TCHAR			szDifference[MAX_PATH];
	BOOL			fSuccess = FALSE;
	BYTE		*	pByte;
	ULONG			nBytes;
	LPTSTR			szTemp = NULL;
	long			diffCount;				// difference count

	crypt.DoInit();
	ZeroMemory((PVOID)szLastOpDate, MAX_PATH * sizeof(TCHAR));
	// read the last op date
	if (reggie.GetStringValue(L"val1", szString, MAX_PATH))
	{
		// convert to byte
		pByte = NULL;
		crypt.TranslateRegistryString(szString, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.DecriptStringFromBytes_Aes(pByte, nBytes, szLastOpDate, MAX_PATH);
//			cng.DecryptCipher(pByte, nBytes, szLastOpDate, MAX_PATH);
			CoTaskMemFree((LPVOID)pByte);
		}
	}
	// read the expiry date
	if (reggie.GetStringValue(L"val2", szString, MAX_PATH))
	{
		// convert to byte
		pByte = NULL;
		crypt.TranslateRegistryString(szString, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.DecriptStringFromBytes_Aes(pByte, nBytes, szExpDate, MAX_PATH);
//			cng.DecryptCipher(pByte, nBytes, szExpDate, MAX_PATH);
			CoTaskMemFree((LPVOID)pByte);
		}
	}
	// difference count
	if (reggie.GetStringValue(L"val3", szString, MAX_PATH))
	{
		// convert to byte
		pByte = NULL;
		crypt.TranslateRegistryString(szString, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.DecriptStringFromBytes_Aes(pByte, nBytes, szDifference, MAX_PATH);
	//		cng.DecryptCipher(pByte, nBytes, szDifference, MAX_PATH);
			CoTaskMemFree((LPVOID)pByte);
			diffCount = 0;
			if (!_stscanf_s(szDifference, L"%d", &diffCount))
			{
				diffCount = 0;
			}
		}
	}

	fSuccess = this->MyCheckDate(szLastOpDate, MAX_PATH, szExpDate, &diffCount);
	if (fSuccess)
	{
		// encrypt the last operating date
//		if (cng.EnscryptString(szLastOpDate, &pByte, &nBytes))
		crypt.EnscriptStringToBytes_Aes(szLastOpDate, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.FormRegistryString(pByte, nBytes, &szTemp);
			if (NULL != szTemp)
			{
				reggie.SetStringValue(L"val1", szTemp);
				CoTaskMemFree((LPVOID)szTemp);
			}
			CoTaskMemFree((LPVOID)pByte);
		}
		// enscript the difference count
		StringCchPrintf(szDifference, MAX_PATH, L"%d", diffCount);
//		if (cng.EnscryptString(szDifference, &pByte, &nBytes))
		crypt.EnscriptStringToBytes_Aes(szDifference, &pByte, &nBytes);
		if (NULL != pByte)
		{
			crypt.FormRegistryString(pByte, nBytes, &szTemp);
			if (NULL != szTemp)
			{
				reggie.SetStringValue(L"val3", szTemp);
				CoTaskMemFree((LPVOID)szTemp);
			}
			CoTaskMemFree((LPVOID)pByte);
		}
	}
	return fSuccess;
}

BOOL CDateCheck::MyCheckDate(
	LPTSTR			szLastOpDate,
	UINT			nBufferSize,
	LPCTSTR			szExpDate,
	long		*	diffCount)
{
	HRESULT			hr;
	DATE			lastOp;			// last operated date
	DATE			expDate;		// expiry date
	LCID			lcid = 0x1009;			// Canadian English
	SYSTEMTIME		st;
	DATE			currentDate;			// current time
	BSTR			bstr;

	hr = VarDateFromStr(szLastOpDate, lcid, 0, &lastOp);
	if (FAILED(hr)) return FALSE;
	hr = VarDateFromStr(szExpDate, lcid, 0, &expDate);
	if (FAILED(hr)) return FALSE;
	// get the current date
	GetSystemTime(&st);
	SystemTimeToVariantTime(&st, &currentDate);
	// check for fail
	if ((lastOp > (currentDate + 2)) || (currentDate > expDate) || (!this->CheckDiffCount(lastOp, expDate, diffCount)))
	{
		// current date changed too much
		return FALSE;
	}
	else
	{
		// update the last operating date
		hr = VarBstrFromDate(currentDate, lcid, VAR_DATEVALUEONLY, &bstr);
		if (SUCCEEDED(hr))
		{
			StringCchCopy(szLastOpDate, nBufferSize, bstr);
			SysFreeString(bstr);
		}
		// check the expiry date
		return currentDate <= expDate;
	}
}

BOOL CDateCheck::CheckDiffCount(
	DATE			lastOp,			// last operated date
	DATE			expDate,		// expiry date
	long		*	diffCount)
{
	if ((expDate - lastOp) > 1.0)
	{
		// OK case
		return TRUE;
	}
	else
	{
		// need to check the difference count
		if ((*diffCount) >= 30)
		{
			return FALSE;
		}
		else
		{
			(*diffCount)++;
			return TRUE;
		}
	}
}
