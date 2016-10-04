#include "stdafx.h"
#include "MyCrypto.h"
#include "Resource.h"
#include "Server.h"

/*
key[32]
56
84






iv[16]








*/

MyCrypto::MyCrypto() :
	m_pdisp(NULL)
{
	HRESULT			hr;
	CLSID			clsid;

	hr = CLSIDFromProgID(L"MyStringHandler.clsMyStringHandler", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&this->m_pdisp);
}


MyCrypto::~MyCrypto()
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

BOOL MyCrypto::DoInit()
{
	BOOL			fSuccess = FALSE;
	BYTE		*	key = NULL;
	ULONG			nKey;
	BYTE		*	iv = NULL;
	ULONG			niv;

	if (this->LoadResourceData(keyData, &key, &nKey))
	{
		//		this->SaveToFile(L"key.txt", key, nKey);
		this->SetKey(key, nKey);
		//		this->GetIV(&iv, &niv);
		//		if (NULL != iv)
		if (this->LoadResourceData(ivData, &iv, &niv))
		{
			this->SetIV(iv, niv);
			fSuccess = TRUE;
			CoTaskMemFree((LPVOID)iv);
		}
		CoTaskMemFree((LPVOID)key);
	}
	return fSuccess;
}

BOOL MyCrypto::RunCrypto()
{
	BYTE		*	key = NULL;
	ULONG			nKey;
	BYTE		*	iv = NULL;
	ULONG			niv;
	BYTE		*	enscript = NULL;
	ULONG			nEnscript;
	TCHAR			szTestString[MAX_PATH];
	TCHAR			szResult[MAX_PATH];
	BOOL			fSuccess = FALSE;

	StringCchCopy(szTestString, MAX_PATH, L"This is a stupid stupid stupid test string");
//	this->doRun();
//	this->GetKey(&key, &nKey);
	if (this->LoadResourceData(keyData, &key, &nKey))
	{
//		this->SaveToFile(L"key.txt", key, nKey);
		this->SetKey(key, nKey);
//		this->GetIV(&iv, &niv);
//		if (NULL != iv)
		if (this->LoadResourceData(ivData, &iv, &niv))
		{
//			this->SaveToFile(L"iv.txt", iv, niv);
			this->SetIV(iv, niv);
			this->EnscriptStringToBytes_Aes(szTestString, &enscript, &nEnscript);
			if (NULL != enscript)
			{
				this->DecriptStringFromBytes_Aes(enscript, nEnscript, szResult, MAX_PATH);
				fSuccess = 0 == lstrcmpi(szResult, szTestString);
				CoTaskMemFree((LPVOID)enscript);
			}
			CoTaskMemFree((LPVOID)iv);
		}
		CoTaskMemFree((LPVOID)key);
	}
	return fSuccess;
}

void MyCrypto::doRun()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"doRun", &dispid);
	Utils_InvokeMethod(this->m_pdisp, dispid, NULL, 0, NULL);
}

void MyCrypto::GetKey(
	BYTE	**	ppKey,
	ULONG	*	nArray)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	long			ub;
	BYTE		*	arr;
	ULONG			i;
	*ppKey = NULL;
	*nArray = 0;
	Utils_GetMemid(this->m_pdisp, L"Key", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_UI1) == varResult.vt && NULL != varResult.parray)
		{
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nArray = ub + 1;
			*ppKey = (BYTE*)CoTaskMemAlloc((*nArray) * sizeof(BYTE));
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nArray); i++)
				(*ppKey)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
}

void MyCrypto::SetKey(
	BYTE	*	pKey,
	ULONG		nArray)
{
	DISPID			dispid;
	VARIANTARG		varg;
	SAFEARRAYBOUND	sab;
	BYTE		*	arr;
	ULONG			i;
	Utils_GetMemid(this->m_pdisp, L"Key", &dispid);
	sab.lLbound = 0;
	sab.cElements = nArray;
	VariantInit(&varg);
	varg.vt = VT_ARRAY | VT_UI1;
	varg.parray = SafeArrayCreate(VT_UI1, 1, &sab);
	SafeArrayAccessData(varg.parray, (void**)&arr);
	for (i = 0; i < nArray; i++)
		arr[i] = pKey[i];
	SafeArrayUnaccessData(varg.parray);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

void MyCrypto::GetIV(
	BYTE	**	ppIV,
	ULONG	*	nArray)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	long			ub;
	BYTE		*	arr;
	ULONG			i;
	*ppIV = NULL;
	*nArray = 0;
	Utils_GetMemid(this->m_pdisp, L"IV", &dispid);
	hr = Utils_InvokePropertyGet(this->m_pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_UI1) == varResult.vt && NULL != varResult.parray)
		{
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nArray = ub + 1;
			*ppIV = (BYTE*)CoTaskMemAlloc((*nArray) * sizeof(BYTE));
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nArray); i++)
				(*ppIV)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
}

void MyCrypto::SetIV(
	BYTE	*	pIV,
	ULONG		nArray)
{
	DISPID			dispid;
	VARIANTARG		varg;
	SAFEARRAYBOUND	sab;
	BYTE		*	arr;
	ULONG			i;
	Utils_GetMemid(this->m_pdisp, L"IV", &dispid);
	sab.lLbound = 0;
	sab.cElements = nArray;
	VariantInit(&varg);
	varg.vt = VT_ARRAY | VT_UI1;
	varg.parray = SafeArrayCreate(VT_UI1, 1, &sab);
	SafeArrayAccessData(varg.parray, (void**)&arr);
	for (i = 0; i < nArray; i++)
		arr[i] = pIV[i];
	SafeArrayUnaccessData(varg.parray);
	Utils_InvokePropertyPut(this->m_pdisp, dispid, &varg, 1);
	VariantClear(&varg);
}

void MyCrypto::EnscriptStringToBytes_Aes(
	LPCTSTR		szString,
	BYTE	**	ppEnscript,
	ULONG	*	nArray)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			ub;
	BYTE		*	arr;
	ULONG			i;
	*ppEnscript = NULL;
	*nArray = 0;
	Utils_GetMemid(this->m_pdisp, L"EncryptStringToBytes_Aes", &dispid);
	InitVariantFromString(szString, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		if ((VT_ARRAY | VT_UI1) == varResult.vt && NULL != varResult.parray)
		{
			SafeArrayGetUBound(varResult.parray, 1, &ub);
			*nArray = ub + 1;
			*ppEnscript = (BYTE*)CoTaskMemAlloc((*nArray) * sizeof(BYTE));
			SafeArrayAccessData(varResult.parray, (void**)&arr);
			for (i = 0; i < (*nArray); i++)
				(*ppEnscript)[i] = arr[i];
			SafeArrayUnaccessData(varResult.parray);
		}
		VariantClear(&varResult);
	}
}

void MyCrypto::DecriptStringFromBytes_Aes(
	BYTE	*	pEnscript,
	ULONG		nBytes,
	LPTSTR		szString,
	UINT		nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	SAFEARRAYBOUND	sab;
	BYTE		*	arr;
	ULONG			i;
	szString[0] = '\0';
	Utils_GetMemid(this->m_pdisp, L"DecryptStringFromBytes_Aes", &dispid);
	VariantInit(&varg);
	varg.vt = VT_ARRAY | VT_UI1;
	sab.lLbound = 0;
	sab.cElements = nBytes;
	varg.parray = SafeArrayCreate(VT_UI1, 1, &sab);
	SafeArrayAccessData(varg.parray, (void**)&arr);
	for (i = 0; i < nBytes; i++)
		arr[i] = pEnscript[i];
	SafeArrayUnaccessData(varg.parray);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szString, nBufferSize);
		VariantClear(&varResult);
	}
}

void MyCrypto::SaveToFile(
	LPCTSTR		szFileName,
	BYTE	*	arr,
	ULONG		nArray)
{
	TCHAR			szFilePath[MAX_PATH];
	HANDLE			hFile;
	ULONG			i;
	TCHAR			szOneLine[MAX_PATH];
	size_t			nWrite;
	DWORD			nWritten;

	// form file name
	StringCchPrintf(szFilePath, MAX_PATH, 
		TEXT("C:\\Users\\Greg\\Documents\\Visual Studio 2015\\Projects\\MonoHandler\\Debug\\%s"), szFileName);
	hFile = CreateFile(szFilePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		for (i = 0; i < nArray; i++)
		{
			StringCchPrintf(szOneLine, MAX_PATH, L"%1d\r\n", arr[i]);
			StringCbLength(szOneLine, MAX_PATH, &nWrite);
			WriteFile(hFile, (LPCVOID)szOneLine, nWrite, &nWritten, NULL);
		}
		CloseHandle(hFile);
	}
}

BOOL MyCrypto::LoadResourceData(
	UINT		resourceID,
	BYTE	**	ppArray,
	ULONG	*	nArray)
{
	HMODULE		hModule = GetServer()->GetInstance();
	HRSRC		hRscr = NULL;
	HGLOBAL		hGlobal;
	WORD	*	arr;
	DWORD		resourceSize;
	BOOL		fSuccess = FALSE;
	UINT		i;
	*ppArray = NULL;
	*nArray = 0;
	hRscr = FindResource(hModule, MAKEINTRESOURCE(resourceID), RT_RCDATA);
	if (NULL != hRscr)
	{
		hGlobal = LoadResource(hModule, hRscr);
		if (NULL != hGlobal)
		{
			// resource size
			resourceSize = SizeofResource(hModule, hRscr);
			// array data
			arr = (WORD*)LockResource(hGlobal);
			// allocate output - convert to bytes
			*nArray = resourceSize / sizeof(WORD);			
			*ppArray = (BYTE*)CoTaskMemAlloc(*nArray);
			fSuccess = TRUE;
			// copy array
			for (i = 0; i < (*nArray); i++)
				(*ppArray)[i] = (BYTE)arr[i];
		}
	}
	return fSuccess;
}

void MyCrypto::FormRegistryString(
	BYTE	*	pEnscript,
	ULONG		nBytes,
	LPTSTR	*	szString)
{
	ULONG			nChars = (nBytes * 2) + 1;
	ULONG			i;
	TCHAR			szTemp[3];
	size_t			slen;

	*szString = (LPTSTR)CoTaskMemAlloc(nChars * sizeof(TCHAR));
	ZeroMemory((PVOID)*szString, nChars * sizeof(TCHAR));
	for (i = 0; i < nBytes; i++)
	{
		if (pEnscript[i] > 15)
		{
			StringCchPrintf(szTemp, 3, L"%2x", pEnscript[i]);
		}
		else
		{
			StringCchPrintf(szTemp, 3, L"0%1x", pEnscript[i]);
		}
		StringCchLength(*szString, nChars, &slen);
		if (slen > 0)
		{
			StringCchCat(*szString, nChars, szTemp);
		}
		else
		{
			StringCchCopy(*szString, nChars, szTemp);
		}
	}
}

void MyCrypto::TranslateRegistryString(
	LPCTSTR			szString,
	BYTE		**	ppEnscript,
	ULONG		*	nBytes)
{
	size_t			slen;
	size_t			i;
	TCHAR			szTemp[MAX_PATH];
	short int		bval;
	BYTE		*	arr;
	TCHAR			szCopy[MAX_PATH];

	*ppEnscript = NULL;
	*nBytes = 0;
	StringCchCopy(szCopy, MAX_PATH, szString);
	StringCchLength(szCopy, MAX_PATH, &slen);
	*nBytes = slen / 2;
	*ppEnscript = (BYTE*)CoTaskMemAlloc(*nBytes);
//	arr = (BYTE*)CoTaskMemAlloc(*nBytes);
	for (i = 0; i < (*nBytes); i++)
	{
		StringCchCopyN(szTemp, MAX_PATH, &(szCopy[i * 2]), 2);
		bval = (this->CharToHex(szTemp[0]) * 16) + this->CharToHex(szTemp[1]);
//		bval = 0;
//		_stscanf_s(szTemp, L"%x", &bval);
		(*(ppEnscript))[i] = (BYTE) bval;
//		arr[i] = (BYTE)bval;
	}
	BOOL			stat = FALSE;
}

// convert char to hex
BYTE MyCrypto::CharToHex(
	TCHAR		szValue)
{
	if (szValue >= '0' && szValue <= '9')
	{
		return (BYTE)(szValue - '0');
	}
	else if (szValue >= 'a' && szValue <= 'f')
	{
		return (BYTE)((szValue - 'a') + 10);
	}
	else if (szValue >= 'A' && szValue <= 'F')
	{
		return (BYTE)((szValue - 'A') + 10);
	}
	else
	{
		return (BYTE)0;
	}
}
