#include "stdafx.h"
#include "MyCNG.h"


#define NT_SUCCESS(Status)          (((NTSTATUS)(Status)) >= 0)
#define STATUS_UNSUCCESSFUL         ((NTSTATUS)0xC0000001L)


CMyCNG::CMyCNG()
{
}


CMyCNG::~CMyCNG()
{
}

static const BYTE rgbAES128Key[] =
{
	0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
	0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F,
	0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
	0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F,
	0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07,
	0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F
};

static const BYTE rgbIV[] =
{
	0xee, 0x01, 0x02, 0x03, 0x04, 0x15, 0x06, 0x07,
	0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F
};

BOOL CMyCNG::MyGenerateKey(
	BCRYPT_ALG_HANDLE	hAesAlg,
	BCRYPT_KEY_HANDLE*	phKey)
{
	NTSTATUS	status = STATUS_UNSUCCESSFUL;
	BOOL		fSuccess = FALSE;
	DWORD		cbKeyObject = 0;
	DWORD		cbData = 0;
	PBYTE		pbKeyObject = NULL;

	*phKey = NULL;
	// generate a symmetric key
	// determine size for the key
	status = BCryptGetProperty(hAesAlg, BCRYPT_OBJECT_LENGTH, (PBYTE)&cbKeyObject, sizeof(DWORD), &cbData, 0);
	if ((NT_SUCCESS(status)))
	{
		pbKeyObject = (PBYTE)HeapAlloc(GetProcessHeap(), 0, cbKeyObject);
		status = BCryptGenerateSymmetricKey(hAesAlg, phKey, pbKeyObject, cbKeyObject, (PBYTE)rgbAES128Key, sizeof(rgbAES128Key), 0);
	}
	return NT_SUCCESS(status);
}

// get IV vector
PBYTE CMyCNG::GetIV(
	BCRYPT_ALG_HANDLE	hAesAlg,
	DWORD		*	dwBlockSize)
{
	NTSTATUS	status = STATUS_UNSUCCESSFUL;
	DWORD		cbBlockLen = 0;
	DWORD		cbData = 0;
	PBYTE		piv = NULL;

	// Calculate the block length for the IV.
	status = BCryptGetProperty(hAesAlg, BCRYPT_BLOCK_LENGTH,
			(PBYTE)&cbBlockLen,
			sizeof(DWORD),
			&cbData,
			0);
	if (NT_SUCCESS(status))
	{
		piv = (PBYTE)HeapAlloc(GetProcessHeap(), 0, cbBlockLen);
		*dwBlockSize = cbBlockLen;
		if (NULL != piv)
		{
			memcpy(piv, rgbIV, cbBlockLen);
		}
	}
	return piv;
}

BOOL CMyCNG::DoEncrypt(
	PUCHAR		szInput,
	UINT		nInput,
	PUCHAR	*	szOutput,
	UINT	*	nBufferSize)
{
	NTSTATUS			status = STATUS_UNSUCCESSFUL;
	BCRYPT_ALG_HANDLE	hAesAlg = NULL;
	BOOL				fSuccess = FALSE;
	BCRYPT_KEY_HANDLE	hKey = NULL;
	PBYTE				piv = NULL;
	DWORD				cbBlockLen = 0;
	DWORD				cbCipherText;

	// 
	*szOutput = NULL;
	*nBufferSize = 0;
	status = BCryptOpenAlgorithmProvider(&hAesAlg, BCRYPT_AES_ALGORITHM, NULL, 0);
	if (NT_SUCCESS(status))
	{
		// generate the key
		if (this->MyGenerateKey(hAesAlg, &hKey))
		{
			// initialization
			piv = this->GetIV(hAesAlg, &cbBlockLen);
			if (NULL != piv)
			{
				// determine size for decripted data
				status = BCryptEncrypt(hKey, szInput, nInput, NULL, piv, cbBlockLen,
					NULL,
					0,
					&cbCipherText,
					BCRYPT_BLOCK_PADDING);
				if (NT_SUCCESS(status))
				{
					*szOutput = (PBYTE)HeapAlloc(GetProcessHeap(), 0, cbCipherText);
					*nBufferSize = cbCipherText;
					status = BCryptEncrypt(hKey, szInput, nInput, NULL, piv, cbBlockLen,
						*szOutput, *nBufferSize, &cbCipherText,
						BCRYPT_BLOCK_PADDING);
					fSuccess = NT_SUCCESS(status);
				}
				HeapFree(GetProcessHeap(), 0, piv);
			}
			BCryptDestroyKey(hKey);
		}
		BCryptCloseAlgorithmProvider(hAesAlg, 0);
	}
	return fSuccess;
}

BOOL CMyCNG::DoDecrypt(
	PUCHAR		szCipher,
	UINT		nCipher,
	PUCHAR	*	szText,
	UINT	*	nBufferSize)
{
	NTSTATUS			status = STATUS_UNSUCCESSFUL;
	BCRYPT_ALG_HANDLE	hAesAlg = NULL;
	BOOL				fSuccess = FALSE;
	BCRYPT_KEY_HANDLE	hKey = NULL;
	PBYTE				piv = NULL;
	DWORD				cbBlockLen = 0;
	DWORD				cbPlainText;

	*szText = NULL;
	*nBufferSize = 0;
	status = BCryptOpenAlgorithmProvider(&hAesAlg, BCRYPT_AES_ALGORITHM, NULL, 0);
	if (NT_SUCCESS(status))
	{
		// generate the key
		if (this->MyGenerateKey(hAesAlg, &hKey))
		{
			// initialization
			piv = this->GetIV(hAesAlg, &cbBlockLen);
			if (NULL != piv)
			{
				// determine the size for the plain text
				status = BCryptDecrypt(hKey, szCipher, nCipher, NULL, piv, cbBlockLen,
					NULL, 0, &cbPlainText, BCRYPT_BLOCK_PADDING);
				if (NT_SUCCESS(status))
				{
					*szText = (PBYTE)HeapAlloc(GetProcessHeap(), 0, cbPlainText);
					*nBufferSize = cbPlainText;
					status = BCryptDecrypt(hKey, szCipher, nCipher, NULL, piv, cbBlockLen,
						*szText, *nBufferSize, &cbPlainText, BCRYPT_BLOCK_PADDING);
					fSuccess = NT_SUCCESS(status);
				}
				HeapFree(GetProcessHeap(), 0, piv);
			}
			BCryptDestroyKey(hKey);
		}
		BCryptCloseAlgorithmProvider(hAesAlg, 0);
	}
	return fSuccess;
}

BOOL CMyCNG::EnscryptString(
	LPCTSTR		szInput,
	BYTE	**	ppEnscript,
	ULONG	*	nArray)
{
	BOOL		fSuccess = FALSE;
	size_t		i;
	size_t		slen;
	PUCHAR		Input = NULL;
	DWORD		nInput;
	PUCHAR		Cipher = NULL;
	UINT		nCipher;

	*ppEnscript = NULL;
	*nArray = 0;
	// form input string
	StringCchLength(szInput, MAX_PATH, &slen);
	nInput = slen;
	Input = (PUCHAR)CoTaskMemAlloc(slen * sizeof(UCHAR));
	for (i = 0; i < slen; i++)
	{
		Input[i] = (UCHAR)szInput[i];
	}
	if (this->DoEncrypt(Input, slen, &Cipher, &nCipher))
	{
		fSuccess = TRUE;
		*nArray = nCipher;
		*ppEnscript = (BYTE*)CoTaskMemAlloc(*nArray);
		for (i = 0; i < (*nArray); i++)
			(*ppEnscript)[i] = (BYTE)Cipher[i];
		HeapFree(GetProcessHeap(), 0, Cipher);
	}
	CoTaskMemFree((LPVOID)Input);
	return fSuccess;
}

BOOL CMyCNG::DecryptCipher(
	BYTE	*	Cipher,
	UINT		nCipher,
	LPTSTR		szString,
	UINT		nBufferSize)
{
	BOOL		fSuccess = FALSE;
	PUCHAR		szCipher = NULL;
	PUCHAR		szText = NULL;
	UINT		nText;
	UINT		i;

	szString[0] = '\0';
	szCipher = (PUCHAR)CoTaskMemAlloc(nCipher * sizeof(UCHAR));
	for (i = 0; i < nCipher; i++)
	{
		szCipher[i] = (UCHAR)Cipher[i];
	}
	if (this->DoDecrypt(szCipher, nCipher, &szText, &nText))
	{
		for (i = 0; i < nText && i<nBufferSize; i++)
		{
			szString[i] = (TCHAR)szText[i];
		}
		szString[i] = '\0';		// null terminate
		fSuccess = TRUE;
		HeapFree(GetProcessHeap(), 0, szText);
	}
	CoTaskMemFree((LPVOID)szCipher);
	return fSuccess;
}
