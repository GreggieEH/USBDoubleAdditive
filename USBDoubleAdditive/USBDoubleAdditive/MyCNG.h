#pragma once
class CMyCNG
{
public:
	CMyCNG();
	~CMyCNG();
	BOOL			EnscryptString(
						LPCTSTR		szInput,
						BYTE	**	ppEnscript,
						ULONG	*	nArray);
	BOOL			DoEncrypt(
						PUCHAR		szInput,
						UINT		nInput,
						PUCHAR	*	szOutput,
						UINT	*	nBufferSize);
	BOOL			DecryptCipher(
						BYTE	*	Cipher,
						UINT		nCipher,
						LPTSTR		szString,
						UINT		nBufferSize);
	BOOL			DoDecrypt(
						PUCHAR		szCipher,
						UINT		nCipher,
						PUCHAR	*	szText,
						UINT	*	nBufferSize);
protected:
	// generate symetric key
	BOOL			MyGenerateKey(
						BCRYPT_ALG_HANDLE	hAesAlg,
						BCRYPT_KEY_HANDLE*	phKey);
	// get IV vector
	PBYTE			GetIV(
						BCRYPT_ALG_HANDLE	hAesAlg,
						DWORD		*	dwBlockSize);
};

