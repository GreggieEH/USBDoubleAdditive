#pragma once
class MyCrypto
{
public:
	MyCrypto();
	~MyCrypto();
	BOOL			DoInit();
	BOOL			RunCrypto();
	void			doRun();
	void			GetKey(
						BYTE	**	ppKey,
						ULONG	*	nArray);
	void			SetKey(
						BYTE	*	pKey,
						ULONG		nArray);
	void			GetIV(
						BYTE	**	ppIV,
						ULONG	*	nArray);
	void			SetIV(
						BYTE	*	pIV,
						ULONG		nArray);
	void			EnscriptStringToBytes_Aes(
						LPCTSTR		szString,
						BYTE	**	ppEnscript,
						ULONG	*	nArray);
	void			DecriptStringFromBytes_Aes(
						BYTE	*	pEnscript,
						ULONG		nBytes,
						LPTSTR		szString,
						UINT		nBufferSize);
	void			FormRegistryString(
						BYTE	*	pEnscript,
						ULONG		nBytes,
						LPTSTR	*	szString);
	void			TranslateRegistryString(
						LPCTSTR			szString,
						BYTE		**	ppEnscript,
						ULONG		*	nBytes);
protected:
	void			SaveToFile(
						LPCTSTR		szFileName,
						BYTE	*	arr,
						ULONG		nArray);
	BOOL			LoadResourceData(
						UINT		resourceID,
						BYTE	**	ppArray,
						ULONG	*	nArray);
	// convert char to hex
	BYTE			CharToHex(
						TCHAR		szValue);
private:
	IDispatch	*	m_pdisp;
};

