#pragma once

class CMyFactory : public IClassFactory2
{
public:
	CMyFactory(BOOL regPropPage = FALSE);
	~CMyFactory(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// IClassFactory methods
	STDMETHODIMP				CreateInstance(
									IUnknown	*	pUnkOuter,
									REFIID			riid,
									void		**	ppvObject);
	STDMETHODIMP				LockServer(
									BOOL			fLock);
	//  IClassFactory2 methods
	STDMETHODIMP				CreateInstanceLic(
									IUnknown	*	pUnkOuter,
									IUnknown	*	pUnkReserved,
									REFIID			riid,
									BSTR			bstrKey,
									PVOID		*	ppvObj);
	STDMETHODIMP				GetLicInfo(
									LICINFO		*	pLicInfo);
	STDMETHODIMP				RequestLicKey(
									DWORD			dwReserved,
									BSTR		*	pBstrKey);
protected:
	HRESULT						doCreateInstance(
									IUnknown	*	pUnkOuter,
									REFIID			riid,
									void		**	ppvObject);
private:
	ULONG						m_cRefs;
	BOOL						m_regPropPage;
};
