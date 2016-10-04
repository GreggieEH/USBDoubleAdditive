#include "StdAfx.h"
#include "MyFactory.h"
#include "Server.h"
#include "MyObject.h"
#include "MyPropPage.h"
#include "MyGuids.h"

CMyFactory::CMyFactory(BOOL regPropPage /* = FALSE*/) :
	m_cRefs(0),
	m_regPropPage(regPropPage)
{
}

CMyFactory::~CMyFactory(void)
{
}

// IUnknown methods
STDMETHODIMP CMyFactory::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IClassFactory == riid || IID_IClassFactory2 == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyFactory::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyFactory::Release()
{
	ULONG				cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		this->m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// IClassFactory methods
STDMETHODIMP CMyFactory::CreateInstance(
									IUnknown	*	pUnkOuter,
									REFIID			riid,
									void		**	ppv)
{
	// check for licensing
	if (!GetServer()->CheckLicence()) return CLASS_E_NOTLICENSED;
	return this->doCreateInstance(pUnkOuter, riid, ppv);
}

STDMETHODIMP CMyFactory::LockServer(
									BOOL			fLock)
{
	if (fLock)
		GetServer()->LocksUp();
	else
		GetServer()->LocksDown();
	return S_OK;
}

//  IClassFactory2 methods
STDMETHODIMP CMyFactory::CreateInstanceLic(
	IUnknown	*	pUnkOuter,
	IUnknown	*	pUnkReserved,
	REFIID			riid,
	BSTR			bstrKey,
	PVOID		*	ppvObj)
{
	// check input license
	WCHAR			wszString[MAX_PATH];
	StringFromGUID2(MY_LIC_KEY, wszString, MAX_PATH);
	if (0 == lstrcmpi(bstrKey, wszString))
	{
		return this->doCreateInstance(pUnkOuter, riid, ppvObj);
	}
	else
	{
		return CLASS_E_NOTLICENSED;
	}
}

STDMETHODIMP CMyFactory::GetLicInfo(
	LICINFO		*	pLicInfo)
{
	HRESULT			hr = S_OK;
	BOOL			bLicensed = GetServer()->CheckLicence();
	if (NULL != pLicInfo)
	{
		pLicInfo->cbLicInfo = sizeof(LICINFO);

		// Inform whether RequestLicKey will work.
		pLicInfo->fRuntimeKeyAvail =bLicensed;

		// Inform whether the standard CreateInstance will work.
		pLicInfo->fLicVerified = bLicensed;
	}
	else
		hr = E_POINTER;
	return hr;
}

STDMETHODIMP CMyFactory::RequestLicKey(
	DWORD			dwReserved,
	BSTR		*	pBstrKey)
{
	return E_NOTIMPL;
}

HRESULT CMyFactory::doCreateInstance(
	IUnknown	*	pUnkOuter,
	REFIID			riid,
	void		**	ppvObject)
{
	// check the input parameters
	HRESULT				hr;
	CMyObject		*	pCob;		// object
	CMyPropPage		*	pPropPage;	// property page

	// check if property page
	if (this->m_regPropPage)
	{
		if (NULL == pUnkOuter)
		{
			pPropPage = new CMyPropPage();
			if (NULL != pPropPage)
			{
				hr = pPropPage->QueryInterface(riid, ppvObject);
				if (FAILED(hr))
					delete pPropPage;
				return hr;
			}
			else return E_OUTOFMEMORY;
		}
		else return CLASS_E_NOAGGREGATION;
	}
	// not property page reg main object
	if (NULL != pUnkOuter && riid != IID_IUnknown)
		hr = CLASS_E_NOAGGREGATION;
	else
	{
		// let's create the object
		pCob = new CMyObject(pUnkOuter);
		if (pCob != NULL)
		{
			// increment the object count
			GetServer()->ObjectsUp();
			// initialize the object
			hr = pCob->Init();
			if (SUCCEEDED(hr))
			{
				hr = pCob->QueryInterface(riid, ppvObject);
			}
			if (FAILED(hr))
			{
				GetServer()->ObjectsDown();
				delete pCob;
			}
		}
		else
			hr = E_OUTOFMEMORY;
	}
	return hr;
}
