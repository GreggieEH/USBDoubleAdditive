#include "StdAfx.h"
#include "MyObject.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"
#include "MyUSBDoubleAdditive.h"
#include "DlgDisplayConfigValues.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIProvideClassInfo2(NULL),
	m_pImp_clsIMono(NULL),
	m_pImpIPersistStreamInit(NULL),
	m_pImpISpecifyPropertyPages(NULL),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// object reference count
	m_cRefs(0),
	m_pMyUSBDoubleAdditive(NULL),
	// _clsIMono interface id
	m_iid_clsIMono(IID_NULL),
	// __clsIMono events
	m_iid__clsIMono(IID_NULL),
	m_dispidGratingChanged(DISPID_UNKNOWN),
	m_dispidBeforeMoveChange(DISPID_UNKNOWN),
	m_dispidHaveAborted(DISPID_UNKNOWN),
	m_dispidMoveCompleted(DISPID_UNKNOWN),
	m_dispidMoveError(DISPID_UNKNOWN),
	m_dispidRequestChangeGrating(DISPID_UNKNOWN),
	m_dispidRequestParentWindow(DISPID_UNKNOWN),
	m_dispidStatusMessage(DISPID_UNKNOWN)

{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
GetServer()->DoLogString(TEXT("Close UP"));

	Utils_DELETE_POINTER(this->m_pImpIDispatch);
	Utils_DELETE_POINTER(this->m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(this->m_pImpIProvideClassInfo2);
	Utils_DELETE_POINTER(this->m_pImp_clsIMono);
	Utils_DELETE_POINTER(this->m_pImpIPersistStreamInit);
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_DELETE_POINTER(this->m_pMyUSBDoubleAdditive);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || MY_IID == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	else if (riid == m_iid_clsIMono)
		*ppv = this->m_pImp_clsIMono;
	else if (IID_IPersist == riid || IID_IPersistStreamInit == riid)
		*ppv = this->m_pImpIPersistStreamInit;
	else if (IID_ISpecifyPropertyPages == riid)
		*ppv = this->m_pImpISpecifyPropertyPages;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;

TCHAR			szMessage[MAX_PATH];
StringCchPrintf(szMessage, MAX_PATH, TEXT("Release count = %1d"), m_cRefs);
GetServer()->DoLogString(szMessage);

	if (0 == cRefs)
	{
		m_cRefs++;
		GetServer()->ObjectsDown();
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT						hr;

	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pImp_clsIMono					= new CImp_clsIMono(this, this->m_pUnkOuter);
	this->m_pImpIPersistStreamInit			= new CImpIPersistStreamInit(this, this->m_pUnkOuter);
	this->m_pImpISpecifyPropertyPages		= new CImpISpecifyPropertyPages(this, this->m_pUnkOuter);
	this->m_pMyUSBDoubleAdditive			= new CMyUSBDoubleAdditive(this);
	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pImp_clsIMono					&&
		NULL != this->m_pImpIPersistStreamInit			&&
		NULL != this->m_pImpISpecifyPropertyPages		&&
		NULL != this->m_pMyUSBDoubleAdditive)
	{
		hr = Utils_CreateConnectionPoint(this, MY_IIDSINK,
			 &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
		if (SUCCEEDED(hr))
		{
			hr = Utils_CreateConnectionPoint(this, IID_IPropertyNotifySink,
				&(this->m_paConnPts[CONN_PT_PROPNOTIFY]));
		}
		if (SUCCEEDED(hr))
		{
			hr = this->Init__clsIMono();
			if (SUCCEEDED(hr))
			{
				hr = Utils_CreateConnectionPoint(this, this->m_iid__clsIMono,
					&(this->m_paConnPts[CONN_PT__clsIMono]));
			}
		}
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}


HRESULT CMyObject::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

HRESULT CMyObject::GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo;
	BOOL			fSuccess	= FALSE;
	*ppTypeInfo	= NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

HRESULT CMyObject::Init__clsIMono()
{
	HRESULT				hr;
//	ITypeLib		*	pTypeLib;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->GetRefTypeInfo(TEXT("__clsIMono") , &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_iid__clsIMono	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, TEXT("GratingChanged"), &m_dispidGratingChanged);
		Utils_GetMemid(pTypeInfo, TEXT("BeforeMoveChange"), &m_dispidBeforeMoveChange);
		Utils_GetMemid(pTypeInfo, TEXT("HaveAborted"), &m_dispidHaveAborted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveCompleted"), &m_dispidMoveCompleted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveError"), &m_dispidMoveError);
		Utils_GetMemid(pTypeInfo, TEXT("RequestChangeGrating"), &m_dispidRequestChangeGrating);
		Utils_GetMemid(pTypeInfo, TEXT("RequestParentWindow"), &m_dispidRequestParentWindow);
		Utils_GetMemid(pTypeInfo, TEXT("StatusMessage"), &m_dispidStatusMessage);
		pTypeInfo->Release();
	}
	return hr;
}

// sink events
void CMyObject::FireMonoPositionChanged(
									LPCTSTR			Mono,
									double			position)
{
	VARIANTARG				avarg[2];
	InitVariantFromString(Mono, &avarg[1]);
	InitVariantFromDouble(position, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_MonoPositionChanged, avarg, 2);
	VariantClear(&avarg[1]);
}

void CMyObject::FireMonoGratingChanged(
									LPCTSTR			Mono,
									long			grating)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(Mono, &avarg[1]);
	InitVariantFromInt32(grating, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_MonoGratingChanged, avarg, 2);
	VariantClear(&avarg[1]);
}

void CMyObject::FireMonoAutoGratingChanged(
									LPCTSTR			Mono,
									BOOL			AutoGrating)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(Mono, &avarg[1]);
	InitVariantFromBoolean(AutoGrating, &avarg[0]);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_MonoAutoGratingChanged, avarg, 2);
	VariantClear(&avarg[1]);
}

// __clsIMono sink events
void CMyObject::FireGratingChanged(
									long			grating)
{
	VARIANTARG			varg;
	InitVariantFromInt32(grating, &varg);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidGratingChanged, &varg, 1);
}

BOOL CMyObject::FireBeforeMoveChange(
									double			newPosition)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	InitVariantFromDouble(newPosition, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fAllow;
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidBeforeMoveChange, avarg, 2);
	return VARIANT_TRUE == fAllow;
}

void CMyObject::FireHaveAborted()
{
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidHaveAborted, NULL, 0);
}

void CMyObject::FireMoveCompleted(
									double			newPosition)
{
	VARIANTARG				varg;
	InitVariantFromDouble(newPosition, &varg);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidMoveCompleted, &varg, 1);
}

void CMyObject::FireMoveError(	
									LPCTSTR			err)
{
	VARIANTARG				varg;
	InitVariantFromString(err, &varg);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidMoveError, &varg, 1);
	VariantClear(&varg);
}

BOOL CMyObject::FireRequestChangeGrating(
									long			newGrating)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	InitVariantFromInt32(newGrating, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fAllow;
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidRequestChangeGrating,
		avarg, 2);
	return VARIANT_TRUE == fAllow;
}

HWND CMyObject::FireRequestParentWindow()
{
	VARIANTARG			varg;
	long				lval		= 0;
	VariantInit(&varg);
	varg.vt		= VT_BYREF | VT_I4;
	varg.plVal	= &lval;
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidRequestParentWindow, &varg, 1);
	return (HWND) lval;
}

void CMyObject::FireStatusMessage(
									LPCTSTR			msg, 
									BOOL			AmBusy)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(msg, &avarg[1]);
	InitVariantFromBoolean(AmBusy, &avarg[0]);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidStatusMessage, avarg, 2);
	VariantClear(&avarg[1]);
}

// nested classes
CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	if (0 != iTInfo) return DISP_E_BADINDEX;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetServer()->GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(MY_IID, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	switch (dispIdMember)
	{
	case DISPID_MonoDecoupled:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMonoDecoupled(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetMonoDecoupled(pDispParams);
		break;
	case DISPID_ActiveMono:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetActiveMono(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetActiveMono(pDispParams);
		break;
	case DISPID_SystemInit:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSystemInit(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSystemInit(pDispParams);
		break;
	case DISPID_Grating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetGrating(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetGrating(pDispParams);
		break;
	case DISPID_DisplayMonoSetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayMonoSetup(pDispParams);
		break;
	case DISPID_SetMonoPosition:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetMonoPosition(pDispParams, pVarResult);
		break;
	case DISPID_GetMonoPosition:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetMonoPosition(pDispParams, pVarResult);
		break;
	case DISPID_DisplayMonoConfigValues:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayMonoConfigValues(pDispParams);
		break;
	case DISPID_GetMonoGrating:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetMonoGrating(pDispParams, pVarResult);
		break;
	case DISPID_SetMonoGrating:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetMonoGrating(pDispParams);
		break;
	case DISPID_GetMonoAutoGrating:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->GetMonoAutoGrating(pDispParams, pVarResult);
		break;
	case DISPID_SetMonoAutoGrating:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetMonoAutoGrating(pDispParams);
		break;
	case DISPID_ReInitOnScanStart:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetReInitOnScanStart(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetReInitOnScanStart(pDispParams);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpIDispatch::GetMonoDecoupled(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetMonoDecoupled(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoDecoupled(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetMonoDecoupled(VARIANT_TRUE == varg.boolVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_MonoDecoupled);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetActiveMono(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	TCHAR				szMono[MAX_PATH];
	if (this->m_pBackObj->m_pMyUSBDoubleAdditive->GetActiveMono(szMono, MAX_PATH))
	{
		InitVariantFromString(szMono, pVarResult);
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetActiveMono(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
//	if (NULL != szString)
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetActiveMono(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_ActiveMono);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetSystemInit(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetSystemInit(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetSystemInit(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetSystemInit(VARIANT_TRUE == varg.boolVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_SystemInit);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetGrating(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetCurrentGrating(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetCurrentGrating(varg.lVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_Grating);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::DisplayMonoSetup(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szMono		= NULL;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szMono);
//	VariantClear(&varg);
	this->m_pBackObj->m_pMyUSBDoubleAdditive->DisplayMonoSetup(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szMono);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	double				position;
	BOOL				fSuccess		= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	position	= varg.dblVal;
	fSuccess	= this->m_pBackObj->m_pMyUSBDoubleAdditive->SetMonoPosition(
					szMono, position);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	VariantInit(&varg);
	if (NULL == pVarResult) return E_INVALIDARG;
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
//	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varg);
	InitVariantFromDouble(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetMonoPosition(szMono), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::DisplayMonoConfigValues(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	this->m_pBackObj->m_pMyUSBDoubleAdditive->DisplayMonoConfigValues(szMono);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoGrating(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	InitVariantFromInt32(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetMonoGrating(szMono),  pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetMonoGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	long				grating;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.lVal;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetMonoGrating(szMono, grating);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetMonoAutoGrating(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
//	TCHAR				szMono[MAX_PATH];
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
//	StringCchCopy(szMono, MAX_PATH, szString);
//	CoTaskMemFree((LPVOID) szString);
	InitVariantFromBoolean(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetMonoAutoGrating(varg.bstrVal), pVarResult);
	VariantClear(&varg);
	return S_OK;

}

HRESULT CMyObject::CImpIDispatch::SetMonoAutoGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szString		= NULL;
	TCHAR				szMono[MAX_PATH];
	BOOL				AutoGrating;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szString);
//	VariantClear(&varg);
	StringCchCopy(szMono, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szString);
	hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	AutoGrating = VARIANT_TRUE == varg.boolVal;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetMonoAutoGrating(szMono, AutoGrating);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetReInitOnScanStart(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pBackObj->m_pMyUSBDoubleAdditive->GetReInitOnScanStart(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetReInitOnScanStart(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pBackObj->m_pMyUSBDoubleAdditive->SetReInitOnScanStart(VARIANT_TRUE == varg.boolVal);
	Utils_OnPropChanged(this->m_pBackObj, DISPID_ReInitOnScanStart);
	return S_OK;
}

CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI) 
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetServer()->GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(MY_CLSID, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID)
{
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID	= MY_IIDSINK;
		return S_OK;
	}
	else
	{
		*pGUID	= GUID_NULL;
		return E_INVALIDARG;	
	}
}

CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP)
{
	HRESULT					hr		= CONNECT_E_NOCONNECTION;
	IConnectionPoint	*	pConnPt	= NULL;
	*ppCP	= NULL;
	if (MY_IIDSINK == riid)
		pConnPt	= this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (IID_IPropertyNotifySink == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_PROPNOTIFY];
	else if (riid == this->m_pBackObj->m_iid__clsIMono)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT__clsIMono];
	if (NULL != pConnPt)
	{
		*ppCP = pConnPt;
		pConnPt->AddRef();
		hr = S_OK;
	}
	return hr;
}


CMyObject::CImp_clsIMono::CImp_clsIMono(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pMyObject(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_dispidCurrentWavelength(DISPID_UNKNOWN),
	m_dispidAutoGrating(DISPID_UNKNOWN),
	m_dispidCurrentGrating(DISPID_UNKNOWN),
	m_dispidAmBusy(DISPID_UNKNOWN),
	m_dispidAmInitialized(DISPID_UNKNOWN),
	m_dispidGratingParams(DISPID_UNKNOWN),
	m_dispidMonoParams(DISPID_UNKNOWN),
	m_dispidIsValidPosition(DISPID_UNKNOWN),
	m_dispidConvertStepsToNM(DISPID_UNKNOWN),
	m_dispidDisplayConfigValues(DISPID_UNKNOWN),
	m_dispidDoSetup(DISPID_UNKNOWN),
	m_dispidReadConfig(DISPID_UNKNOWN),
	m_dispidWriteConfig(DISPID_UNKNOWN),
	m_dispidGetGratingDispersion(DISPID_UNKNOWN),
	m_dispidScanStart(DISPID_UNKNOWN)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pMyObject->GetRefTypeInfo(TEXT("_clsIMono"), &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// store the type info
		this->m_pTypeInfo	= pTypeInfo;
		this->m_pTypeInfo->AddRef();
		// get the interface ID
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pMyObject->m_iid_clsIMono	= pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(this->m_pTypeInfo, TEXT("CurrentWavelength"), &m_dispidCurrentWavelength);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AutoGrating"), &m_dispidAutoGrating);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("CurrentGrating"), &m_dispidCurrentGrating);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AmBusy"), &m_dispidAmBusy);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AmInitialized"), &m_dispidAmInitialized);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GratingParams"), &m_dispidGratingParams);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("MonoParams"), &m_dispidMonoParams);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("IsValidPosition"), &m_dispidIsValidPosition);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ConvertStepsToNM"), &m_dispidConvertStepsToNM);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("DisplayConfigValues"), &m_dispidDisplayConfigValues);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("DoSetup"), &m_dispidDoSetup);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadConfig"), &m_dispidReadConfig);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("WriteConfig"), &m_dispidWriteConfig);
		Utils_GetMemid(this->m_pTypeInfo, L"GetGratingDispersion", &m_dispidGetGratingDispersion);
		Utils_GetMemid(this->m_pTypeInfo, L"ScanStart", &m_dispidScanStart);
		pTypeInfo->Release();
	}
}

CMyObject::CImp_clsIMono::~CImp_clsIMono()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImp_clsIMono::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIMono::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIMono::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImp_clsIMono::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImp_clsIMono::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	if (NULL != this->m_pTypeInfo)
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
		return S_OK;
	}
	else
	{
		*ppTInfo	= NULL;
		return E_FAIL;
	}
}

STDMETHODIMP CMyObject::CImp_clsIMono::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImp_clsIMono::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	if (dispIdMember == m_dispidCurrentWavelength)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetCurrentWavelength(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetCurrentWavelength(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidAutoGrating)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAutoGrating(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetAutoGrating(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidCurrentGrating)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetCurrentGrating(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetCurrentGrating(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidAmBusy)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAmBusy(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidAmInitialized)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAmInitialized(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetAmInitialized(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidGratingParams)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetGratingParams(pDispParams, pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetGratingParams(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidMonoParams)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetMonoParams(pDispParams, pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetMonoParams(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidIsValidPosition)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->IsValidPosition(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == m_dispidConvertStepsToNM)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ConvertStepsToNM(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == m_dispidDisplayConfigValues)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->DisplayConfigValues();
		}
	}
	else if (dispIdMember == m_dispidDoSetup)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->DoSetup();
		}
	}
	else if (dispIdMember == m_dispidReadConfig)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ReadConfig(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == m_dispidWriteConfig)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->WriteConfig(pDispParams);
		}
	}
	else if (dispIdMember == this->m_dispidGetGratingDispersion)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->GetGratingDispersion(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == this->m_dispidScanStart)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ScanStart();
		}
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImp_clsIMono::GetCurrentWavelength(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetCurrentWavelength(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetCurrentWavelength(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetCurrentWavelength(varg.dblVal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetAutoGrating(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetAutoGrating(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetAutoGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetAutoGrating(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetCurrentGrating(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetCurrentGrating(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetCurrentGrating(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetCurrentGrating(varg.lVal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetAmBusy(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetAmBusy(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetAmInitialized(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetAmInitialized(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetAmInitialized(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetAmInitialized(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetGratingParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
//	UINT				uArgErr;
	VariantInit(&varg);

	if (NULL == pVarResult) return E_INVALIDARG;
	short int iGrating;
	short int param;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	param = varg.iVal;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	iGrating = varg.iVal;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->GetGratingParams(iGrating, param, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetGratingParams(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
//	UINT				uArgErr;
	VariantInit(&varg);
	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	short int iGrating;
	short int param;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	param = varg.iVal;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[2]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	iGrating = varg.iVal;
	VariantCopy(&varg, &(pDispParams->rgvarg[0]));
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetGratingParams(iGrating, param, &varg);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetMonoParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
//	UINT				uArgErr;
	VariantInit(&varg);

	if (NULL == pVarResult) return E_INVALIDARG;
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMyUSBDoubleAdditive->GetMonoParams(varg.iVal, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetMonoParams(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
//	UINT				uArgErr;
	VariantInit(&varg);
	short int param;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
	if (FAILED(hr)) return hr;
	hr = VariantChangeType(&varg, &varg, 0, VT_I2);
	if (FAILED(hr)) return hr;
	param = varg.iVal;
	VariantCopy(&varg, &(pDispParams->rgvarg[0]));
	this->m_pMyObject->m_pMyUSBDoubleAdditive->SetMonoParams(param, &varg);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::IsValidPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);

	if (NULL == pVarResult) return E_INVALIDARG;
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromBoolean(this->m_pMyObject->m_pMyUSBDoubleAdditive->IsValidPosition(varg.dblVal), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
//	HRESULT				hr;
	VARIANTARG			varg;
//	UINT				uArgErr;
	VariantInit(&varg);

	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::DisplayConfigValues()
{
	CDlgDisplayConfigValues		dlg(this->m_pMyObject);
	dlg.DoOpenDialog(GetServer()->GetInstance(), this->m_pMyObject->FireRequestParentWindow());
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::DoSetup()
{
	this->m_pMyObject->m_pMyUSBDoubleAdditive->DoSetup();
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szConfig		= NULL;
	BOOL				fSuccess		= FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szConfig);
//	VariantClear(&varg);
	fSuccess	= this->m_pMyObject->m_pMyUSBDoubleAdditive->ReadConfig(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szConfig);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::WriteConfig(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
//	LPTSTR				szConfig		= NULL;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
//	Utils_UnicodeToAnsi(varg.bstrVal, &szConfig);
//	VariantClear(&varg);
	this->m_pMyObject->m_pMyUSBDoubleAdditive->WriteConfig(varg.bstrVal);
	VariantClear(&varg);
//	CoTaskMemFree((LPVOID) szConfig);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetGratingDispersion(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromDouble(this->m_pMyObject->m_pMyUSBDoubleAdditive->GetGratingDispersion(varg.lVal), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::ScanStart()
{
	this->m_pMyObject->m_pMyUSBDoubleAdditive->ScanStart();
	return S_OK;
}


CMyObject::CImpIPersistStreamInit::CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIPersistStreamInit::~CImpIPersistStreamInit()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetClassID(
									CLSID			*	pClassID)
{
	*pClassID	= MY_CLSID;
	return S_OK;
}

// IPersistStreamInit methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::IsDirty(void)
{
	return this->m_pBackObj->m_pMyUSBDoubleAdditive->GetDirty() ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Load(
									LPSTREAM			pStm)
{
	return this->m_pBackObj->m_pMyUSBDoubleAdditive->Load(pStm);
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Save(
									LPSTREAM			pStm ,  
									BOOL				fClearDirty)
{
	return this->m_pBackObj->m_pMyUSBDoubleAdditive->Save(pStm, fClearDirty);
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetSizeMax(
									ULARGE_INTEGER	*	pcbSize)
{
	pcbSize->QuadPart	= this->m_pBackObj->m_pMyUSBDoubleAdditive->GetSizeMax();
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::InitNew(void)
{
	return this->m_pBackObj->m_pMyUSBDoubleAdditive->InitNew();
}

CMyObject::CImpISpecifyPropertyPages::CImpISpecifyPropertyPages(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpISpecifyPropertyPages::~CImpISpecifyPropertyPages()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpISpecifyPropertyPages::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpISpecifyPropertyPages::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpISpecifyPropertyPages::Release()
{
	return this->m_punkOuter->Release();
}

// ISpecifyPropertyPages method
STDMETHODIMP CMyObject::CImpISpecifyPropertyPages::GetPages(
									CAUUID			*	pPages)
{
	pPages->cElems		= 1;
	pPages->pElems		= (GUID*) CoTaskMemAlloc(1 * sizeof(GUID));
	pPages->pElems[0]	= CLSID_MyPropPage;
	return S_OK;
}