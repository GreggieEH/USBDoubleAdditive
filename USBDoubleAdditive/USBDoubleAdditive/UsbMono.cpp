#include "stdafx.h"
#include "UsbMono.h"
#include "MyUSBDoubleAdditive.h"
#include "Server.h"

CUsbMono::CUsbMono(CMyUSBDoubleAdditive * pMyUSBDoubleAdditive, LPCTSTR szName) :
	m_pMyUSBDoubleAdditive(pMyUSBDoubleAdditive),
	m_pdisp(NULL),
	// properties and methods
	m_dispidReadConfig(DISPID_UNKNOWN),
	m_dispidAmInitialized(DISPID_UNKNOWN),
	m_dispidCurrentWavelength(DISPID_UNKNOWN),
	m_dispidCurrentGrating(DISPID_UNKNOWN),
	m_dispidConvertStepsToNM(DISPID_UNKNOWN),
	m_dispidDoSetup(DISPID_UNKNOWN),
	m_dispidDisplayConfigValues(DISPID_UNKNOWN),
	// sink connection
	m_iid__clsIMono(IID_NULL),
	m_dwCookie(0)
{
	StringCchCopy(this->m_szName, MAX_PATH, szName);
}

CUsbMono::~CUsbMono(void)
{
	if (NULL != this->m_pdisp)
	{
		IDispatch	*	pdisp;
		HRESULT			hr;
		DISPID			dispid;
		hr = this->m_pdisp->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
		if (SUCCEEDED(hr))
		{
			Utils_GetMemid(pdisp, TEXT("RemoveNamedObject"), &dispid);
			Utils_InvokeMethod(pdisp, dispid, NULL, 0, NULL);
			pdisp->Release();
		}

TCHAR			szMessage[MAX_PATH];
StringCchPrintf(szMessage, MAX_PATH, TEXT("Closing IMono connection cookie = %3d"),
				this->m_dwCookie);
GetServer()->DoLogString(szMessage);

		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, this->m_iid__clsIMono, FALSE,
			&(this->m_dwCookie));
		ULONG			cRefs		= this->m_pdisp->Release();
		this->m_pdisp	= NULL;
//		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
}

BOOL CUsbMono::doInit()
{
	HRESULT				hr;
	LPOLESTR			ProgID		= NULL;
	CLSID				clsid;
	IUnknown		*	pUnk		= NULL;
	IProvideClassInfo*	pProvideClassInfo;
	ITypeInfo		*	pClassInfo	= NULL;
	ITypeInfo		*	pTypeInfo	= NULL;
	TYPEATTR		*	pTypeAttr;
	CImpISink		*	pSink;
	IUnknown		*	pUnkSink;
	
	Utils_RELEASE_INTERFACE(this->m_pdisp);
	SHStrDup(TEXT("Sciencetech.SciUsbMono.1"), &ProgID);
	hr = CLSIDFromProgID(ProgID, &clsid);
	CoTaskMemFree((LPVOID*) ProgID);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (LPVOID*) &pUnk);
	}
	if (SUCCEEDED(hr))
	{
		hr = pUnk->QueryInterface(IID_IProvideClassInfo, (LPVOID*) &pProvideClassInfo);
		if (SUCCEEDED(hr))
		{
			hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
			pProvideClassInfo->Release();
		}
		if (SUCCEEDED(hr))
		{
			hr = Utils_FindImplClassName(pClassInfo, TEXT("_clsIMono"), &pTypeInfo) ? S_OK : E_FAIL;
			pClassInfo->Release();
		}
		if (SUCCEEDED(hr))
		{
			hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
			if (SUCCEEDED(hr))
			{
				hr = pUnk->QueryInterface(pTypeAttr->guid, (LPVOID*) &(this->m_pdisp));
				pTypeInfo->ReleaseTypeAttr(pTypeAttr);
			}
			pTypeInfo->Release();
		}
		pUnk->Release();
	}
	if (SUCCEEDED(hr))
	{
		ULONG			cRefs = this->m_pdisp->AddRef();
		this->m_pdisp->Release();
	}
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(this->m_pdisp, TEXT("ReadConfig"), &m_dispidReadConfig);
		Utils_GetMemid(this->m_pdisp, TEXT("AmInitialized"), &m_dispidAmInitialized);
		Utils_GetMemid(this->m_pdisp, TEXT("CurrentWavelength"), &m_dispidCurrentWavelength);
		Utils_GetMemid(this->m_pdisp, TEXT("CurrentGrating"), &m_dispidCurrentGrating);
		Utils_GetMemid(this->m_pdisp, TEXT("ConvertStepsToNM"), &m_dispidConvertStepsToNM);
		Utils_GetMemid(this->m_pdisp, TEXT("DoSetup"), &m_dispidDoSetup);
		Utils_GetMemid(this->m_pdisp, TEXT("DisplayConfigValues"), &m_dispidDisplayConfigValues);
		ULONG			cRefs = this->m_pdisp->AddRef();
		this->m_pdisp->Release();
		// set the display name
		this->SetDisplayName();
		cRefs = this->m_pdisp->AddRef();
		this->m_pdisp->Release();
		// connect the sink
		pSink	= new CImpISink(this);
		cRefs = this->m_pdisp->AddRef();
		this->m_pdisp->Release();
		hr = pSink->QueryInterface(this->m_iid__clsIMono, (LPVOID*) &pUnkSink);
		if (SUCCEEDED(hr))
		{
			Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, this->m_iid__clsIMono,
				TRUE, &(this->m_dwCookie));
			pUnkSink->Release();
		}
		else
			delete pSink;
	}
	if (SUCCEEDED(hr))
	{
		ULONG			cRefs = this->m_pdisp->AddRef();
		this->m_pdisp->Release();
	}
	return SUCCEEDED(hr);
}

// set the display name
void CUsbMono::SetDisplayName()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	DISPID			dispid;

	if (NULL == this->m_pdisp) return;
	hr = this->m_pdisp->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, TEXT("DisplayName"), &dispid);
		Utils_SetStringProperty(pdisp, dispid, this->m_szName);
		pdisp->Release();
	}
}

BOOL CUsbMono::ReadConfig(
								LPCTSTR				szReadConfig)
{
	HRESULT			hr;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess		= FALSE;
	InitVariantFromString(szReadConfig, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, this->m_dispidReadConfig, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr))VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

void CUsbMono::WriteConfig(
								LPCTSTR				configfile)
{
	DISPID			dispid;
	VARIANTARG		varg;
	Utils_GetMemid(this->m_pdisp, TEXT("WriteConfig"), &dispid);
	InitVariantFromString(configfile, &varg);
	Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, NULL);
	VariantClear(&varg);
}

BOOL CUsbMono::GetAmInitialized()
{
	return Utils_GetBoolProperty(this->m_pdisp, this->m_dispidAmInitialized);
}

void CUsbMono::SetAmInitialized(
								BOOL				AmInitialized)
{
	Utils_SetBoolProperty(this->m_pdisp, this->m_dispidAmInitialized, AmInitialized);
}

double CUsbMono::GetCurrentWavelength()
{
	return Utils_GetDoubleProperty(this->m_pdisp, this->m_dispidCurrentWavelength);
}

void CUsbMono::SetCurrentWavelength(
								double				CurrentWavelength)
{
	Utils_SetDoubleProperty(this->m_pdisp, this->m_dispidCurrentWavelength, CurrentWavelength);
}

long CUsbMono::GetCurrentGrating()
{
	return Utils_GetLongProperty(this->m_pdisp, this->m_dispidCurrentGrating);
}

void CUsbMono::SetCurrentGrating(
								long				CurrentGrating)
{
	Utils_SetLongProperty(this->m_pdisp, this->m_dispidCurrentGrating, CurrentGrating);
}

BOOL CUsbMono::ConvertStepsToNM(
								BOOL				fStepsToNM, 
								long				gratingID, 
								long			*	steps, 
								double			*	nm)
{
	HRESULT				hr;
	VARIANTARG			avarg[4];
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	InitVariantFromBoolean(fStepsToNM, &avarg[3]);
	InitVariantFromInt32(gratingID, &avarg[2]);
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I4;
	avarg[1].plVal		= steps;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= nm;
	hr = Utils_InvokeMethod(this->m_pdisp, this->m_dispidConvertStepsToNM, avarg, 4, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

void CUsbMono::DoSetup()
{
	Utils_InvokeMethod(this->m_pdisp, this->m_dispidDoSetup, NULL, 0, NULL);
}

void CUsbMono::DisplayConfigValues()
{
	Utils_InvokeMethod(this->m_pdisp, this->m_dispidDisplayConfigValues, NULL, 0, NULL);
}

// persistence
BOOL CUsbMono::GetDirty()
{
	HRESULT				hr;
	IPersistStreamInit*	pPersistStreamInit;
	BOOL				fDirty		= FALSE;
	if (NULL == this->m_pdisp) return FALSE;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
	if (SUCCEEDED(hr))
	{
		fDirty = S_OK == pPersistStreamInit->IsDirty();
		pPersistStreamInit->Release();
	}
	return fDirty;
}

HRESULT CUsbMono::Load(
								LPSTREAM			pStm)
{
	HRESULT				hr;
	IPersistStreamInit*	pPersistStreamInit;
	if (NULL == this->m_pdisp) return E_FAIL;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
	if (SUCCEEDED(hr))
	{
		hr = pPersistStreamInit->Load(pStm);
		pPersistStreamInit->Release();
	}
	return hr;
}


HRESULT CUsbMono::Save(
								LPSTREAM			pStm ,  
								BOOL				fClearDirty)
{
	HRESULT				hr;
	IPersistStreamInit*	pPersistStreamInit;
	if (NULL == this->m_pdisp) return E_FAIL;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
	if (SUCCEEDED(hr))
	{
		hr = pPersistStreamInit->Save(pStm, fClearDirty);
		pPersistStreamInit->Release();
	}
	return hr;
}

long CUsbMono::GetSizeMax()
{
	HRESULT				hr;
	IPersistStreamInit*	pPersistStreamInit;
	ULARGE_INTEGER		cbSize;
	long				nSize		= 0;
	if (NULL == this->m_pdisp) return 0;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
	if (SUCCEEDED(hr))
	{
		hr = pPersistStreamInit->GetSizeMax(&cbSize);
		if (SUCCEEDED(hr)) nSize = (long) cbSize.QuadPart;
		pPersistStreamInit->Release();
	}
	return nSize;
}

HRESULT CUsbMono::InitNew()
{
	HRESULT				hr;
	IPersistStreamInit*	pPersistStreamInit;
	if (NULL == this->m_pdisp) return E_FAIL;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStreamInit);
	if (SUCCEEDED(hr))
	{
		hr = pPersistStreamInit->InitNew();
		pPersistStreamInit->Release();
	}
	return hr;
}

BOOL CUsbMono::GetAutoGrating()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AutoGrating"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CUsbMono::SetAutoGrating(
								BOOL				AutoGrating)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AutoGrating"), &dispid);
	Utils_SetBoolProperty(this->m_pdisp, dispid, AutoGrating);
}

BOOL CUsbMono::GetAmBusy()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdisp, TEXT("AmBusy"), &dispid);
	return Utils_GetBoolProperty(this->m_pdisp, dispid);
}

void CUsbMono::GetGratingParams(
								long				grating, 
								short				paramIndex, 
								VARIANT			*	Value)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdisp, TEXT("GratingParams"), &dispid);
	InitVariantFromInt32(grating, &avarg[1]);
	InitVariantFromInt16(paramIndex, &avarg[0]);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYGET, avarg, 2, Value);
}

void CUsbMono::SetGratingParams(
								long				grating, 
								short				paramIndex, 
								VARIANT			*	Value)
{
	DISPID				dispid;
	VARIANTARG			avarg[3];
	Utils_GetMemid(this->m_pdisp, TEXT("GratingParams"), &dispid);
	InitVariantFromInt32(grating, &avarg[2]);
	InitVariantFromInt16(paramIndex, &avarg[1]);
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], Value);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 3, NULL);
}

void CUsbMono::GetMonoParams(
								short				paramIndex, 
								VARIANT			*	Value)
{
	DISPID				dispid;
	VARIANTARG			varg;
	Utils_GetMemid(this->m_pdisp, TEXT("MonoParams"), &dispid);
	InitVariantFromInt16(paramIndex, &varg);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYGET, &varg, 1, Value);
}

void CUsbMono::SetMonoParams(
								short				paramIndex, 
								VARIANT			*	Value)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdisp, TEXT("MonoParams"), &dispid);
	InitVariantFromInt16(paramIndex, &avarg[1]);
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], Value);
	Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYPUT, avarg, 2, NULL);
	VariantClear(&avarg[0]);
}

BOOL CUsbMono::IsValidPosition(
								double				position)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(this->m_pdisp, TEXT("IsValidPosition"), &dispid);
	InitVariantFromDouble(position, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

// get the property page
BOOL CUsbMono::GetPropPage(
								IUnknown	**		ppUnk,
								CLSID		*		pClsidPage)
{
	return ::GetObjectPropPage(this->m_pdisp, ppUnk, pClsidPage);
}

// sink events
void CUsbMono::OnGratingChanged(
								long				grating)
{
	this->m_pMyUSBDoubleAdditive->OnMonoGratingChanged(this->m_szName, grating);
}

void CUsbMono::OnMoveCompleted(
								double				newPosition)
{
	this->m_pMyUSBDoubleAdditive->OnMonoMoveCompleted(this->m_szName, newPosition);
}

void CUsbMono::OnAutoGratingPropChanged(
								BOOL				fAutoGrating)
{
	this->m_pMyUSBDoubleAdditive->OnMonoAutoGratingPropChange(this->m_szName, fAutoGrating);
}

CUsbMono::CImpISink::CImpISink(CUsbMono * pUsbMono) :
	m_pUsbMono(pUsbMono),
	m_cRefs(0),
	m_dispidGratingChanged(DISPID_UNKNOWN),
	m_dispidMoveCompleted(DISPID_UNKNOWN),
	m_dispidAutoGratingPropChanged(DISPID_UNKNOWN)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo;
	ITypeInfo		*	pClassInfo		= NULL;
	ITypeInfo		*	pTypeInfo		= NULL;
	TYPEATTR		*	pTypeAttr;

	if (NULL == this->m_pUsbMono->m_pdisp) return;
	hr = this->m_pUsbMono->m_pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*) &pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = Utils_FindImplClassName(pClassInfo, TEXT("__clsIMono"), &pTypeInfo)
			? S_OK : E_FAIL;
		pClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pUsbMono->m_iid__clsIMono	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, TEXT("GratingChanged"), &m_dispidGratingChanged);
		Utils_GetMemid(pTypeInfo, TEXT("MoveCompleted"), &m_dispidMoveCompleted);
		Utils_GetMemid(pTypeInfo, TEXT("AutoGratingPropChanged"), &m_dispidAutoGratingPropChanged);
		pTypeInfo->Release();
	}
}

CUsbMono::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CUsbMono::CImpISink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || riid == this->m_pUsbMono->m_iid__clsIMono)
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

STDMETHODIMP_(ULONG) CUsbMono::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CUsbMono::CImpISink::Release()
{
	ULONG			cRefs		= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CUsbMono::CImpISink::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CUsbMono::CImpISink::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CUsbMono::CImpISink::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CUsbMono::CImpISink::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	if (dispIdMember == this->m_dispidGratingChanged)
	{
		hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
		if (SUCCEEDED(hr)) this->m_pUsbMono->OnGratingChanged(varg.lVal);
	}
	else if (dispIdMember == this->m_dispidMoveCompleted)
	{
		hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
		if (SUCCEEDED(hr)) this->m_pUsbMono->OnMoveCompleted(varg.dblVal);
	}
	else if (dispIdMember == this->m_dispidAutoGratingPropChanged)
	{
		hr = DispGetParam(pDispParams, 0, VT_BOOL, &varg, &uArgErr);
		if (SUCCEEDED(hr)) this->m_pUsbMono->OnAutoGratingPropChanged(
			VARIANT_TRUE == varg.boolVal);
	}
	return S_OK;
}