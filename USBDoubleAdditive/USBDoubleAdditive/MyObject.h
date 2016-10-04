#pragma once

class CMyUSBDoubleAdditive;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	void						FireMonoPositionChanged(
									LPCTSTR			Mono,
									double			position);
	void						FireMonoGratingChanged(
									LPCTSTR			Mono,
									long			grating);
	void						FireMonoAutoGratingChanged(
									LPCTSTR			Mono,
									BOOL			AutoGrating);
	// __clsIMono sink events
	void						FireGratingChanged(
									long			grating);
	BOOL						FireBeforeMoveChange(
									double			newPosition);
	void						FireHaveAborted();
	void						FireMoveCompleted(
									double			newPosition);
	void						FireMoveError(	
									LPCTSTR			err);
	BOOL						FireRequestChangeGrating(
									long			newGrating);
	HWND						FireRequestParentWindow();
	void						FireStatusMessage(
									LPCTSTR			msg, 
									BOOL			AmBusy);
protected:
	HRESULT						GetClassInfo(
									ITypeInfo	**	ppTI);
	HRESULT						GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo);
	HRESULT						Init__clsIMono();
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					GetMonoDecoupled(
									VARIANT		*	pVarResult);
		HRESULT					SetMonoDecoupled(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetActiveMono(
									VARIANT		*	pVarResult);
		HRESULT					SetActiveMono(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetSystemInit(
									VARIANT		*	pVarResult);
		HRESULT					SetSystemInit(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGrating(
									VARIANT		*	pVarResult);
		HRESULT					SetGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					DisplayMonoSetup(
									DISPPARAMS	*	pDispParams);
		HRESULT					SetMonoPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetMonoPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					DisplayMonoConfigValues(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoGrating(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetMonoGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoAutoGrating(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetMonoAutoGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetReInitOnScanStart(
									VARIANT		*	pVarResult);
		HRESULT					SetReInitOnScanStart(
									DISPPARAMS	*	pDispParams);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImp_clsIMono : public IDispatch
	{
	public:
		CImp_clsIMono(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImp_clsIMono();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					GetCurrentWavelength(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentWavelength(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAutoGrating(
									VARIANT		*	pVarResult);
		HRESULT					SetAutoGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetCurrentGrating(
									VARIANT		*	pVarResult);
		HRESULT					SetCurrentGrating(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetAmBusy(
									VARIANT		*	pVarResult);
		HRESULT					GetAmInitialized(
									VARIANT		*	pVarResult);
		HRESULT					SetAmInitialized(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGratingParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetGratingParams(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetMonoParams(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					SetMonoParams(
									DISPPARAMS	*	pDispParams);
		HRESULT					IsValidPosition(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					DisplayConfigValues();
		HRESULT					DoSetup();
		HRESULT					ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					WriteConfig(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGratingDispersion(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ScanStart();
	private:
		CMyObject			*	m_pMyObject;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		DISPID					m_dispidCurrentWavelength;
		DISPID					m_dispidAutoGrating;
		DISPID					m_dispidCurrentGrating;
		DISPID					m_dispidAmBusy;
		DISPID					m_dispidAmInitialized;
		DISPID					m_dispidGratingParams;
		DISPID					m_dispidMonoParams;
		DISPID					m_dispidIsValidPosition;
		DISPID					m_dispidConvertStepsToNM;
		DISPID					m_dispidDisplayConfigValues;
		DISPID					m_dispidDoSetup;
		DISPID					m_dispidReadConfig;
		DISPID					m_dispidWriteConfig;
		DISPID					m_dispidGetGratingDispersion;
		DISPID					m_dispidScanStart;
	};
	class CImpIPersistStreamInit : public IPersistStreamInit
	{
	public:
		CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistStreamInit();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID			*	pClassID);
		// IPersistStreamInit methods
		STDMETHODIMP			IsDirty(void);
		STDMETHODIMP			Load(
									LPSTREAM			pStm);
		STDMETHODIMP			Save(
									LPSTREAM			pStm ,  
									BOOL				fClearDirty);
		STDMETHODIMP			GetSizeMax(
									ULARGE_INTEGER	*	pcbSize);
		STDMETHODIMP			InitNew(void);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpISpecifyPropertyPages : public ISpecifyPropertyPages 
	{
	public:
		CImpISpecifyPropertyPages(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpISpecifyPropertyPages();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// ISpecifyPropertyPages method
		STDMETHODIMP			GetPages(
									CAUUID			*	pPages);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	friend CImpIDispatch;
	friend CImpIConnectionPointContainer;
	friend CImpIProvideClassInfo2;
	friend CImp_clsIMono;
	friend CImpIPersistStreamInit;
	friend CImpISpecifyPropertyPages;
	// data members
	CImpIDispatch			*	m_pImpIDispatch;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIProvideClassInfo2	*	m_pImpIProvideClassInfo2;
	CImp_clsIMono			*	m_pImp_clsIMono;
	CImpIPersistStreamInit	*	m_pImpIPersistStreamInit;
	CImpISpecifyPropertyPages*	m_pImpISpecifyPropertyPages;
	// outer unknown for aggregation
	IUnknown				*	m_pUnkOuter;
	// object reference count
	ULONG						m_cRefs;
	// connection points array
	IConnectionPoint		*	m_paConnPts[MAX_CONN_PTS];
	CMyUSBDoubleAdditive	*	m_pMyUSBDoubleAdditive;
	// _clsIMono interface id
	IID							m_iid_clsIMono;
	// __clsIMono events	
	IID							m_iid__clsIMono;
	DISPID						m_dispidGratingChanged;
	DISPID						m_dispidBeforeMoveChange;
	DISPID						m_dispidHaveAborted;
	DISPID						m_dispidMoveCompleted;
	DISPID						m_dispidMoveError;
	DISPID						m_dispidRequestChangeGrating;
	DISPID						m_dispidRequestParentWindow;
	DISPID						m_dispidStatusMessage;
};
