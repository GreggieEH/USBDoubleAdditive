#pragma once

class CMyUSBDoubleAdditive;

class CUsbMono
{
public:
	CUsbMono(CMyUSBDoubleAdditive * pMyUSBDoubleAdditive, LPCTSTR szName);
	~CUsbMono(void);
	BOOL					doInit();
	BOOL					ReadConfig(
								LPCTSTR				szReadConfig);
	void					WriteConfig(
								LPCTSTR				configfile);
	BOOL					GetAmInitialized();
	void					SetAmInitialized(
								BOOL				AmInitialized);
	double					GetCurrentWavelength();
	void					SetCurrentWavelength(
								double				CurrentWavelength);
	long					GetCurrentGrating();
	void					SetCurrentGrating(
								long				CurrentGrating);
	BOOL					ConvertStepsToNM(
								BOOL				fStepsToNM, 
								long				gratingID, 
								long			*	steps, 
								double			*	nm);
	void					DoSetup();
	void					DisplayConfigValues();
	BOOL					GetAutoGrating();
	void					SetAutoGrating(
								BOOL				AutoGrating);
	BOOL					GetAmBusy();
	void					GetGratingParams(
								long				grating, 
								short				paramIndex, 
								VARIANT			*	Value);
	void					SetGratingParams(
								long				grating, 
								short				paramIndex, 
								VARIANT			*	Value);
	void					GetMonoParams(
								short				paramIndex, 
								VARIANT			*	Value);
	void					SetMonoParams(
								short				paramIndex, 
								VARIANT			*	Value);
	BOOL					IsValidPosition(
								double				position);
	// persistence
	BOOL					GetDirty();
	HRESULT					Load(
								LPSTREAM			pStm);
	HRESULT					Save(
								LPSTREAM			pStm ,  
								BOOL				fClearDirty);
	long					GetSizeMax();
	HRESULT					InitNew();
	// get the property page
	BOOL					GetPropPage(
								IUnknown	**		ppUnk,
								CLSID		*		pClsidPage);
protected:
	// set the display name
	void					SetDisplayName();
	// sink events
	void					OnGratingChanged(
								long				grating);
	void					OnMoveCompleted(
								double				newPosition);
	void					OnAutoGratingPropChanged(
								BOOL				fAutoGrating);
private:
	CMyUSBDoubleAdditive*	m_pMyUSBDoubleAdditive;
	TCHAR					m_szName[MAX_PATH];
	IDispatch			*	m_pdisp;
	// properties and methods
	DISPID					m_dispidReadConfig;
	DISPID					m_dispidAmInitialized;
	DISPID					m_dispidCurrentWavelength;
	DISPID					m_dispidCurrentGrating;
	DISPID					m_dispidConvertStepsToNM;
	DISPID					m_dispidDoSetup;
	DISPID					m_dispidDisplayConfigValues;
	// sink connection
	IID						m_iid__clsIMono;
	DWORD					m_dwCookie;

	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CUsbMono * pUsbMono);
		~CImpISink();
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
	private:
		CUsbMono			*	m_pUsbMono;
		ULONG					m_cRefs;
		DISPID					m_dispidGratingChanged;
		DISPID					m_dispidMoveCompleted;
		DISPID					m_dispidAutoGratingPropChanged;
	};
	friend CImpISink;

};
