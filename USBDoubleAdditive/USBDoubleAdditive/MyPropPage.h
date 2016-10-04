#pragma once

class CMyPropPage : public IPropertyPage
{
public:
	CMyPropPage(void);
	~CMyPropPage(void);
	// IUnknown methods
	STDMETHODIMP			QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();
	// IPropertyPage methods
	STDMETHODIMP			SetPageSite( 
								IPropertyPageSite *pPageSite);
	STDMETHODIMP			Activate( 
								HWND			hWndParent ,  //Parent window handle
								LPCRECT			prc ,      //Pointer to RECT structure
								BOOL			bModal);
	STDMETHODIMP			Deactivate(void);
	STDMETHODIMP			GetPageInfo( 
								PROPPAGEINFO *	pPageInfo);
	STDMETHODIMP			SetObjects( 
								ULONG			cObjects ,  //Number of IUnknown pointers in the ppUnk array
								IUnknown	**	ppUnk);
	STDMETHODIMP			Show(
								UINT			nCmdShow);
	STDMETHODIMP			Move( 
								LPCRECT			prc);
	STDMETHODIMP			IsPageDirty(void);
	STDMETHODIMP			Apply(void);
	STDMETHODIMP			Help( 
								LPCOLESTR		pszHelpDir);
	STDMETHODIMP			TranslateAccelerator( 
								LPMSG			pMsg);
protected:
	// map dialog units to pixels
	void					DialogUnitsToPixels(
								HDC				hdc,
								LPRECT			prc);
	// set the dirty flag
	void					SetDirty(
								BOOL			fDirty);
	// dialog procedure
	BOOL					DlgProc(
								UINT			uMsg,
								WPARAM			wParam,
								LPARAM			lParam);
	BOOL					OnInitDialog();		// handler for WM_INITDIALOG
	BOOL					OnCommand(			// command handler
								WORD			wmID,
								WORD			wmEvent);
	BOOL					OnNotify(
								LPNMHDR			pnmh);
	void					OnEditBoxReturnClicked(
								UINT			nEditBox);
	// subclass an edit box
	void					SubclassEditBox(
								UINT			uControlID);
	// display mono coupled
	void					DisplayMonoCoupled();
	void					DisplayActiveMono();
	void					DisplaySystemInit();
	// property change notifications
	void					OnPropChanged(
								DISPID			dispid);
	// mono position
	double					GetCurrentWavelength();
	void					SetCurrentWavelength(
								double			CurrentWavelength);
	long					GetCurrentGrating();
	void					SetCurrentGrating(
								long			CurrentGrating);
	BOOL					Get_clsIMono(
								IDispatch	**	ppdisp);
	// get mono position
	double					GetMonoPosition(
								LPCTSTR			szMono);
	// set mono position
	void					SetMonoPosition(
								LPCTSTR			szMono,
								double			position);
	void					OnMonoPositionChanged(
								LPCTSTR			szMono,
								double			position);
	// mono grating changed
	void					OnMonoGratingChanged(
								LPCTSTR			szMono,
								long			grating);
	long					GetMonoGrating(
								LPCTSTR			Mono);
	void					SetMonoGrating(
								LPCTSTR			Mono,
								long			grating);
	// auto grating prop change
	void					OnMonoAutoGratingChanged(
								LPCTSTR			szMono,
								BOOL			AutoGrating);
	BOOL					GetMonoAutoGrating(
								LPCTSTR			szMono);
	void					SetMonoAutoGrating(
								LPCTSTR			szMono,
								BOOL			AutoGrating);
	void					OnSetGrating();
	void					OnSetPosition();
	BOOL					GetActiveMono(
								LPTSTR			szMono,
								UINT			nBufferSize);
	void					OnClickedChkAutoGrating();		// click event on auto grating control
	void					DisplayDispersion();
	void					DisplayReInitOnScanStart();
	void					OnClickedInitOnScanStart();
private:
	ULONG					m_cRefs;			// object reference count
	HWND					m_hwndDlg;			// property page window handle
	IPropertyPageSite	*	m_pPropertyPageSite;	// property page site object
	IDispatch			*	m_pdisp;			// object which we are communicating
	BOOL					m_fDirty;			// dirty flag
	DWORD					m_dwPropNotify;		//  property notify cookie
	DWORD					m_dwCookie;			// connection cookie

// property page procedure
friend LRESULT CALLBACK MyPropertyPageProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
// edit box subclass
friend LRESULT CALLBACK MyEditBoxSubclass(HWND, UINT, WPARAM, LPARAM);

// structure sent in with an edit box
struct MY_SUBCLASS_STRUCT
{
	CMyPropPage			*	pMyPropPage;
	WNDPROC					wpOrig;
};

	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CMyPropPage * pMyPropPage);
		~CImpIPropNotify();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotifySink methods
		STDMETHODIMP			OnChanged( 
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit( 
									DISPID			dispID);
	private:
		CMyPropPage		*	m_pMyPropPage;
		ULONG				m_cRefs;
	};
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CMyPropPage * pMyPropPage);
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
	protected:
		BOOL				GetMono(
								DISPPARAMS	*	pDispParams,
								LPTSTR			szMono,
								UINT			nBufferSize);
	private:
		CMyPropPage		*	m_pMyPropPage;
		ULONG				m_cRefs;
	};
	friend CImpIPropNotify;
	friend CImpISink;
};
