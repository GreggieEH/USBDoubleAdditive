#include "StdAfx.h"
#include "MyPropPage.h"
#include "Server.h"
#include "dispids.h"
#include "MyGuids.h"

CMyPropPage::CMyPropPage(void) :
	m_cRefs(0),					// object reference count
	m_hwndDlg(NULL),			// property page window handle
	m_pPropertyPageSite(NULL),	// property page site object
	m_pdisp(NULL),				// object which we are communicating
	m_fDirty(FALSE),			// dirty flag
	m_dwPropNotify(0),			//  property notify cookie
	m_dwCookie(0)				// connection cookie
{
}

CMyPropPage::~CMyPropPage(void)
{
	Utils_RELEASE_INTERFACE(this->m_pPropertyPageSite);
	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, MY_IIDSINK, FALSE,
			&(this->m_dwCookie));
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, IID_IPropertyNotifySink,
			FALSE, &(this->m_dwPropNotify));
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
}


// IUnknown methods
STDMETHODIMP CMyPropPage::QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyPage == riid)
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

STDMETHODIMP_(ULONG) CMyPropPage::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyPropPage::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IPropertyPage methods
STDMETHODIMP CMyPropPage::SetPageSite( 
								IPropertyPageSite *pPageSite)
{
	if (NULL == this->m_pPropertyPageSite)
	{
		if (NULL == pPageSite) return E_INVALIDARG;
		this->m_pPropertyPageSite	= pPageSite;
		this->m_pPropertyPageSite->AddRef();
	}
	else
	{
		// non null value
		if (NULL != pPageSite) return E_UNEXPECTED;
		Utils_RELEASE_INTERFACE(this->m_pPropertyPageSite);
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::Activate( 
								HWND			hWndParent ,  //Parent window handle
								LPCRECT			prc ,      //Pointer to RECT structure
								BOOL			bModal)
{
	HWND			hwndDlg;
	hwndDlg = CreateDialogParam(GetServer()->GetInstance(),
		MAKEINTRESOURCE(IDD_PROPPAGEUSBDoubleAdditive), 
		hWndParent,
		(DLGPROC) MyPropertyPageProc,
		(LPARAM) this);
	return NULL != hwndDlg ? S_OK : E_FAIL;
}

STDMETHODIMP CMyPropPage::Deactivate(void)
{
	if (NULL != this->m_hwndDlg)
	{
		DestroyWindow(this->m_hwndDlg);
		this->m_hwndDlg = NULL;
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::GetPageInfo( 
								PROPPAGEINFO *	pPageInfo)
{
	RECT				rc;							// client rectangle
	HDC					hdc;

	pPageInfo->cb		= sizeof(PROPPAGEINFO);
	// set the title
	pPageInfo->pszTitle	= NULL;
	// default title
	SHStrDup(TEXT("Double Additive Mono"), &(pPageInfo->pszTitle));
	pPageInfo->pszHelpFile	= NULL;
	pPageInfo->pszDocString = NULL;
	pPageInfo->dwHelpContext	= 0;
	if (NULL != this->m_hwndDlg)
	{
		GetWindowRect(this->m_hwndDlg, &rc);
	}
	else
	{
		hdc			= GetDC(NULL);
		rc.left		= 0;
		rc.top		= 0;
		rc.right	= 250;
		rc.bottom	= 110;
		this->DialogUnitsToPixels(hdc, &rc);
		ReleaseDC(NULL, hdc);
	}
	pPageInfo->size.cx = rc.right - rc.left;
	pPageInfo->size.cy = rc.bottom - rc.top;
	return S_OK;
}

// map dialog units to pixels
void CMyPropPage::DialogUnitsToPixels(
							HDC				hdc,
							LPRECT			prc)
{
	TEXTMETRIC				txtMetric;
	long					baseunitX;
	long					baseunitY;
	RECT					rcCopy;

	// get the text size
	GetTextMetrics(hdc, &txtMetric);
	baseunitX	= txtMetric.tmAveCharWidth;
	baseunitY	= (long) floor((txtMetric.tmHeight * 1.2) + 0.5);
	// copy the input rectangle
	CopyRect(&rcCopy, prc);
	// determine the output dimensions
	prc->left	= MulDiv(rcCopy.left, baseunitX, 4);
	prc->top	= MulDiv(rcCopy.top, baseunitY, 8);
	prc->right	= MulDiv(rcCopy.right, baseunitX, 4);
	prc->bottom	= MulDiv(rcCopy.bottom, baseunitY, 8);
}


STDMETHODIMP CMyPropPage::SetObjects( 
								ULONG			cObjects ,  //Number of IUnknown pointers in the ppUnk array
								IUnknown	**	ppUnk)
{
	HRESULT				hr;
	ULONG				i;
//	IDispatch		*	pdisp;
	BOOL				fDone;
//	LPTSTR				szName;
	// property notification sink
	CImpIPropNotify	*	pSink;
	IUnknown		*	pUnkSink;
	CImpISink		*	pISink;

	// disconnect the sinks
//	this->DisconnectCustomSink();
//	this->DisconnectPropertyNotification();
	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, MY_IIDSINK, FALSE,
			&(this->m_dwCookie));
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, IID_IPropertyNotifySink,
			FALSE, &(this->m_dwPropNotify));
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
	if (cObjects > 0)
	{
		// find the first non selected page
		fDone	= FALSE;
		i		= 0;
		while (i < cObjects && !fDone)
		{
			// have found the desired object
			hr = (ppUnk[i])->QueryInterface(IID_IUSBDoubleAdditive, (LPVOID*) &(this->m_pdisp));
			fDone = SUCCEEDED(hr);
			if (fDone)
			{
				pISink	= new CImpISink(this);
				hr = pISink->QueryInterface(MY_IIDSINK, (LPVOID*) &pUnkSink);
				if (SUCCEEDED(hr))
				{
					Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, MY_IIDSINK,
						TRUE, &(this->m_dwCookie));
					pUnkSink->Release();
				}
				else
					delete pISink;
				// property change notifications
				pSink	= new CImpIPropNotify(this);
				hr = pSink->QueryInterface(IID_IPropertyNotifySink, (LPVOID*) &pUnkSink);
				if (SUCCEEDED(hr))
				{	
					Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink,
						IID_IPropertyNotifySink, TRUE, &(this->m_dwPropNotify));
					pUnkSink->Release();
				}
				else delete pSink;

			}
			if (!fDone) i++;
		}
		if (!fDone) return E_NOINTERFACE;
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::Show(
								UINT			nCmdShow)
{
	if (NULL != this->m_hwndDlg)
	{
		ShowWindow(this->m_hwndDlg, nCmdShow);
		return S_OK;
	}
	else return E_UNEXPECTED;
}

STDMETHODIMP CMyPropPage::Move( 
								LPCRECT			prc)
{
	if (NULL != this->m_hwndDlg)
	{
		MoveWindow(this->m_hwndDlg,
			prc->left, prc->top,
			prc->right - prc->left,
			prc->bottom - prc->top,
			TRUE);
		return S_OK;
	}
	else return E_UNEXPECTED;
}

STDMETHODIMP CMyPropPage::IsPageDirty(void)
{
	return this->m_fDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CMyPropPage::Apply(void)
{
	this->m_fDirty	= FALSE;
	return S_OK;
}

STDMETHODIMP CMyPropPage::Help( 
								LPCOLESTR		pszHelpDir)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyPropPage::TranslateAccelerator( 
								LPMSG			pMsg)
{
	return E_NOTIMPL;
}

// property page procedure
LRESULT CALLBACK MyPropertyPageProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage			*	pMyPropPage		= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pMyPropPage		= (CMyPropPage *) lParam;
		pMyPropPage->m_hwndDlg	= hwndDlg;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, (LONG_PTR) lParam);
	}
	else
	{
		pMyPropPage		= (CMyPropPage *) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pMyPropPage)
		return pMyPropPage->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

// dialog procedure
BOOL CMyPropPage::DlgProc(
								UINT			uMsg,
								WPARAM			wParam,
								LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR) lParam);
	default:
		break;
	}
	return FALSE;
}

BOOL CMyPropPage::OnInitDialog()		// handler for WM_INITDIALOG
{
	this->DisplayMonoCoupled();
	this->DisplayActiveMono();
	// display the dispersion
	this->DisplayDispersion();
	this->DisplayReInitOnScanStart();
	return TRUE;
}

BOOL CMyPropPage::OnCommand(			// command handler
								WORD			wmID,
								WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_RADIOTURNOFFDECOUPLE:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled, FALSE);
			return TRUE;
		}
		break;
	case IDC_RADIODECOUPLE:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled, TRUE);
			return TRUE;
		}
		break;
	case IDC_RADIOENTRANCEMONOCHROMATOR:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetStringProperty(this->m_pdisp, DISPID_ActiveMono, ENTRANCE_MONO);
			return TRUE;
		}
		break;
	case IDC_RADIOEXITMONO:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetStringProperty(this->m_pdisp, DISPID_ActiveMono, EXIT_MONO); 
			return TRUE;
		}
		break;
	case IDC_SETGRATING:
		this->OnSetGrating();
		return TRUE;
	case IDC_SETPOSITION:
		this->OnSetPosition();
		return TRUE;
	case IDC_CHKAUTOGRATING:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedChkAutoGrating();
			return TRUE;
		}
		break;
	case IDC_CHKINITONSCANSTART:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedInitOnScanStart();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

void CMyPropPage::OnSetGrating()
{
	return;
	TCHAR		szMono[MAX_PATH];
	long		grating		= GetDlgItemInt(this->m_hwndDlg, IDC_EDITGRATING, NULL, TRUE);
/*
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		this->SetMonoGrating(szMono, grating);
	}
*/
//	this->SetCurrentGrating(grating);
}

void CMyPropPage::OnSetPosition()
{
	return;
	TCHAR		szMono[MAX_PATH];
	TCHAR		szPosition[MAX_PATH];
	float		fval;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szPosition, MAX_PATH))
	{
		if (1 == _stscanf_s(szPosition, TEXT("%f"), &fval))
		{
/*
			if (this->GetActiveMono(szMono, MAX_PATH))
			{
				this->SetMonoPosition(szMono, fval);
			}
*/
	//		this->SetCurrentWavelength(fval);
		}
	}
}

BOOL CMyPropPage::GetActiveMono(
								LPTSTR			szMono,
								UINT			nBufferSize)
{
	LPTSTR				szString		= NULL;
	BOOL				fSuccess		= FALSE;
	szMono[0]	= '\0';
	Utils_GetStringProperty(this->m_pdisp, DISPID_ActiveMono, &szString);
	if (NULL != szString)
	{
		fSuccess	= TRUE;
		StringCchCopy(szMono, nBufferSize, szString);
		CoTaskMemFree((LPVOID) szString);
	}
	return fSuccess;
}


BOOL CMyPropPage::OnNotify(
								LPNMHDR			pnmh)
{
	return FALSE;
}

void CMyPropPage::OnEditBoxReturnClicked(
								UINT			nEditBox)
{
	if (IDC_EDITGRATING == nEditBox)
	{
		this->OnSetGrating();
	}
	else if (IDC_EDITPOSITION == nEditBox)
	{
		this->OnSetPosition();
	}
}

// subclass an edit box
void CMyPropPage::SubclassEditBox(
								UINT			uControlID)
{
	MY_SUBCLASS_STRUCT	*	pSubclass	= new MY_SUBCLASS_STRUCT;
	HWND					hwndEdit	= GetDlgItem(this->m_hwndDlg, uControlID);

	pSubclass->pMyPropPage	= this;
	pSubclass->wpOrig		= (WNDPROC) SetWindowLongPtr(hwndEdit, GWLP_WNDPROC, (LONG) MyEditBoxSubclass);
	SetWindowLongPtr(hwndEdit, GWLP_USERDATA, (LONG) pSubclass);
}

// edit box subclass
LRESULT CALLBACK MyEditBoxSubclass(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage::MY_SUBCLASS_STRUCT*	pSubclass = (CMyPropPage::MY_SUBCLASS_STRUCT*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_GETDLGCODE:
		return DLGC_WANTALLKEYS;			// want the return key
	case WM_CHAR:
		if (0 != (0x8000 & GetKeyState(VK_RETURN)))
		{
			pSubclass->pMyPropPage->OnEditBoxReturnClicked(GetDlgCtrlID(hwnd));
			return 0;
		}
		break;
	case WM_DESTROY:
		// remove the subclass
		{
			WNDPROC			wpOrig	= pSubclass->wpOrig;
			delete pSubclass;
			SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
			SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG) wpOrig);
			return CallWindowProc(wpOrig, hwnd, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return CallWindowProc(pSubclass->wpOrig, hwnd, uMsg, wParam, lParam);
}

void CMyPropPage::DisplaySystemInit()
{
	BOOL			fSystemInit		= Utils_GetBoolProperty(this->m_pdisp, DISPID_SystemInit);
	TCHAR			szString[50];
	if (!fSystemInit)
	{
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLGRATING), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITGRATING), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETGRATING), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLENTRANCEMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITENTRANCEMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDENTRANCEMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLEXITMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITEXITMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDEXITMONOPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETPOSITION), FALSE);
	}
	else
	{
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLGRATING), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITGRATING), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETGRATING), TRUE);
		SetDlgItemInt(this->m_hwndDlg, IDC_EDITGRATING,
			Utils_GetLongProperty(this->m_pdisp, DISPID_Grating), FALSE);
		if (Utils_GetLongProperty(this->m_pdisp, DISPID_MonoDecoupled))
		{
			// mono decoupled
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETPOSITION), FALSE);

			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLENTRANCEMONOPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITENTRANCEMONOPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDENTRANCEMONOPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLEXITMONOPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITEXITMONOPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDEXITMONOPOSITION), TRUE);
			_stprintf_s(szString, 50, TEXT("%5.0f"), this->GetMonoPosition(ENTRANCE_MONO));
			SetDlgItemText(this->m_hwndDlg, IDC_EDITENTRANCEMONOPOSITION, szString);
			_stprintf_s(szString, 50, TEXT("%5.0f"), this->GetMonoPosition(EXIT_MONO));
			SetDlgItemText(this->m_hwndDlg, IDC_EDITENTRANCEMONOPOSITION, szString);
		}
		else
		{
			// monos coupled
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITPOSITION), TRUE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_SETPOSITION), TRUE);
			_stprintf_s(szString, 50, TEXT("%5.0f"), this->GetMonoPosition(ENTRANCE_MONO));
			SetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szString);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLENTRANCEMONOPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITENTRANCEMONOPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDENTRANCEMONOPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_LBLEXITMONOPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITEXITMONOPOSITION), FALSE);
			EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDEXITMONOPOSITION), FALSE);
		}
	}
}

// display mono coupled
void CMyPropPage::DisplayMonoCoupled()
{
	if (Utils_GetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled))
	{
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_FRAMEACTIVEMONO), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_RADIOENTRANCEMONOCHROMATOR), TRUE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_RADIOEXITMONO), TRUE);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOTURNOFFDECOUPLE), BM_SETCHECK,
			BST_UNCHECKED, 0);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIODECOUPLE), BM_SETCHECK,
			BST_CHECKED, 0);
		this->DisplayActiveMono();
	}
	else
	{
		// monos not decoupled
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_FRAMEACTIVEMONO), FALSE);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOENTRANCEMONOCHROMATOR),
			BM_SETCHECK, BST_CHECKED, 0);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOEXITMONO),
			BM_SETCHECK, BST_CHECKED, 0);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_RADIOENTRANCEMONOCHROMATOR), FALSE);
		EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_RADIOEXITMONO), FALSE);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOTURNOFFDECOUPLE), BM_SETCHECK,
			BST_CHECKED, 0);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIODECOUPLE), BM_SETCHECK,
			BST_UNCHECKED, 0);
	}
	this->DisplaySystemInit();
}

void CMyPropPage::DisplayActiveMono()
{
	if (Utils_GetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled))
	{
		LPTSTR			szString		= NULL;
		Utils_GetStringProperty(this->m_pdisp, DISPID_ActiveMono, &szString);
		if (NULL != szString)
		{
			SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOENTRANCEMONOCHROMATOR),
				BM_SETCHECK, BST_UNCHECKED, 0);
			SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOEXITMONO),
				BM_SETCHECK, BST_UNCHECKED, 0);
			if (0 == lstrcmpi(szString, ENTRANCE_MONO))
			{
				SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOENTRANCEMONOCHROMATOR),
					BM_SETCHECK, BST_CHECKED, 0);
			}
			else if (0 == lstrcmpi(szString, EXIT_MONO))
			{
				SendMessage(GetDlgItem(this->m_hwndDlg, IDC_RADIOEXITMONO),
					BM_SETCHECK, BST_CHECKED, 0);
			}
			CoTaskMemFree((LPVOID) szString);
		}
	}
	// display grating and position
	TCHAR			szMono[MAX_PATH];
	TCHAR			szPosition[MAX_PATH];
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		_stprintf_s(szPosition, MAX_PATH, TEXT("%6.1f"), 
			(float) this->GetMonoPosition(szMono));
		SetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szPosition);
		SetDlgItemInt(this->m_hwndDlg, IDC_EDITGRATING, this->GetMonoGrating(szMono), TRUE);
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_CHKAUTOGRATING), BM_SETCHECK,
			this->GetMonoAutoGrating(szMono) ? BST_CHECKED : BST_UNCHECKED, 0);
	}
}

// property change notifications
void CMyPropPage::OnPropChanged(
								DISPID			dispid)
{
	if (DISPID_MonoDecoupled == dispid)
	{
		this->DisplayMonoCoupled();
		this->DisplayActiveMono();
	}
	else if (DISPID_ActiveMono == dispid)
	{
		this->DisplayActiveMono();
	}
	else if (DISPID_ReInitOnScanStart == dispid)
	{
		this->DisplayReInitOnScanStart();
	}
}

// get mono position
double CMyPropPage::GetMonoPosition(
								LPCTSTR			szMono)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	double				dval		= 0.0;
	InitVariantFromString(szMono, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, DISPID_GetMonoPosition, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &dval);
	return dval;
}

// set mono position
void CMyPropPage::SetMonoPosition(
								LPCTSTR			szMono,
								double			position)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(szMono, &avarg[1]);
	InitVariantFromDouble(position, &avarg[0]);
	Utils_InvokeMethod(this->m_pdisp, DISPID_SetMonoPosition, avarg, 2, NULL);
	VariantClear(&avarg[1]);
}

void CMyPropPage::OnMonoPositionChanged(
								LPCTSTR			Mono,
								double			position)
{
	TCHAR			szMono[MAX_PATH];
	TCHAR			szPosition[MAX_PATH];
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		_stprintf_s(szPosition, MAX_PATH, TEXT("%6.1f"), (float) this->GetMonoPosition(szMono));
		SetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szPosition);
	}

/*
	// dont do anything if mono decoupled
	if (!Utils_GetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled))
	{
		TCHAR			szMono[MAX_PATH];
		if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
			StringCchCopy(szMono, MAX_PATH, EXIT_MONO);
		else if (0 == lstrcmpi(Mono, EXIT_MONO))
			StringCchCopy(szMono, MAX_PATH, ENTRANCE_MONO);
		else	
			return;
		double			diff = abs(this->GetMonoPosition(szMono) - position);
		if (diff > 1.0)
		{
			this->SetMonoPosition(szMono, position);
		}
		else
		{
			// display the current position
			TCHAR		szString[50];
			_stprintf_s(szString, 50, TEXT("%5.0f"), this->GetMonoPosition(ENTRANCE_MONO));
			SetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szString);
		}
	}
*/
}

// mono grating changed
void CMyPropPage::OnMonoGratingChanged(
								LPCTSTR			Mono,
								long			grating)
{
/*
	// dont do anything if mono decoupled
	if (!Utils_GetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled))
	{
		TCHAR			szMono[MAX_PATH];
		if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
			StringCchCopy(szMono, MAX_PATH, EXIT_MONO);
		else if (0 == lstrcmpi(Mono, EXIT_MONO))
			StringCchCopy(szMono, MAX_PATH, ENTRANCE_MONO);
		else
			return;
		if (this->GetMonoGrating(szMono) != grating)
		{
			this->SetMonoGrating(szMono, grating);
		}
		else
		{
			SetDlgItemInt(this->m_hwndDlg, IDC_EDITGRATING, grating, TRUE);
		}
	}
*/
	TCHAR			szMono[MAX_PATH];
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		SetDlgItemInt(this->m_hwndDlg, IDC_EDITGRATING, this->GetMonoGrating(szMono), TRUE);
	}
}

long CMyPropPage::GetMonoGrating(
								LPCTSTR			Mono)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	long				grating		= -1;
	InitVariantFromString(Mono, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, DISPID_GetMonoGrating, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &grating);
	return grating;
}

void CMyPropPage::SetMonoGrating(
								LPCTSTR			Mono,
								long			grating)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(Mono, &avarg[1]);
	InitVariantFromInt32(grating, &avarg[0]);
	Utils_InvokeMethod(this->m_pdisp, DISPID_SetMonoGrating, avarg, 2, NULL);
	VariantClear(&avarg[1]);
}

// auto grating prop change
void CMyPropPage::OnMonoAutoGratingChanged(
								LPCTSTR			Mono,
								BOOL			AutoGrating)
{
/*
	// dont do anything if mono decoupled
	if (!Utils_GetBoolProperty(this->m_pdisp, DISPID_MonoDecoupled))
	{
		TCHAR			szMono[MAX_PATH];
		if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
			StringCchCopy(szMono, MAX_PATH, EXIT_MONO);
		else if (0 == lstrcmpi(Mono, EXIT_MONO))
			StringCchCopy(szMono, MAX_PATH, ENTRANCE_MONO);
		else
			return;
		if (this->GetMonoAutoGrating(szMono) != AutoGrating)
		{
			this->SetMonoAutoGrating(szMono, AutoGrating);
		}
	}
	*/
	TCHAR			szMono[MAX_PATH];
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		SendMessage(GetDlgItem(this->m_hwndDlg, IDC_CHKAUTOGRATING), BM_SETCHECK,
			this->GetMonoAutoGrating(szMono) ? BST_CHECKED : BST_UNCHECKED, 0);
	}

	// display the dispersion
	this->DisplayDispersion();
}

BOOL CMyPropPage::GetMonoAutoGrating(
								LPCTSTR			szMono)
{
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				AutoGrating		= FALSE;
	InitVariantFromString(szMono, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, DISPID_GetMonoAutoGrating, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &AutoGrating);
	return AutoGrating;
}

void CMyPropPage::SetMonoAutoGrating(
								LPCTSTR			szMono,
								BOOL			AutoGrating)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(szMono, &avarg[1]);
	InitVariantFromBoolean(AutoGrating, &avarg[0]);
	Utils_InvokeMethod(this->m_pdisp, DISPID_SetMonoAutoGrating, avarg, 2, NULL);
	VariantClear(&avarg[1]);
}

void CMyPropPage::OnClickedChkAutoGrating()		// click event on auto grating control
{
	TCHAR			szMono[MAX_PATH];
	if (this->GetActiveMono(szMono, MAX_PATH))
	{
		// toggle on/off auto grating
		BOOL		fAutoGrating		= this->GetMonoAutoGrating(szMono);
		this->SetMonoAutoGrating(szMono, !fAutoGrating);
	}
}

// mono position
double CMyPropPage::GetCurrentWavelength()
{
	IDispatch	*	pdisp;
	double			currentWavelength = 0.0;
	DISPID			dispid;
	if (this->Get_clsIMono(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurrentWavelength", &dispid);
		currentWavelength = Utils_GetDoubleProperty(pdisp, dispid);
		pdisp->Release();
	}
	return currentWavelength;
}
void CMyPropPage::SetCurrentWavelength(
	double			CurrentWavelength)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	if (this->Get_clsIMono(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurrentWavelength", &dispid);
		Utils_SetDoubleProperty(pdisp, dispid, CurrentWavelength);
		pdisp->Release();
	}
}

long CMyPropPage::GetCurrentGrating()
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	long			grating = -1;
	if (this->Get_clsIMono(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurrentGrating", &dispid);
		grating = Utils_GetLongProperty(pdisp, dispid);
		pdisp->Release();
	}
	return grating;

}

void CMyPropPage::SetCurrentGrating(
	long			CurrentGrating)
{
	IDispatch	*	pdisp;
	DISPID			dispid;
	if (this->Get_clsIMono(&pdisp))
	{
		Utils_GetMemid(pdisp, L"CurrentGrating", &dispid);
		Utils_SetLongProperty(pdisp, dispid, CurrentGrating);
		pdisp->Release();
	}
}

BOOL CMyPropPage::Get_clsIMono(
	IDispatch	**	ppdisp)
{
	HRESULT				hr;
	IProvideClassInfo*	pProvideClassInfo = NULL;
	ITypeInfo		*	pClassInfo		= NULL;
	ITypeInfo		*	pTypeInfo = NULL;
	TYPEATTR		*	pTypeAttr;
	BOOL				fSuccess = FALSE;
	*ppdisp = NULL;
	hr = this->m_pdisp->QueryInterface(IID_IProvideClassInfo, (LPVOID*)&pProvideClassInfo);
	if (SUCCEEDED(hr))
	{
		hr = pProvideClassInfo->GetClassInfo(&pClassInfo);
		pProvideClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		if (!Utils_FindImplClassName(pClassInfo, L"_clsIMono", &pTypeInfo))
			hr = E_FAIL;
		pClassInfo->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			hr = this->m_pdisp->QueryInterface(pTypeAttr->guid, (LPVOID*)ppdisp);
			fSuccess = SUCCEEDED(hr);
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
	return fSuccess;
}

CMyPropPage::CImpIPropNotify::CImpIPropNotify(CMyPropPage * pMyPropPage) :
	m_pMyPropPage(pMyPropPage),
	m_cRefs(0)
{
}

CMyPropPage::CImpIPropNotify::~CImpIPropNotify()
{
}

// IUnknown methods
STDMETHODIMP CMyPropPage::CImpIPropNotify::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyNotifySink == riid)
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

STDMETHODIMP_(ULONG) CMyPropPage::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyPropPage::CImpIPropNotify::Release()
{
	ULONG			cRefs		= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IPropertyNotifySink methods
STDMETHODIMP CMyPropPage::CImpIPropNotify::OnChanged( 
									DISPID			dispID)
{
	this->m_pMyPropPage->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CMyPropPage::CImpIPropNotify::OnRequestEdit( 
									DISPID			dispID)
{
	return S_OK;
}

CMyPropPage::CImpISink::CImpISink(CMyPropPage * pMyPropPage) :
	m_pMyPropPage(pMyPropPage),
	m_cRefs(0)
{
}

CMyPropPage::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CMyPropPage::CImpISink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || MY_IIDSINK == riid)
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

STDMETHODIMP_(ULONG) CMyPropPage::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyPropPage::CImpISink::Release()
{
	ULONG		cRefs	= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CMyPropPage::CImpISink::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CMyPropPage::CImpISink::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyPropPage::CImpISink::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyPropPage::CImpISink::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	TCHAR				szMono[MAX_PATH];

	VariantInit(&varg);
	if (dispIdMember == DISPID_MonoGratingChanged)
	{
		if (this->GetMono(pDispParams, szMono, MAX_PATH))
		{
			hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
			if (SUCCEEDED(hr))
				this->m_pMyPropPage->OnMonoGratingChanged(szMono, varg.lVal);
		}
	}
	else if (dispIdMember == DISPID_MonoPositionChanged)
	{
		if (this->GetMono(pDispParams, szMono, MAX_PATH))
		{
			hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
			if (SUCCEEDED(hr))
				this->m_pMyPropPage->OnMonoPositionChanged(szMono, varg.dblVal);
		}
	}
	else if (dispIdMember == DISPID_MonoAutoGratingChanged)
	{
		if (this->GetMono(pDispParams, szMono, MAX_PATH))
		{
			hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
			if (SUCCEEDED(hr))
				this->m_pMyPropPage->OnMonoAutoGratingChanged(szMono, VARIANT_TRUE == varg.boolVal);
		}
	}
	return S_OK;
}

BOOL CMyPropPage::CImpISink::GetMono(
								DISPPARAMS	*	pDispParams,
								LPTSTR			szMono,
								UINT			nBufferSize)
{
	HRESULT			hr;
	VARIANTARG		varg;
//	LPTSTR			szString		= NULL;
//	long			grating;
	BOOL			fSuccess		= FALSE;
	UINT			uArgErr;
	szMono[0]	= '\0';
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (SUCCEEDED(hr))
	{
//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
		StringCchCopy(szMono, nBufferSize, varg.bstrVal);
		fSuccess	= TRUE;
//		CoTaskMemFree((LPVOID) szString);
		VariantClear(&varg);
	}
	return fSuccess;
}

// display the dispersion
void CMyPropPage::DisplayDispersion()
{

}

void CMyPropPage::DisplayReInitOnScanStart()
{
	BOOL		fReInitOnScanStart = Utils_GetBoolProperty(this->m_pdisp, DISPID_ReInitOnScanStart);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINITONSCANSTART), fReInitOnScanStart ? BST_CHECKED : BST_UNCHECKED);
}

void CMyPropPage::OnClickedInitOnScanStart()
{
	BOOL		fReInitOnScanStart = BST_CHECKED == Button_GetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKINITONSCANSTART));
	Utils_SetBoolProperty(this->m_pdisp, DISPID_ReInitOnScanStart, !fReInitOnScanStart);
}
