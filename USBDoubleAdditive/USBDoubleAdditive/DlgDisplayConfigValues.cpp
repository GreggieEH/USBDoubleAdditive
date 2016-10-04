#include "stdafx.h"
#include "DlgDisplayConfigValues.h"
#include "MyObject.h"
#include "dispids.h"

CDlgDisplayConfigValues::CDlgDisplayConfigValues(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_hwndDlg(NULL)
{
}

CDlgDisplayConfigValues::~CDlgDisplayConfigValues()
{
}

void CDlgDisplayConfigValues::DoOpenDialog(
	HINSTANCE		hInst,
	HWND			hwndParent)
{
	DialogBoxParam(hInst, MAKEINTRESOURCE(IDD_DIALOGDisplayConfigValues), hwndParent, (DLGPROC)DlgProcDisplayConfigValues, (LPARAM)this);
}

LRESULT CALLBACK DlgProcDisplayConfigValues(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgDisplayConfigValues*	pDlg = NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgDisplayConfigValues*)lParam;
		pDlg->m_hwndDlg = hwndDlg;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
	}
	else
	{
		pDlg = (CDlgDisplayConfigValues*)GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
	{
		switch (uMsg)
		{
		case WM_INITDIALOG:
			return pDlg->OnInitDialog();
		case WM_COMMAND:
			return pDlg->OnCommand(LOWORD(wParam), HIWORD(wParam));
		default:
			break;
		}
	}
	return FALSE;
}

BOOL CDlgDisplayConfigValues::OnInitDialog()
{
	Utils_CenterWindow(this->m_hwndDlg);
	return TRUE;
}

BOOL CDlgDisplayConfigValues::OnCommand(
	WORD			wmID,
	WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	case IDC_ENTRANCEMONO:
		this->DisplayMonoConfigValues(ENTRANCE_MONO);
		return TRUE;
	case IDC_EXITMONO:
		this->DisplayMonoConfigValues(EXIT_MONO);
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

void CDlgDisplayConfigValues::DisplayMonoConfigValues(
	LPCTSTR			szMono)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANTARG		varg;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		InitVariantFromString(szMono, &varg);
		Utils_InvokeMethod(pdisp, DISPID_DisplayMonoConfigValues, &varg, 1, NULL);
		VariantClear(&varg);
		pdisp->Release();
	}
}