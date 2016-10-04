#pragma once

class CMyObject;

class CDlgDisplayConfigValues
{
public:
	CDlgDisplayConfigValues(CMyObject * pMyObject);
	~CDlgDisplayConfigValues();
	void			DoOpenDialog(
						HINSTANCE		hInst,
						HWND			hwndParent);
protected:
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	void			DisplayMonoConfigValues(
						LPCTSTR			szMono);
private:
	CMyObject		*	m_pMyObject;
	HWND				m_hwndDlg;

	friend LRESULT CALLBACK DlgProcDisplayConfigValues(HWND, UINT, WPARAM, LPARAM);
};

