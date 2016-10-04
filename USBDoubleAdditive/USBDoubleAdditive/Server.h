#pragma once

class CServer
{
public:
	CServer(HINSTANCE hInst);
	~CServer(void);
	HINSTANCE			GetInstance();
	BOOL				CanUnloadNow();
	HRESULT				GetClassObject(
							REFCLSID	rclsid,
							REFIID		riid,
							LPVOID	*	ppv);
	HRESULT				RegisterServer();
	HRESULT				UnregisterServer();
	void				ObjectsUp();
	void				ObjectsDown();
	void				LocksUp();
	void				LocksDown();
	HRESULT				GetTypeLib(
							ITypeLib**	ppTypeLib);
	void				DoLogString(
							LPCTSTR		szLogString);
	// check for licence
	BOOL				CheckLicence();
protected:
	BOOL				SetRegKeyValue(
							LPTSTR		pszKey,
							LPTSTR		pszSubkey,
							LPTSTR		pszValue);
	HRESULT				MyStringFromCLSID(
							REFGUID		refGuid,
							LPTSTR		szCLSID);
	void				RegisterPropPage(
							REFCLSID	clsid,
							LPTSTR		szName,
							LPTSTR		szProgID);
	void				UnregisterPropPage(
							REFCLSID	clsid,
							LPTSTR		szName,
							LPTSTR		szProgID);
private:
	HINSTANCE			m_hInst;
	ULONG				m_cObjects;
	ULONG				m_cLocks;
	HANDLE				m_hFile;
};
