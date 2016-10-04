#include "stdafx.h"
#include "Server.h"
#include "MyGuids.h"
#include "MyFactory.h"
#include "DateCheck.h"
#include "MyReggie.h"

CServer::CServer(HINSTANCE hInst) :
	m_hInst(hInst),
	m_cObjects(0),
	m_cLocks(0),
	m_hFile(NULL)
{
	// create a logging file
	TCHAR				szLogFile[MAX_PATH];

	GetModuleFileName(this->m_hInst, szLogFile, MAX_PATH);
	PathRemoveFileSpec(szLogFile);
	PathAppend(szLogFile, TEXT("USBDoubleAdditive.txt"));
	this->m_hFile	= CreateFile(szLogFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
						FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == this->m_hFile)
		this->m_hFile = NULL;
}

CServer::~CServer(void)
{
	if (NULL != this->m_hFile)
	{
		CloseHandle(this->m_hFile);
		this->m_hFile = NULL;
	}
}

HINSTANCE CServer::GetInstance()
{
	return this->m_hInst;
}

BOOL CServer::CanUnloadNow()
{
	return 0 == this->m_cObjects && 0 == this->m_cLocks;
}

HRESULT CServer::GetClassObject(
							REFCLSID	rclsid,
							REFIID		riid,
							LPVOID	*	ppv)
{
	CMyFactory		*	pMyFactory	= NULL;
	HRESULT				hr = CLASS_E_CLASSNOTAVAILABLE;

	*ppv = NULL;
	if (MY_CLSID == rclsid)
	{
		// main COM object
		pMyFactory = new CMyFactory();
	}
	else if (CLSID_MyPropPage == rclsid)
	{
		pMyFactory = new CMyFactory(TRUE);
	}
	else
	{
		return hr;
	}
	if (NULL != pMyFactory)
	{
		this->ObjectsUp();
		hr = pMyFactory->QueryInterface(riid, ppv);
		if (FAILED(hr))
		{
			this->ObjectsDown();
			delete pMyFactory;
		}
	}
	else
		hr = E_OUTOFMEMORY;
	return hr;
}

HRESULT CServer::RegisterServer()
{
	HRESULT			hr = NOERROR;
	TCHAR			szID[MAX_PATH];
	TCHAR			szCLSID[MAX_PATH];
	TCHAR			szModulePath[MAX_PATH];
//	WCHAR			wszModulePath[MAX_PATH];
	ITypeLib	*	pITypeLib;
//	TCHAR			szString[MAX_PATH];
//	TCHAR			szStr2[MAX_PATH];
	WCHAR			wszID[MAX_PATH];
	LPTSTR			szTypeLib		= NULL;

	// Obtain the path to this module's executable file for later use.
	GetModuleFileName(this->m_hInst, szModulePath, sizeof(szModulePath)/sizeof(TCHAR));

	// Create some base key strings.
	MyStringFromCLSID(MY_CLSID, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCat(szCLSID, MAX_PATH, szID);

	// Create ProgID keys.
	SetRegKeyValue(PROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(PROGID, TEXT("CLSID"), szID);

	// Create VersionIndependentProgID keys.
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CurVer"), PROGID);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CLSID"), szID);
	// Create entries under CLSID.
	SetRegKeyValue(szCLSID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(szCLSID, TEXT("ProgID"), PROGID);
	SetRegKeyValue(szCLSID, TEXT("VersionIndependentProgID"), VERSIONINDEPENDENTPROGID);
	SetRegKeyValue(szCLSID, TEXT("NotInsertable"), NULL);
	SetRegKeyValue(szCLSID, TEXT("InprocServer32"), szModulePath);
	// type library
	StringFromGUID2(MY_LIBID, wszID, MAX_PATH);
//	Utils_UnicodeToAnsi(wszID, &szTypeLib);
//	if (NULL != szTypeLib)
//	{
	SetRegKeyValue(szCLSID, TEXT("TypeLib"), wszID);
	//	CoTaskMemFree((LPVOID) szTypeLib);
//	}
	// version number
	SetRegKeyValue(szCLSID, TEXT("Version"), TEXT("1.0"));
	// register the type library
/*
	MultiByteToWideChar(
		CP_ACP,
		0,
		szModulePath,
		-1,
		wszModulePath,
		MAX_PATH);
*/
	hr = LoadTypeLibEx(szModulePath, REGKIND_REGISTER, &pITypeLib);
	if (SUCCEEDED(hr))
		pITypeLib->Release();

	// register the property page
	RegisterPropPage(CLSID_MyPropPage,
		TEXT("PropPageUSBDoubleAdditive"), TEXT("Sciencetech.PropPageUSBDoubleAdditive"));

	return hr;
}

HRESULT CServer::UnregisterServer()
{
	HRESULT  hr = S_OK;
	TCHAR    szID[MAX_PATH];
	TCHAR    szCLSID[MAX_PATH];
	TCHAR    szTemp[MAX_PATH];

	//Create some base key strings.
	MyStringFromCLSID(MY_CLSID, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCopy(szCLSID, MAX_PATH, szID);

	// un register the entries under Version independent prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CurVer"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, VERSIONINDEPENDENTPROGID); 

	// delete the entries under prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), PROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, PROGID);

	// delete the entries under CLSID
	StringCchPrintf(szTemp, MAX_PATH,TEXT("%s\\%s"), szCLSID, TEXT("ProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("VersionIndependentProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH,TEXT("%s\\%s"), szCLSID, TEXT("NotInsertable"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH,TEXT("%s\\%s"), szCLSID, TEXT("InprocServer32"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szCLSID);

	// unregister the type library
	hr = UnRegisterTypeLib(MY_LIBID, 1, 0, 0x09, SYS_WIN32);

	// unregister the property page
	UnregisterPropPage(CLSID_MyPropPage,
		TEXT("PropPageUSBDoubleAdditive"), TEXT("Sciencetech.PropPageUSBDoubleAdditive"));


	return S_OK;
}

void CServer::ObjectsUp()
{
	this->m_cObjects++;
}

void CServer::ObjectsDown()
{
	this->m_cObjects--;
}

void CServer::LocksUp()
{
	this->m_cLocks++;
}

void CServer::LocksDown()
{
	this->m_cLocks--;
}


void CServer::RegisterPropPage(
								REFCLSID		clsid,
								LPTSTR			szName,
								LPTSTR			szProgID)
{
	TCHAR		szID[MAX_PATH];
	WCHAR		wszID[MAX_PATH];
	TCHAR		szCLSID[MAX_PATH];
	TCHAR		szModulePath[MAX_PATH];

	// Create some base key strings.
	StringFromGUID2(clsid, wszID, MAX_PATH);
//	SHUnicodeToAnsi(wszID, szID, MAX_PATH);
	StringCchPrintf(szCLSID, MAX_PATH, TEXT("CLSID\\%s"), wszID);
	SetRegKeyValue(szCLSID, NULL, szName);
	// Obtain the path to this module's executable file for later use.
	GetModuleFileName(this->m_hInst, szModulePath, sizeof(szModulePath)/sizeof(TCHAR));
	SetRegKeyValue(szCLSID, TEXT("InprocServer32"), szModulePath);
	// Create ProgID keys.
	SetRegKeyValue(szProgID, NULL, szName);
	SetRegKeyValue(szProgID, TEXT("CLSID"), szID);
}

void CServer::UnregisterPropPage(
								REFCLSID		clsid,
								LPTSTR			szName,
								LPTSTR			szProgID)
{
	TCHAR				szTemp[MAX_PATH];
	WCHAR				wszID[MAX_PATH];
//	TCHAR				szID[MAX_PATH];
	TCHAR				szCLSID[MAX_PATH];

	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), szProgID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szProgID);
	StringFromGUID2(clsid, wszID, MAX_PATH);
//	SHUnicodeToAnsi(wszID, szID, MAX_PATH);
	StringCchPrintf(szCLSID, MAX_PATH, TEXT("CLSID\\%s"), wszID);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\InprocServer32"), szCLSID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szCLSID);
}


/*F+F++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: SetRegKeyValue

  Summary:  Internal utility function to set a Key, Subkey, and value
            in the system Registry under HKEY_CLASSES_ROOT.

  Args:     LPTSTR pszKey,
            LPTSTR pszSubkey,
            LPTSTR pszValue)

  Returns:  BOOL
              TRUE if success; FALSE if not.
------------------------------------------------------------------------F-F*/
BOOL CServer::SetRegKeyValue(
       LPTSTR pszKey,
       LPTSTR pszSubkey,
       LPTSTR pszValue)
{
  BOOL bOk = FALSE;
  LONG ec;
  HKEY hKey;
  TCHAR szKey[MAX_PATH];

  StringCchCopy(szKey, MAX_PATH, pszKey);

	if (NULL != pszSubkey)
	{
		StringCchCat(szKey, MAX_PATH, TEXT("\\"));
		StringCchCat(szKey, MAX_PATH, pszSubkey);
	}

	ec = RegCreateKeyEx(
         HKEY_CLASSES_ROOT,
         szKey,
         0,
         NULL,
         REG_OPTION_NON_VOLATILE,
         KEY_ALL_ACCESS,
         NULL,
         &hKey,
         NULL);

  if (ERROR_SUCCESS == ec)
  {
    if (NULL != pszValue)
    {
      ec = RegSetValueEx(
             hKey,
             NULL,
             0,
             REG_SZ,
             (BYTE *)pszValue,
             (lstrlen(pszValue)+1)*sizeof(TCHAR));
    }
    if (ERROR_SUCCESS == ec)
      bOk = TRUE;
    RegCloseKey(hKey);
  }
  return bOk;
}

HRESULT CServer::MyStringFromCLSID(
							    REFGUID		refGuid,
								LPTSTR		szCLSID)
{
	HRESULT				hr;
	LPOLESTR			wszCLSID;
//	LPTSTR				szTemp;

	hr = StringFromCLSID(refGuid, &wszCLSID);
	if (SUCCEEDED(hr))
	{
		// copy to output string
//		Utils_UnicodeToAnsi(wszCLSID, &szTemp);
		StringCchCopy(szCLSID, MAX_PATH, wszCLSID);
//		CoTaskMemFree((LPVOID) szTemp);
		CoTaskMemFree((LPVOID) wszCLSID);
	}
	return hr;
}

HRESULT CServer::GetTypeLib(
						ITypeLib	**	ppTypeLib)
{
	HRESULT				hr;

	*ppTypeLib = NULL;
	hr = LoadRegTypeLib(MY_LIBID, 1, 0, LOCALE_SYSTEM_DEFAULT, ppTypeLib);
	if (FAILED(hr))
	{
		TCHAR				szModule[MAX_PATH];
//		LPOLESTR			wszModule		= NULL;

		GetModuleFileName((HMODULE) (this->GetInstance()), szModule, MAX_PATH);
//		Utils_AnsiToUnicode(szModule, &wszModule);
		hr = LoadTypeLibEx(szModule, REGKIND_REGISTER, ppTypeLib);
//		CoTaskMemFree((LPVOID) wszModule);
	}
	return hr;
}

void CServer::DoLogString(
							LPCTSTR		szString)
{
	TCHAR				szLine[MAX_PATH];
	size_t				slen;
	DWORD				dwNWritten;

	if (NULL != this->m_hFile)
	{
		StringCchPrintf(szLine, MAX_PATH, TEXT("%s\r\n"), szString);
		StringCbLength(szLine, MAX_PATH * sizeof(TCHAR), &slen);
		WriteFile(this->m_hFile, (LPCVOID) szLine, slen, &dwNWritten, NULL);
	}
}

// check for licence
BOOL CServer::CheckLicence()
{
	// load string from resource
	WCHAR			wszString[MAX_PATH];
	WCHAR			szLicense[MAX_PATH];
	TCHAR			szString[MAX_PATH];
	CMyReggie		reg;

	// form licence key
	StringFromGUID2(MY_LIC_KEY, szLicense, MAX_PATH);
	// check for resource string
	if (LoadString(this->GetInstance(), IDS_NICESTRING, wszString, MAX_PATH) > 0)
	{
		if (0 == lstrcmpi(szLicense, wszString))
		{
			return TRUE;
		}
	}
	// check windows registry for key value
	if (reg.GetStringValue(L"val2", szString, MAX_PATH))
	{
		if (0 == lstrcmpi(szLicense, szString))
		{
			return TRUE;
		}
	}
	// resource string license not found, check in windows registry
	CDateCheck			dateCheck;
	return dateCheck.MyDateCheck();
}
