// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "Server.h"
#include <initguid.h>
#include "MyGuids.h"

CServer * g_pServer = NULL;

BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD fdwReason,
	LPVOID lpvReserved
)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		g_pServer = new CServer(hinstDLL);
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pServer);
		break;
	default:
		break;
	}
	return TRUE;
}

CServer * GetServer()
{
	return g_pServer;
}

// get the property page for an object
BOOL GetObjectPropPage(
	IDispatch	*	pdisp,
	IUnknown	**	ppUnk,
	CLSID		*	pClsidPage)
{
	HRESULT					hr;
	ISpecifyPropertyPages*	pSpecifyPropertyPages;
	CAUUID					pages;
	BOOL					fSuccess = FALSE;
	*ppUnk = NULL;
	*pClsidPage = CLSID_NULL;
	if (NULL == pdisp) return FALSE;
	hr = pdisp->QueryInterface(IID_ISpecifyPropertyPages, (LPVOID*)&pSpecifyPropertyPages);
	if (SUCCEEDED(hr))
	{
		hr = pSpecifyPropertyPages->GetPages(&pages);
		if (SUCCEEDED(hr))
		{
			if (1 == pages.cElems)
			{
				*pClsidPage = pages.pElems[0];
				hr = pSpecifyPropertyPages->QueryInterface(IID_IUnknown, (LPVOID*)ppUnk);
				fSuccess = SUCCEEDED(hr);
			}
			CoTaskMemFree((LPVOID)pages.pElems);
		}
		pSpecifyPropertyPages->Release();
	}
	return fSuccess;
}
