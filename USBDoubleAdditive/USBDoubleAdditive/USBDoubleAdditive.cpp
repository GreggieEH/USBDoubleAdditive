// USBDoubleAdditive.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "Server.h"

STDAPI DllCanUnloadNow(void)
{
	return GetServer()->CanUnloadNow() ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(
	REFCLSID rclsid,
	REFIID riid,
	LPVOID * ppv)
{
	return GetServer()->GetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	return GetServer()->RegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	return GetServer()->UnregisterServer();
}


