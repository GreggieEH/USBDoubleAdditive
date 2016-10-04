// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <Windowsx.h>
#include <ole2.h>
#include <olectl.h>

#include <shlobj.h>
#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <math.h>
#include <Propvarutil.h>

// CNG
#include <bcrypt.h>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

class CServer;
CServer * GetServer();

// get the property page for an object
BOOL GetObjectPropPage(
	IDispatch	*	pdisp,
	IUnknown	**	ppUnk,
	CLSID		*	pClsidPage);

// definitions
#define				MY_CLSID					CLSID_USBDoubleAdditive
#define				MY_LIBID					LIBID_USBDoubleAdditive
#define				MY_IID						IID_IUSBDoubleAdditive
#define				MY_IIDSINK					IID__USBDoubleAdditive

#define				MAX_CONN_PTS				3
#define				CONN_PT_CUSTOMSINK			0
#define				CONN_PT_PROPNOTIFY			1
#define				CONN_PT__clsIMono			2

#define				FRIENDLY_NAME				TEXT("USBDoubleAdditive")
#define				PROGID						TEXT("Sciencetech.USBDoubleAdditive.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.USBDoubleAdditive")

// mono info parameters

#define				MONO_INFO_NUMGRATINGS		100

// 
#define				ENTRANCE_MONO				TEXT("EntranceMono")
#define				EXIT_MONO					TEXT("ExitMono")

