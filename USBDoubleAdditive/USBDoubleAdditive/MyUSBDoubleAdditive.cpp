#include "stdafx.h"
#include "MyUSBDoubleAdditive.h"
#include "MyObject.h"
#include "UsbMono.h"

// position offsets for mono info
#define			STREAM_SIZE						750
#define			SECTION_SIZE					250
#define			POSITION_ENTRANCEMONO			250
#define			POSITION_EXITMONO				500

CMyUSBDoubleAdditive::CMyUSBDoubleAdditive(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pEntranceMono(NULL),
	m_pExitMono(NULL),
	m_fMonoDecoupled(FALSE),
	m_fDirty(FALSE),
	m_fAmInitializing(FALSE),			// am initializing flag
	m_ReInitOnScanStart(TRUE)			// flag reinitialization on scan start
{
	// default is entrance mono
	StringCchCopy(this->m_szActiveMono, MAX_PATH, ENTRANCE_MONO);
	this->m_pEntranceMono	= new CUsbMono(this, ENTRANCE_MONO);
	this->m_pEntranceMono->doInit();
	this->m_pExitMono		= new CUsbMono(this, EXIT_MONO);
	this->m_pExitMono->doInit();
}

CMyUSBDoubleAdditive::~CMyUSBDoubleAdditive(void)
{
	Utils_DELETE_POINTER(this->m_pEntranceMono);
	Utils_DELETE_POINTER(this->m_pExitMono);
}

// persistence
BOOL CMyUSBDoubleAdditive::GetDirty()
{
	if (this->m_fDirty) return TRUE;
	if (this->m_pEntranceMono->GetDirty())
		return TRUE;
	if (this->m_pExitMono->GetDirty())
		return TRUE;
	return FALSE;
}

HRESULT CMyUSBDoubleAdditive::Load(
									LPSTREAM			pStm)
{
	HRESULT			hr;
	BYTE			arr[STREAM_SIZE];
	BOOL			bVal;
	TCHAR			szString[MAX_PATH];
	IStream		*	pStream;
	UINT			uStart;

	hr = IStream_Read(pStm, (void*) arr, STREAM_SIZE);
	if (FAILED(hr)) return hr;
	CopyMemory((PVOID) &bVal, arr, sizeof(BOOL));
	this->SetMonoDecoupled(bVal);
	uStart = sizeof(BOOL);
	CopyMemory((PVOID) szString, &(arr[uStart]), MAX_PATH);
	this->SetActiveMono(szString);
	// entrance mono
	pStream		= SHCreateMemStream(&(arr[POSITION_ENTRANCEMONO]), SECTION_SIZE);
	if (NULL != pStream)
	{
		hr = this->m_pEntranceMono->Load(pStream);
		pStream->Release();
	}
	// exit mono
	pStream		= SHCreateMemStream(&(arr[POSITION_EXITMONO]), SECTION_SIZE);
	if (NULL != pStream)
	{
		hr = this->m_pExitMono->Load(pStream);
		pStream->Release();
	}
	return hr;
}

HRESULT CMyUSBDoubleAdditive::Save(
									LPSTREAM			pStm ,  
									BOOL				fClearDirty)
{
	HRESULT			hr;
	BYTE			arr[STREAM_SIZE];
	size_t			slen;
	IStream		*	pStream;			// temp stream
	BOOL			fMonoCoupled		= this->GetMonoDecoupled();
	UINT			uStart;

	ZeroMemory((PVOID) arr, STREAM_SIZE);
	CopyMemory((PVOID) arr, (const void*) &fMonoCoupled, sizeof(BOOL));
	StringCbLength(this->m_szActiveMono, MAX_PATH * sizeof(TCHAR), &slen);
	uStart = sizeof(BOOL);
	CopyMemory((PVOID) &(arr[uStart]), (const void*) this->m_szActiveMono, slen);
	pStream		= SHCreateMemStream(&(arr[POSITION_ENTRANCEMONO]), SECTION_SIZE);
	if (NULL != pStream)
	{
		hr = this->m_pEntranceMono->Save(pStream, fClearDirty);
		pStream->Release();
	}
	pStream		= SHCreateMemStream(&(arr[POSITION_EXITMONO]), SECTION_SIZE);
	if (NULL != pStream)
	{
		hr = this->m_pExitMono->Save(pStream, fClearDirty);
		pStream->Release();
	}
	hr = IStream_Write(pStm, (const void*) arr, STREAM_SIZE);
	if (SUCCEEDED(hr) && fClearDirty) this->m_fDirty	= FALSE;
	return hr;
}

long CMyUSBDoubleAdditive::GetSizeMax()
{
	return 750;
}

HRESULT CMyUSBDoubleAdditive::InitNew()
{
	this->m_pEntranceMono->InitNew();
	this->m_pExitMono->InitNew();
	return S_OK;
}

CUsbMono* CMyUSBDoubleAdditive::GetActiveMono()
{
	CUsbMono * pActiveMono	= this->m_pEntranceMono;
	if (0 == lstrcmpi(this->m_szActiveMono, EXIT_MONO))
		pActiveMono	= this->m_pExitMono;
	return pActiveMono;
}

//
BOOL CMyUSBDoubleAdditive::GetMonoDecoupled()
{
	return this->m_fMonoDecoupled;
}

void CMyUSBDoubleAdditive::SetMonoDecoupled(
							BOOL			fMonoDecoupled)
{
	this->m_fMonoDecoupled	= fMonoDecoupled;
	if (this->m_fMonoDecoupled)
	{
		this->SetActiveMono(ENTRANCE_MONO);
	}
	else
	{
		// make sure exit mono has same position as entrance mono
		this->m_pExitMono->SetCurrentGrating(this->m_pEntranceMono->GetCurrentGrating());
		this->m_pExitMono->SetCurrentWavelength(this->m_pEntranceMono->GetCurrentWavelength());
		this->m_pExitMono->SetAutoGrating(this->m_pEntranceMono->GetAutoGrating());
	}
	this->m_fDirty	= TRUE;
}

BOOL CMyUSBDoubleAdditive::GetActiveMono(
							LPTSTR			szMono,
							UINT			nBufferSize)
{
	StringCchCopy(szMono, nBufferSize, this->m_szActiveMono);
	return TRUE;
}

void CMyUSBDoubleAdditive::SetActiveMono(
							LPCTSTR			szMono)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO) ||
		0 == lstrcmpi(szMono, EXIT_MONO))
	{
		StringCchCopy(this->m_szActiveMono, MAX_PATH, szMono);
		this->m_fDirty		= TRUE;
	}
}

BOOL CMyUSBDoubleAdditive::GetSystemInit()
{
	return this->GetAmInitialized();
}

void CMyUSBDoubleAdditive::SetSystemInit(
							BOOL			fSystemInit)
{
	this->SetAmInitialized(fSystemInit);
}

void CMyUSBDoubleAdditive::DisplayMonoSetup(
							LPCTSTR			Mono)
{
	if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
	{
		this->m_pEntranceMono->DoSetup();
	}
	else if (0 == lstrcmpi(Mono, EXIT_MONO))
	{
		this->m_pExitMono->DoSetup();
	}
}

BOOL CMyUSBDoubleAdditive::SetMonoPosition(
							LPCTSTR			Mono,
							double			position)
{
	if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
	{
		this->m_pEntranceMono->SetCurrentWavelength(position);
	}
	else if (0 == lstrcmpi(Mono, EXIT_MONO))
	{
		this->m_pExitMono->SetCurrentWavelength(position);
	}
	return TRUE;
}

double CMyUSBDoubleAdditive::GetMonoPosition(
							LPCTSTR			Mono)
{
	if (0 == lstrcmpi(Mono, ENTRANCE_MONO))
	{
		return this->m_pEntranceMono->GetCurrentWavelength();
	}
	else if (0 == lstrcmpi(Mono, EXIT_MONO))
	{
		return this->m_pExitMono->GetCurrentWavelength();
	}
	return 0.0;
}

void CMyUSBDoubleAdditive::DisplayMonoConfigValues(
							LPCTSTR			szMono)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO))
	{
		this->m_pEntranceMono->DisplayConfigValues();
	}
	else if (0 == lstrcmpi(szMono, EXIT_MONO))
	{
		this->m_pExitMono->DisplayConfigValues();
	}
}

long CMyUSBDoubleAdditive::GetMonoGrating(
							LPCTSTR			szMono)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO))
	{
		return this->m_pEntranceMono->GetCurrentGrating();
	}
	else if (0 == lstrcmpi(szMono, EXIT_MONO))
	{
		return this->m_pExitMono->GetCurrentGrating();
	}
	else
	{
		return -1;
	}
}

void CMyUSBDoubleAdditive::SetMonoGrating(
							LPCTSTR			szMono,
							long			grating)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO))
	{
		this->m_pEntranceMono->SetCurrentGrating(grating);
	}
	else if (0 == lstrcmpi(szMono, EXIT_MONO))
	{
		this->m_pExitMono->SetCurrentGrating(grating);
	}
}

BOOL CMyUSBDoubleAdditive::GetMonoAutoGrating(
							LPCTSTR			szMono)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO))
	{
		return this->m_pEntranceMono->GetAutoGrating();
	}
	else if (0 == lstrcmpi(szMono, EXIT_MONO))
	{
		return this->m_pExitMono->GetAutoGrating();
	}
	else
	{
		return FALSE;
	}
}

void CMyUSBDoubleAdditive::SetMonoAutoGrating(
							LPCTSTR			szMono,
							BOOL			AutoGrating)
{
	if (0 == lstrcmpi(szMono, ENTRANCE_MONO))
	{
		this->m_pEntranceMono->SetAutoGrating(AutoGrating);
	}
	else if (0 == lstrcmpi(szMono, EXIT_MONO))
	{
		this->m_pExitMono->SetAutoGrating(AutoGrating);
	}
}

BOOL CMyUSBDoubleAdditive::GetReInitOnScanStart()
{
	return this->m_ReInitOnScanStart;
}

void CMyUSBDoubleAdditive::SetReInitOnScanStart(
	BOOL		ReInitOnScanStart)
{
	this->m_ReInitOnScanStart = ReInitOnScanStart;
}


// IMono properties and methods
double CMyUSBDoubleAdditive::GetCurrentWavelength()
{
	if (this->GetMonoDecoupled())
	{
		return this->GetActiveMono()->GetCurrentWavelength();
	}
	else
		return this->m_pEntranceMono->GetCurrentWavelength();
}

void CMyUSBDoubleAdditive::SetCurrentWavelength(
							double			CurrentWavelength)
{
	if (this->GetMonoDecoupled())
	{
		this->GetActiveMono()->SetCurrentWavelength(CurrentWavelength);
	}
	else
	{
		this->m_pEntranceMono->SetCurrentWavelength(CurrentWavelength);
		this->m_pExitMono->SetCurrentWavelength(CurrentWavelength);
	}
	this->m_pMyObject->FireMoveCompleted(CurrentWavelength);
}

BOOL CMyUSBDoubleAdditive::GetAutoGrating()
{
	if (this->GetMonoDecoupled())
	{
		return this->GetActiveMono()->GetAutoGrating();
	}
	else
	{
		return this->m_pEntranceMono->GetAutoGrating();
	}
}

void CMyUSBDoubleAdditive::SetAutoGrating(
							BOOL			AutoGrating)
{
	if (this->GetMonoDecoupled())
	{
		this->GetActiveMono()->SetAutoGrating(AutoGrating);
	}
	else
	{
		this->m_pEntranceMono->SetAutoGrating(AutoGrating);
	}
}

long CMyUSBDoubleAdditive::GetCurrentGrating()
{
	if (this->GetMonoDecoupled())
	{
		return this->GetActiveMono()->GetCurrentGrating();
	}
	else
	{
		return this->m_pEntranceMono->GetCurrentGrating();
	}
}

void CMyUSBDoubleAdditive::SetCurrentGrating(
							long			CurrentGrating)
{
	if (this->GetMonoDecoupled())
	{
		this->GetActiveMono()->SetCurrentGrating(CurrentGrating);
	}
	else
	{
		this->m_pEntranceMono->SetCurrentGrating(CurrentGrating);
		this->m_pExitMono->SetCurrentGrating(CurrentGrating);
	}
}

BOOL CMyUSBDoubleAdditive::GetAmBusy()
{
	if (this->m_pEntranceMono->GetAmBusy())
		return TRUE;
	if (this->m_pExitMono->GetAmBusy())
		return TRUE;
	return FALSE;
}

BOOL CMyUSBDoubleAdditive::GetAmInitialized()
{
	if (!this->m_pEntranceMono->GetAmInitialized())
		return FALSE;
	if (!this->m_pExitMono->GetAmInitialized())
		return FALSE;
	return TRUE;
}

void CMyUSBDoubleAdditive::SetAmInitialized(
							BOOL			AmInitialized)
{
	this->m_fAmInitializing = TRUE;
	this->m_pEntranceMono->SetAmInitialized(AmInitialized);
	this->m_pExitMono->SetAmInitialized(AmInitialized);
	this->m_fAmInitializing = FALSE;
	// set default initial state
	this->SetAutoGrating(TRUE);
	this->SetCurrentWavelength(50.0);
}

void CMyUSBDoubleAdditive::GetGratingParams(
							long			grating, 
							short			paramIndex, 
							VARIANT		*	Value)
{
	VariantInit(Value);
//	if (this->GetMonoDecoupled())
//	{
//		this->GetActiveMono()->GetGratingParams(grating, paramIndex, Value);
//	}
//	else
//	{
		this->m_pEntranceMono->GetGratingParams(grating, paramIndex, Value);
//	}
}

void CMyUSBDoubleAdditive::SetGratingParams(
							long			grating, 
							short			paramIndex, 
							VARIANT		*	Value)
{
//	if (this->GetMonoDecoupled())
//	{
//		this->GetActiveMono()->SetGratingParams(grating, paramIndex, Value);
//	}
//	else
//	{
		this->m_pEntranceMono->SetGratingParams(grating, paramIndex, Value);
		this->m_pExitMono->SetGratingParams(grating, paramIndex, Value);
//	}
}

void CMyUSBDoubleAdditive::GetMonoParams(
							short			paramIndex, 
							VARIANT		*	Value)
{
	VariantInit(Value);
//	if (this->GetMonoDecoupled())
//	{
//		this->GetActiveMono()->GetMonoParams(paramIndex, Value);
//	}
//	else
//	{
		this->m_pEntranceMono->GetMonoParams(paramIndex, Value);
//	}
}

void CMyUSBDoubleAdditive::SetMonoParams(
							short			paramIndex, 
							VARIANT		*	Value)
{
//	if (this->GetMonoDecoupled())
//	{
//		this->GetActiveMono()->SetMonoParams(paramIndex, Value);
//	}
//	else
//	{
		this->m_pEntranceMono->SetMonoParams(paramIndex, Value);
		this->m_pExitMono->SetMonoParams(paramIndex, Value);
//	}
}

BOOL CMyUSBDoubleAdditive::IsValidPosition(
							double			position)
{
//	if (this->GetMonoDecoupled())
//	{
//		return this->GetActiveMono()->IsValidPosition(position);
//	}
//	else
//	{
		return this->m_pEntranceMono->IsValidPosition(position);
//	}
//	return FALSE;
}

BOOL CMyUSBDoubleAdditive::ConvertStepsToNM(
							BOOL			fStepsToNM, 
							long			gratingID, 
							long		*	steps, 
							double		*	nm)
{
	if (this->GetMonoDecoupled())
	{
		return this->GetActiveMono()->ConvertStepsToNM(fStepsToNM, gratingID, steps, nm);
	}
	else
		return this->m_pEntranceMono->ConvertStepsToNM(fStepsToNM, gratingID, steps, nm);
}

void CMyUSBDoubleAdditive::DisplayConfigValues()
{

}

void CMyUSBDoubleAdditive::DoSetup()
{
	IDispatch			*	pdisp;
	long					nPages			= 3;		// number of property pages
	IUnknown			**	ppUnk			= NULL;		// array of property pages
	CLSID				*	pClsid			= NULL;		// array of class ids
	HRESULT					hr;
	long					i;
	LPOLESTR				Caption			= NULL;		// caption for the property frame

	// allocate space
	ppUnk	= new IUnknown * [nPages];
	pClsid	= new CLSID [nPages];
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		::GetObjectPropPage(pdisp, &(ppUnk[0]), &(pClsid[0]));
		pdisp->Release();
	}
	this->m_pEntranceMono->GetPropPage(&(ppUnk[1]), &(pClsid[1]));
	this->m_pExitMono->GetPropPage(&(ppUnk[2]), &(pClsid[2]));
	// caption for the frame
	SHStrDup(TEXT("Instrument Properties"), &Caption);
	// create the property frame
	hr = OleCreatePropertyFrame(this->m_pMyObject->FireRequestParentWindow(), 
		0, 0, Caption, nPages,
		ppUnk, nPages, pClsid, LOCALE_USER_DEFAULT, 0, NULL);
	// cleanup
	for (i=0; i<nPages; i++)
		Utils_RELEASE_INTERFACE(ppUnk[i]);
	delete [] ppUnk;
	delete [] pClsid;
}

BOOL CMyUSBDoubleAdditive::ReadConfig(
							LPCTSTR			configfile)
{
	if (!this->m_pEntranceMono->ReadConfig(configfile))
		return FALSE;
	TCHAR			szExitMono[MAX_PATH];
	StringCchCopy(szExitMono, MAX_PATH, configfile);
	PathRemoveFileSpec(szExitMono);
	PathAppend(szExitMono, TEXT("ExitMono.cfg"));
	if (!this->m_pExitMono->ReadConfig(szExitMono))
		return FALSE;
	return TRUE;
}

void CMyUSBDoubleAdditive::WriteConfig(
							LPCTSTR			configfile)
{
	this->m_pEntranceMono->WriteConfig(configfile);
	TCHAR			szExitMono[MAX_PATH];
	StringCchCopy(szExitMono, MAX_PATH, configfile);
	PathRemoveFileSpec(szExitMono);
	PathAppend(szExitMono, TEXT("ExitMono.cfg"));
	this->m_pExitMono->WriteConfig(szExitMono);
}

double CMyUSBDoubleAdditive::GetGratingDispersion(
	long			gratingID)
{
	// get the grating pitch
	VARIANT			Value;
	long			pitch = 1200;
	this->GetGratingParams(gratingID, 0, &Value);
	VariantToInt32(Value, &pitch);
	if (0 == pitch) pitch = 1200;
	return 2.0 * 1200 / pitch;
}

// scan start actions
void CMyUSBDoubleAdditive::ScanStart()
{
	if (this->GetReInitOnScanStart() && !this->m_fAmInitializing)
	{
		this->m_fAmInitializing = TRUE;
		this->m_pEntranceMono->SetAmInitialized(TRUE);
		this->m_pExitMono->SetAmInitialized(TRUE);
		this->m_fAmInitializing = FALSE;
	}
}

// mono sink events
void CMyUSBDoubleAdditive::OnMonoGratingChanged(
							LPCTSTR			szMono,
							long			grating)
{
	if (this->m_fAmInitializing) return;
	this->m_pMyObject->FireMonoGratingChanged(szMono, grating);
}

void CMyUSBDoubleAdditive::OnMonoMoveCompleted(
							LPCTSTR			szMono,
							double			newPosition)
{
	if (this->m_fAmInitializing) return;
	this->m_pMyObject->FireMonoPositionChanged(szMono, newPosition);
}

void CMyUSBDoubleAdditive::OnMonoAutoGratingPropChange(
							LPCTSTR			szMono,
							BOOL			AutoGrating)
{
	if (this->m_fAmInitializing) return;
	this->m_pMyObject->FireMonoAutoGratingChanged(szMono, AutoGrating);
	// if the monos are coupled make sure both flags are set the same
	if (!this->GetMonoDecoupled())
	{
		if (0 == lstrcmpi(ENTRANCE_MONO, szMono))
		{
			if (this->m_pExitMono->GetAutoGrating() != AutoGrating)
				this->m_pExitMono->SetAutoGrating(AutoGrating);
		}
		else if (0 == lstrcmpi(EXIT_MONO, szMono))
		{
			if (this->m_pEntranceMono->GetAutoGrating() != AutoGrating)
				this->m_pEntranceMono->SetAutoGrating(AutoGrating);
		}
	}
}
