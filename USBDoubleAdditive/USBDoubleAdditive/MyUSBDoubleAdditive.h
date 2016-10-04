#pragma once

class CMyObject;
class CUsbMono;

class CMyUSBDoubleAdditive
{
public:
	CMyUSBDoubleAdditive(CMyObject * pMyObject);
	~CMyUSBDoubleAdditive(void);
	// persistence
	BOOL				GetDirty();
	HRESULT				Load(
							LPSTREAM		pStm);
	HRESULT				Save(
							LPSTREAM		pStm ,  
							BOOL			fClearDirty);
	long				GetSizeMax();
	HRESULT				InitNew();
	//
	BOOL				GetMonoDecoupled();
	void				SetMonoDecoupled(
							BOOL			fMonoDecoupled);
	BOOL				GetActiveMono(
							LPTSTR			szMono,
							UINT			nBufferSize);
	void				SetActiveMono(
							LPCTSTR			szMono);
	BOOL				GetSystemInit();
	void				SetSystemInit(
							BOOL			fSystemInit);
	void				DisplayMonoSetup(
							LPCTSTR			Mono);
	BOOL				SetMonoPosition(
							LPCTSTR			Mono,
							double			position);
	double				GetMonoPosition(
							LPCTSTR			Mono);
	void				DisplayMonoConfigValues(
							LPCTSTR			szMono);
	long				GetMonoGrating(
							LPCTSTR			szMono);
	void				SetMonoGrating(
							LPCTSTR			szMono,
							long			grating);
	BOOL				GetMonoAutoGrating(
							LPCTSTR			szMono);
	void				SetMonoAutoGrating(
							LPCTSTR			szMono,
							BOOL			AutoGrating);
	BOOL				GetReInitOnScanStart();
	void				SetReInitOnScanStart(
							BOOL		ReInitOnScanStart);
	// IMono properties and methods
	double				GetCurrentWavelength();
	void				SetCurrentWavelength(
							double			CurrentWavelength);
	BOOL				GetAutoGrating();
	void				SetAutoGrating(
							BOOL			AutoGrating);
	long				GetCurrentGrating();
	void				SetCurrentGrating(
							long			CurrentGrating);
	BOOL				GetAmBusy();
	BOOL				GetAmInitialized();
	void				SetAmInitialized(
							BOOL			AmInitialized);
	void				GetGratingParams(
							long			grating, 
							short			paramIndex, 
							VARIANT		*	Value);
	void				SetGratingParams(
							long			grating, 
							short			paramIndex, 
							VARIANT		*	Value);
	void				GetMonoParams(
							short			paramIndex, 
							VARIANT		*	Value);
	void				SetMonoParams(
							short			paramIndex, 
							VARIANT		*	Value);
	BOOL				IsValidPosition(
							double			position);
	BOOL				ConvertStepsToNM(
							BOOL			fStepsToNM, 
							long			gratingID, 
							long		*	steps, 
							double		*	nm);
	void				DisplayConfigValues();
	void				DoSetup();
	BOOL				ReadConfig(
							LPCTSTR			configfile);
	void				WriteConfig(
							LPCTSTR			configfile);
	double				GetGratingDispersion(
							long			gratingID);
	// scan start actions
	void				ScanStart();
	// mono sink events
	void				OnMonoGratingChanged(
							LPCTSTR			szMono,
							long			grating);
	void				OnMonoMoveCompleted(
							LPCTSTR			szMono,
							double			newPosition);
	void				OnMonoAutoGratingPropChange(
							LPCTSTR			szMono,
							BOOL			AutoGrating);
protected:
	CUsbMono		*	GetActiveMono();


private:
	CMyObject		*	m_pMyObject;
	CUsbMono		*	m_pEntranceMono;
	CUsbMono		*	m_pExitMono;
	BOOL				m_fMonoDecoupled;
	TCHAR				m_szActiveMono[MAX_PATH];
	BOOL				m_fDirty;
	BOOL				m_fAmInitializing;
	BOOL				m_ReInitOnScanStart;			// flag reinitialization on scan start
};
