// dispids.h
// dispatch ids

#pragma once

// properties
#define				DISPID_MonoDecoupled			0x0100
#define				DISPID_ActiveMono				0x0101
#define				DISPID_SystemInit				0x0102
#define				DISPID_Grating					0x0103
#define				DISPID_Dispersion				0x0104
#define				DISPID_ReInitOnScanStart		0x0105

// methods
#define				DISPID_DisplayMonoSetup			0x0120
#define				DISPID_SetMonoPosition			0x0122
#define				DISPID_GetMonoPosition			0x0123
#define				DISPID_DisplayMonoConfigValues	0x0124
#define				DISPID_GetMonoGrating			0x0125
#define				DISPID_SetMonoGrating			0x0126
#define				DISPID_GetMonoAutoGrating		0x0127
#define				DISPID_SetMonoAutoGrating		0x0128

// events
#define				DISPID_MonoPositionChanged		0x0140
#define				DISPID_MonoGratingChanged		0x0141
#define				DISPID_MonoAutoGratingChanged	0x0142
