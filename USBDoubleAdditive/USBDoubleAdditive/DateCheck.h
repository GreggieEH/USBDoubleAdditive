#pragma once
class CDateCheck
{
public:
	CDateCheck();
	~CDateCheck();
	BOOL			MyDateCheck();
protected:
	BOOL			MyCheckDate(
						LPTSTR			szLastOpDate,
						UINT			nBufferSize,
						LPCTSTR			szExpDate,
						long		*	diffCount);
	BOOL			CheckDiffCount(
						DATE			lastOp,			// last operated date
						DATE			expDate,		// expiry date
						long		*	diffCount);
};

