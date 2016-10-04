#pragma once
class CMyReggie
{
public:
	CMyReggie();
	~CMyReggie();
	BOOL				GetStringValue(
							LPCTSTR				szValueName,
							LPTSTR				szValue,
							UINT				nBufferSize);
	void				SetStringValue(
							LPCTSTR				szValueName,
							LPCTSTR				szValue);
protected:
	HKEY				GetNamedKey(
							LPCTSTR				szValueName);
};

