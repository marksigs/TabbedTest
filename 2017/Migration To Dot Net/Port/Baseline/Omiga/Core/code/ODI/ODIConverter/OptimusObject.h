// OptimusObject.h: interface for the COptimusObject class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OPTIMUSOBJECT_H__2F7372ED_130B_4B21_9816_F8A97F253AD2__INCLUDED_)
#define AFX_OPTIMUSOBJECT_H__2F7372ED_130B_4B21_9816_F8A97F253AD2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CCodePage;

class COptimusObject  
{
public:
	COptimusObject(BYTE* pObject = NULL, LPCWSTR szTag = L"", LPCWSTR lpUnicode = NULL);
	virtual ~COptimusObject();
	static void LogDebugHeader();
	void LogDebug(bool bLogHex = false);
	
protected:
	void Decode(BYTE* pObject, LPCWSTR szTag = L"", LPCWSTR lpUnicode = NULL, CCodePage* pCodePage = NULL);
	short GetShort(const BYTE* pStream);
	long GetLong(const BYTE* pStream);

private:
	void LogDebugItem(_TCHAR* pszStream, short nItem, bool bLogHex = false);
	void LogDebugItem(_TCHAR* pszStream, long lItem, bool bLogHex = false);
	void LogDebugItem(_TCHAR* pszStream, LPCWSTR lpItem, bool bLogHex = false);
	void LogDebugHexString(_TCHAR* pszStream, BYTE* pHex, int nMaxSizeHex, bool bLogHex = false, bool bReversed = false);
	TCHAR* HexToString(_TCHAR* pszStream, BYTE* pHex, int nMaxSizeHex, bool bReversed = false);

private:
	BYTE*				m_pObject;
	short				m_nLength;
	long				m_lNumber;
	long				m_lParent;
	short				m_nId;
	_bstr_t				m_bstrTag;
	wchar_t*			m_lpUnicode;
	int					m_nMaxSizeUnicode;
	bool				m_bOwnsUnicode;
	LPCSTR				m_lpEBCDIC;
	int					m_nMaxSizeEBCDIC;
};

#endif // !defined(AFX_OPTIMUSOBJECT_H__2F7372ED_130B_4B21_9816_F8A97F253AD2__INCLUDED_)
