///////////////////////////////////////////////////////////////////////////////
//	FILE:			Exception.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_EXCEPTION_H__B95AFCF6_D9B2_46FE_BCE3_8839122DB5AE__INCLUDED_)
#define AFX_EXCEPTION_H__B95AFCF6_D9B2_46FE_BCE3_8839122DB5AE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <exception>

class CException : public std::exception
{
public:
	CException(long lError, LPCSTR lpszFile, int nLine);
	CException(long lError, LPCSTR lpszFile, int nLine, LPCWSTR lpszFormat, ...);
	CException(long lError, _com_error& e, LPCSTR lpszFile, int nLine);
	CException(long lError, exception& e, LPCSTR lpszFile, int nLine);
	virtual ~CException() {};

	void SetType(LPCWSTR lpszType) { ::ZeroMemory(m_szType, _countof(m_szType)); wcscpy(m_szType, lpszType); }
	LPCWSTR GetType() const { return m_szType; }
	void SetFile(LPCSTR lpszFile) { ::ZeroMemory(m_szFile, _countof(m_szFile)); AtlA2WHelper(m_szFile, lpszFile, sizeof(m_szFile)); }
	LPCWSTR GetFile() const { return m_szFile; }
	void SetLine(int nLine)	{ m_nLine = nLine; }
	int GetLine() const { return m_nLine; }
	void SetError(long lError) { m_lError = lError; }
	long GetError() const { return m_lError; }
	void SetDescription(LPCSTR lpszError);
	void SetDescription(LPCWSTR lpszFormat, ...);
	void SetDescription(LPCWSTR* ppszFormatplusArgs);
	LPCWSTR GetDescription() const { return m_szDescription; }
	void SetIID(IID iid) { m_iid = iid; }
	IID GetIID() const { return m_iid; }
	void SetHResult(HRESULT hResult) { m_hResult = hResult; }
	HRESULT GetHResult() const { return m_hResult; }

protected:
	long			m_lError;
	WCHAR			m_szType[64];
	WCHAR			m_szDescription[256];
	WCHAR			m_szFile[_MAX_PATH];
	int				m_nLine;
	IID				m_iid;
	HRESULT			m_hResult;
};


#endif // !defined(AFX_EXCEPTION_H__B95AFCF6_D9B2_46FE_BCE3_8839122DB5AE__INCLUDED_)
