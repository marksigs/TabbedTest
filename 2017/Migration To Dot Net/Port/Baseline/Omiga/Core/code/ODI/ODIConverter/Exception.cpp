///////////////////////////////////////////////////////////////////////////////
//	FILE:			Exception.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CException::CException(long lError, LPCSTR lpszFile, int nLine)
{
	SetType(L"BASE");
	SetError(lError);
	SetFile(lpszFile);
	SetLine(nLine);
	SetDescription(L"");
	SetIID(GUID_NULL);
	SetHResult(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, lError));
}

CException::CException(long lError, LPCSTR lpszFile, int nLine, LPCWSTR lpszFormat, ...)
{
	SetType(L"BASE");
	SetError(lError);
	SetFile(lpszFile);
	SetLine(nLine);
	SetDescription(&lpszFormat);
	SetIID(GUID_NULL);
	SetHResult(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, lError));
}

CException::CException(long lError, _com_error& e, LPCSTR lpszFile, int nLine) 
{
	SetType(L"COM");
	SetError(lError);
	SetFile(lpszFile);
	SetLine(nLine);
	SetDescription(L"%s", static_cast<LPCWSTR>(e.Description()));
	SetIID(e.GUID());
	SetHResult(e.Error());
}

CException::CException(long lError, exception& e, LPCSTR lpszFile, int nLine)
{
	SetType(L"STL");
	SetError(lError);
	SetFile(lpszFile);
	SetLine(nLine);
	SetDescription(e.what());
	SetIID(GUID_NULL);
	SetHResult(MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, lError));
}

void CException::SetDescription(LPCWSTR lpszFormat, ...)
{
	SetDescription(&lpszFormat);
}

void CException::SetDescription(LPCWSTR* ppszFormatplusArgs)
{
    va_list pArg;

    va_start(pArg, *ppszFormatplusArgs);
    vswprintf(m_szDescription, *ppszFormatplusArgs, pArg);
    va_end(pArg);
}

void CException::SetDescription(LPCSTR lpszError)
{
#ifdef UNICODE
	AtlA2WHelper(m_szDescription, lpszError, _countof(m_szDescription));
#else
	wcscpy(m_szDescription, lpszError);
#endif
}

