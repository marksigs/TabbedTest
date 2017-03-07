///////////////////////////////////////////////////////////////////////////////
//	FILE:			StdAfx.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
///////////////////////////////////////////////////////////////////////////////

//   include file for standard system include files,
//   or project specific include files that are used frequently,
//   but are changed infrequently

#if !defined(AFX_STDAFX_H__E3DA1BC9_63CB_11D4_8241_005004E8D1A7__INCLUDED_)
#define AFX_STDAFX_H__E3DA1BC9_63CB_11D4_8241_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

// Provides facility for messages to be displayed on compilation with its file
// and line number. Use as follows:
// #pragma chMSG("message text")
#define chSTR(x) #x
#define chSTR2(x) chSTR(x)
#define chMSG(desc) message(__FILE__ "(" \
	chSTR2(__LINE__) ") : message: " #desc)

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

#ifdef _DEBUG
#define VERIFY(expr) _ASSERTE(expr)
#else
#define VERIFY(expr) (expr);
#endif
#define verify(expr) VERIFY(expr)

#ifdef _DEBUG
#define TRACE(x) OutputDebugString(x)
#else
#define TRACE(x)
#endif

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__E3DA1BC9_63CB_11D4_8241_005004E8D1A7__INCLUDED)
