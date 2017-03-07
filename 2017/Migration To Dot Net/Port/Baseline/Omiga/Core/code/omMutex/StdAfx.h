/* <VERSION  CORELABEL="" LABEL="R15" DATE="10/10/2003 15:52:58" VERSION="254" PATH="$/CodeCoreCust/4UATCust/Code/Synchronisation/omMutex/StdAfx.h"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			stdafx.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      13/03/03    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_STDAFX_H__1A891E9C_8901_4D74_9606_6D56E835C03C__INCLUDED_)
#define AFX_STDAFX_H__1A891E9C_8901_4D74_9606_6D56E835C03C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0500
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

//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module

///////////////////////////////////////////////////////////////////////////////

// Comment out the next line to prevent #pragma messages from being displayed.
//#define MESSAGES
#ifdef MESSAGES
#define MESSAGE(X) X
#else
#define MESSAGE(X) ""
#endif

///////////////////////////////////////////////////////////////////////////////
//	CObject
//		Base class for all dynamically allocated objects.
///////////////////////////////////////////////////////////////////////////////

#pragma chMSG("memory management to reenable")
class CObject
{
public:
	// use MP heap for dynamically allocated objects of this class
//#ifdef _DEBUG
//	inline void* operator new(size_t nSize, int nBlockType, const char* pszFileName, int nLine) 
//		throw(std::bad_alloc) 
//	{ 
//		return g_pMpHeap ? g_pMpHeap->New(nSize, nBlockType, pszFileName, nLine) : NULL; 
//	}
//#if _MSC_VER >= 1200
//	inline void operator delete(void* pMem, int /* nBlockType */, const char* /* pszFileName */, int /* nLine */) { if (g_pMpHeap) g_pMpHeap->Free(pMem); }
//#endif
//#else
//	inline void* operator new(size_t nSize) throw(std::bad_alloc) { return g_pMpHeap ? g_pMpHeap->New(nSize) : NULL; }
//#endif
//	inline void operator delete(void* pMem) { if (g_pMpHeap) g_pMpHeap->Free(pMem); }
};


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

///////////////////////////////////////////////////////////////////////////////

#include "mutex.h"


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__1A891E9C_8901_4D74_9606_6D56E835C03C__INCLUDED)
