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
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_STDAFX_H__C65A6F92_FD1C_4429_A678_9B447D4A39B9__INCLUDED_)
#define AFX_STDAFX_H__C65A6F92_FD1C_4429_A678_9B447D4A39B9__INCLUDED_

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

#include "Messages.h" // messages report to event log
#include "Mutex.h"
#include "comutil.h"


// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
//class CServiceModule : public CComModule
class CServiceModule : public CComAutoThreadModule<CComSimpleThreadAllocator>
{
public:
	HRESULT RegisterServer(BOOL bRegTypeLib, BOOL bService);
	HRESULT UnregisterServer();
	void Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID, const GUID* plibid = NULL);
    void Start();
	void ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    void Handler(DWORD dwOpcode);
    void Run();
    BOOL IsInstalled();
    BOOL Install();
    BOOL Uninstall();
	LONG Unlock();

	void LogEventError(LPCTSTR pszFormat, ...) {LogEvent(EVENTLOG_ERROR_TYPE, MESSAGEQUEUELISTENERLOG_ERROR_TYPE, &pszFormat);}
	void LogEventInfo(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_INFORMATION_TYPE, MESSAGEQUEUELISTENERLOG_INFORMATION_TYPE, &pszFormat);}
	void LogEventWarning(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_WARNING_TYPE, MESSAGEQUEUELISTENERLOG_WARNING_TYPE, &pszFormat);}

    void SetServiceStatus(DWORD dwState);
    void SetupAsLocalServer();
	_bstr_t GetVersion();

//Implementation
private:
	static void WINAPI _ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    static void WINAPI _Handler(DWORD dwOpcode);

	void LogEvent(WORD wEventType, DWORD dwEventId, LPCTSTR* ppszFormatplusArgs);
	namespaceMutex::CCriticalSection m_csEventSource;

// data members
public:
    TCHAR m_szServiceName[256];
    SERVICE_STATUS_HANDLE m_hServiceStatus;
    SERVICE_STATUS m_status;
	DWORD dwThreadID;
	BOOL m_bService;
};

extern CServiceModule _Module;
#include <atlcom.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__C65A6F92_FD1C_4429_A678_9B447D4A39B9__INCLUDED)
