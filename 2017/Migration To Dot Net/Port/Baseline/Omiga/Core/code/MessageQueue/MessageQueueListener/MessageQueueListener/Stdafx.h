///////////////////////////////////////////////////////////////////////////////
//	FILE:			Stdafx.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

//	include file for standard system include files,
//  or project specific include files that are used frequently,
//  but are changed infrequently

#if !defined(AFX_STDAFX_H__2B0E56B4_4B55_11D4_8237_005004E8D1A7__INCLUDED_)
#define AFX_STDAFX_H__2B0E56B4_4B55_11D4_8237_005004E8D1A7__INCLUDED_

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


#include "Messages.h" // messages report to event log
#include "Mutex.h"
#include "comutil.h"
#include "Log.h"
#include "ScheduleManager.h"

enum DayOfWeek
{
	eSunday = 0,
	eMonday = 1,
	eTuesday = 2,
	eWednesday = 3,
	eThursday = 4,
	eFriday = 5,
	eSaturday = 6
};


// Error codes
#define ERROR_UNABLETOSTARTQUEUE		L"600"
#define ERROR_UNABLETOFINDNAMEDQUEUE	L"601"
#define ERROR_UNABLETOCREATEDOMDOCUMENT	L"602"
#define ERROR_UNABLETOGETROOTELEMENT	L"603"
#define ERROR_UNABLETOPROCESSTAG		L"604"
#define ERROR_UNDEFINEDERROR			L"605"
#define ERROR_SEARCHOFQUEUEFAILED		L"606"
#define ERROR_ADDEVENTTOQUEUEFAILED		L"607"
#define ERROR_FAILEDTOUNPACKSAFEARRAY	L"608"

///////////////////////////////////////////////////////////////////////////////

class CServiceModule : public CComModule, public CLog
{
public:
	CServiceModule();
	virtual ~CServiceModule();

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
	void LogEventError(LPCTSTR pszFormat, ...) {LogEvent(EVENTLOG_ERROR_TYPE, MESSAGEQUEUELISTENER_ERROR_TYPE, &pszFormat);}
	void LogEventInfo(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_INFORMATION_TYPE, MESSAGEQUEUELISTENER_INFORMATION_TYPE, &pszFormat);}
	void LogEventWarning(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_WARNING_TYPE, MESSAGEQUEUELISTENER_WARNING_TYPE, &pszFormat);}
    BOOL SetServiceStatus(DWORD dwState, DWORD dwWaitHint = 0);
    void SetupAsLocalServer();

	_bstr_t GetXmlConfigurationFileName();
	bool XmlConfigurationFileExists();

protected:
	void ReadSCMWaitHints();

	_bstr_t GetVersion();
	_bstr_t ReadDependantServices();

//Implementation
private:
	static void WINAPI _ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    static void WINAPI _Handler(DWORD dwOpcode);

	void LogEvent(WORD wEventType, DWORD dwEventId, LPCTSTR* ppszFormatplusArgs);
	namespaceMutex::CCriticalSection m_csEventSource;

	void TestMessageQueueListenerMTS1();

// Poller used to prevent the SCM from declaring this service in error during
// length start/stop operations
private:
	class CSCMCheckPoint : public CScheduleManager
	{
	public:
		CSCMCheckPoint(CServiceModule& ServiceModule);
		virtual ~CSCMCheckPoint();

		void put_dwNextWaitIntervalms(DWORD dwWaitIntervalms) {m_dwNextWaitIntervalms = dwWaitIntervalms;}

		virtual bool StartUp();
		virtual void CloseDown();

	private:
		virtual DWORD GetdwNextWaitIntervalms() {return m_dwNextWaitIntervalms;}
		virtual void OnEventSchedule();
		
		DWORD m_dwNextWaitIntervalms;
		CServiceModule& m_ServiceModule;
		DWORD m_dwWaitHintPending;
	} m_SCMCheckPoint;
	friend class CSCMCheckPoint;

// data members
public:
    TCHAR m_szServiceName[256];
    SERVICE_STATUS_HANDLE m_hServiceStatus;
    SERVICE_STATUS m_status;
	DWORD dwThreadID;
	BOOL m_bService;
	BOOL m_bExecutingRegistrationCode;
};

extern CServiceModule _Module;
#include <atlcom.h>


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__2B0E56B4_4B55_11D4_8237_005004E8D1A7__INCLUDED)
