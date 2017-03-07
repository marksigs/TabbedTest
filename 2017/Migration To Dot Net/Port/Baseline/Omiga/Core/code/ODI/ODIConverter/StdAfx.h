// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__605B51D1_45BC_45C4_B7C7_F7BD2C71F580__INCLUDED_)
#define AFX_STDAFX_H__605B51D1_45BC_45C4_B7C7_F7BD2C71F580__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define _WIN32_WINNT 0x0500

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
//#define _ATL_FREE_THREADED
#define _ATL_APARTMENT_THREADED

#ifdef _DEBUG
// _ATL_DEBUG_INTERFACES not thread safe in release builds - see MSDN Q272024.
//#define _ATL_DEBUG_INTERFACES
#endif

#include <atlbase.h>
#include <comdef.h>
#include <exception>
#pragma warning(disable: 4245)
// "'function' was declared deprecated" - generally for unsafe string functions.
#pragma warning(disable: 4996)

// Comment out the next line to prevent #pragma messages from being displayed.
//#define MESSAGES
#ifdef MESSAGES
#define MESSAGE(X) X
#else
#define MESSAGE(X) ""
#endif

// Provides facility for messages to be displayed on compilation with its file
// and line number. Use as follows:
// #pragma chMSG("message text")
#define chSTR(x) #x
#define chSTR2(x) chSTR(x)
#ifdef MESSAGES
#define chMSG(desc) message(__FILE__ "(" chSTR2(__LINE__) ") : message: " #desc)
#else
#define chMSG(desc)
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

#include "Mutex.h"
#include "Messages.h"

//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module

// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
//class CServiceModule : public CComModule
class CServiceModule : public CComAutoThreadModule<CComSimpleThreadAllocator>
{
public:
	HRESULT RegisterServer(BOOL bRegTypeLib, BOOL bService, BOOL bAuto = FALSE);
	HRESULT UnregisterServer();
	void Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID, const GUID* plibid = NULL);
    void Start();
	void ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    void Handler(DWORD dwOpcode);
    void Run();
    BOOL IsInstalled();
    BOOL Install(BOOL bAuto = FALSE);
    BOOL Uninstall();
	LONG Unlock();
	void LogEventError(LPCTSTR pszFormat, ...) {LogEvent(EVENTLOG_ERROR_TYPE, ODICONVERTER_ERROR_TYPE, &pszFormat);}
	void LogEventInfo(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_INFORMATION_TYPE, ODICONVERTER_INFORMATION_TYPE, &pszFormat);}
	void LogEventWarning(LPCTSTR pszFormat, ...)  {LogEvent(EVENTLOG_WARNING_TYPE, ODICONVERTER_WARNING_TYPE, &pszFormat);}
	void LogDebug(LPCTSTR pszFormat, ...);
    void SetServiceStatus(DWORD dwState);
    void SetupAsLocalServer();
	_bstr_t MakeModulePath(LPCWSTR szOldPath);
	_bstr_t RemoveModulePath(LPCWSTR szOldPath);
	_TCHAR* GetThreadID(_TCHAR* pszThreadID, int nMaxLenThreadID) const;
	HRESULT Help();

//Implementation
private:
	static void WINAPI _ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
    static void WINAPI _Handler(DWORD dwOpcode);

	void LogEvent(WORD wEventType, DWORD dwEventId, LPCTSTR* ppszFormatplusArgs);
	namespaceMutex::CCriticalSection m_csEventSource;

// data members
public:
    _TCHAR					m_szServiceName[256];
    SERVICE_STATUS_HANDLE	m_hServiceStatus;
    SERVICE_STATUS			m_status;
	DWORD					dwThreadID;
	BOOL					m_bService;
    _TCHAR					m_szModuleFileName[_MAX_PATH];
	_TCHAR					m_szDrive[_MAX_DRIVE];
	_TCHAR					m_szDir[_MAX_DIR];
	_TCHAR					m_szFName[_MAX_FNAME];
	_TCHAR					m_szExt[_MAX_EXT];
};

extern CServiceModule _Module;
#include <atlcom.h>

#import "msxml3.dll" rename_namespace("MSXML")
#include "XMLAssist.tlh"
using namespace XMLAssist;


// determine number of elements in an array (not bytes)
#ifndef _countof
#define _countof(array) (sizeof(array)/sizeof(array[0]))
#endif

#ifdef _UNICODE
#define _tstring wstring
#define _tostringstream wostringstream
#define _ttoupper towupper
#else
#define _tstring string
#define _tostringstream ostringstream
#define _ttoupper toupper
#endif

///////////////////////////////////////////////////////////////////////////////
//	From C++ Programming Language, Bjarne Stroustrup.
//	Used for case declaring case insensitive maps of strings to other objects, 
//	e.g., map<string, int, Nocase>
///////////////////////////////////////////////////////////////////////////////
class Nocase
{
public:
	bool operator()(const _bstr_t& x, const _bstr_t& y) const;
	bool operator()(const std::wstring&, const std::wstring&) const;
	bool operator()(const std::string&, const std::string&) const;
};


int hatoi(const char* string);
int hwtoi(const wchar_t* string);
long hatol(const char* string);
long hwtol(const wchar_t* string);

#ifdef UNICODE
#define httoi(s) hwtoi(s)
#define httol(s) hwtol(s)
#else
#define httoi(s) hatoi(s)
#define httol(s) hatol(s)
#endif

bool IsTrueString(LPCWSTR szValue);


//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__605B51D1_45BC_45C4_B7C7_F7BD2C71F580__INCLUDED)
