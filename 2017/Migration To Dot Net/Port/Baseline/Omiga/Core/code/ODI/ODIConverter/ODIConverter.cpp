///////////////////////////////////////////////////////////////////////////////
//	FILE:			ODIConverter.cpp
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


// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL, 
//      run nmake -f ODIConverterps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include <stdio.h>
#include "About.h"
#include "Exception.h"
#include "MetaData.h"
#include "ODIConverter.h"
#include "ODIConverter_i.c"
#include "ODIConverter1.h"
#include "WSASocket.h"

CServiceModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_ODIConverter, CODIConverter1)
END_OBJECT_MAP()


LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
    while (p1 != NULL && *p1 != NULL)
    {
        LPCTSTR p = p2;
        while (p != NULL && *p != NULL)
        {
            if (*p1 == *p)
                return CharNext(p1);
            p = CharNext(p);
        }
        p1 = CharNext(p1);
    }
    return NULL;
}

// Although some of these functions are big they are declared inline since they are only used once

inline HRESULT CServiceModule::RegisterServer(BOOL bRegTypeLib, BOOL bService, BOOL bAuto)
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
        return hr;

    // Remove any previous service since it may point to
    // the incorrect file
    Uninstall();

    // Add service entries
    UpdateRegistryFromResource(IDR_ODIConverter, TRUE);

    // Adjust the AppID for Local Server or Service
    CRegKey keyAppID;
    LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_WRITE);
    if (lRes != ERROR_SUCCESS)
        return lRes;

    CRegKey key;
    lRes = key.Open(keyAppID, _T("{8CBEA9F1-106D-4B90-A84A-E47FB2788E71}"), KEY_WRITE);
    if (lRes != ERROR_SUCCESS)
        return lRes;
    key.DeleteValue(_T("LocalService"));
    
    if (bService)
    {
        //key.SetValue(_T("ODIConverter"), _T("LocalService"));
		key.SetValue(m_szServiceName, _T("LocalService"));
        key.SetValue(_T("-Service"), _T("ServiceParameters"));
        // Create service
        Install(bAuto);
    }

    // Add object entries
    hr = CComModule::RegisterServer(bRegTypeLib);

    CoUninitialize();

    return hr;
}

inline HRESULT CServiceModule::Help()
{
    TCHAR    chMsg[1024];
	_sntprintf(
		chMsg, 
		_countof(chMsg) - 1, 
		_T("%s service options:\r\n\r\n")
		_T("-UnregServer\tUnregister the service.\r\n")
		_T("-RegServer\tRegisters the COM object and type library.\r\n")
		_T("-Service\t\tRegisters the COM object and type library, and installs as an manual start Windows service.\r\n")
		_T("-ServiceAuto\tRegisters the COM object and type library, and installs as an auto start Windows service.\r\n"), 
		m_szServiceName);
	MessageBox(NULL, chMsg, m_szServiceName, MB_OK);
    return S_OK;
}


inline HRESULT CServiceModule::UnregisterServer()
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
        return hr;

    // Remove service entries
    UpdateRegistryFromResource(IDR_ODIConverter, FALSE);
    // Remove service
    Uninstall();
    // Remove object entries
    CComModule::UnregisterServer(TRUE);
    CoUninitialize();
    return S_OK;
}

inline void CServiceModule::Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID, const GUID* plibid)
{
	// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
    //CComModule::Init(p, h, plibid);
	// Fourth parameter is thread pool size; defaults to number of CPUs * 4.
	// CComSimpleThreadAllocator is the default, round robin, thread allocator.
	CComAutoThreadModule<CComSimpleThreadAllocator>::Init(p, h, plibid);

    m_bService = TRUE;

    // Get the executable file path.
    ::GetModuleFileName(NULL, m_szModuleFileName, _countof(m_szModuleFileName));
	_tsplitpath(m_szModuleFileName, m_szDrive, m_szDir, m_szFName, m_szExt);

	// Use the execute file name as the service name; this enables multiple ODI Converter 
	// services on the same machine.
	_tcscpy(m_szServiceName, m_szFName);
    //LoadString(h, nServiceNameID, m_szServiceName, _countof(m_szServiceName));

    // set up the initial service status 
    m_hServiceStatus = NULL;
    m_status.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
    m_status.dwCurrentState = SERVICE_STOPPED;
    m_status.dwControlsAccepted = SERVICE_ACCEPT_STOP;
    m_status.dwWin32ExitCode = 0;
    m_status.dwServiceSpecificExitCode = 0;
    m_status.dwCheckPoint = 0;
    m_status.dwWaitHint = 0;
}

LONG CServiceModule::Unlock()
{
	// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
    //LONG l = CComModule::Unlock();
	LONG l = CComAutoThreadModule<CComSimpleThreadAllocator>::Unlock();

#ifndef _DEBUG
    if (l == 0 && !m_bService)
	{
        PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
	}
#endif

    return l;
}

BOOL CServiceModule::IsInstalled()
{
    BOOL bResult = FALSE;

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);

    if (hSCM != NULL)
    {
        SC_HANDLE hService = ::OpenService(hSCM, m_szServiceName, SERVICE_QUERY_CONFIG);
        if (hService != NULL)
        {
            bResult = TRUE;
            ::CloseServiceHandle(hService);
        }
        ::CloseServiceHandle(hSCM);
    }
    return bResult;
}

inline BOOL CServiceModule::Install(BOOL bAuto)
{
    if (IsInstalled())
        return TRUE;

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
    if (hSCM == NULL)
    {
        MessageBox(NULL, _T("Couldn't open service manager"), m_szServiceName, MB_OK);
        return FALSE;
    }

	DWORD dwServiceType = SERVICE_WIN32_OWN_PROCESS;
#ifdef _DEBUG
	// Allows ASSERT message boxes to be displayed.
	dwServiceType |= SERVICE_INTERACTIVE_PROCESS;
#endif

    SC_HANDLE hService = ::CreateService(
        hSCM, m_szServiceName, m_szServiceName,
        SERVICE_ALL_ACCESS, dwServiceType,
        bAuto ? SERVICE_AUTO_START : SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL,
        m_szModuleFileName, NULL, NULL, _T("RPCSS\0"), NULL, NULL);

    if (hService == NULL)
    {
        ::CloseServiceHandle(hSCM);
        MessageBox(NULL, _T("Couldn't create service"), m_szServiceName, MB_OK);
        return FALSE;
    }

    ::CloseServiceHandle(hService);
    ::CloseServiceHandle(hSCM);

    return TRUE;
}

inline BOOL CServiceModule::Uninstall()
{
    if (!IsInstalled())
        return TRUE;

    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);

    if (hSCM == NULL)
    {
        MessageBox(NULL, _T("Couldn't open service manager"), m_szServiceName, MB_OK);
        return FALSE;
    }

    SC_HANDLE hService = ::OpenService(hSCM, m_szServiceName, SERVICE_STOP | DELETE);

    if (hService == NULL)
    {
        ::CloseServiceHandle(hSCM);
        MessageBox(NULL, _T("Couldn't open service"), m_szServiceName, MB_OK);
        return FALSE;
    }
    SERVICE_STATUS status;
    ::ControlService(hService, SERVICE_CONTROL_STOP, &status);

    BOOL bDelete = ::DeleteService(hService);
    ::CloseServiceHandle(hService);
    ::CloseServiceHandle(hSCM);

    if (bDelete)
        return TRUE;

    MessageBox(NULL, _T("Service could not be deleted"), m_szServiceName, MB_OK);
    return FALSE;
}

///////////////////////////////////////////////////////////////////////////////////////
// Logging functions
void CServiceModule::LogEvent(WORD wEventType, DWORD dwEventId, LPCTSTR* ppszFormatplusArgs)
{
    TCHAR    chMsg[256];
    HANDLE  hEventSource;
    LPTSTR  lpszStrings[1];
    va_list pArg;

    va_start(pArg, *ppszFormatplusArgs);
    _vstprintf(chMsg, *ppszFormatplusArgs, pArg);
    va_end(pArg);

    lpszStrings[0] = chMsg;

//	if (m_bService)
	if (TRUE)	// AS 06/08/01 Always log events to event log.
	{
		// prevent more than one thread reporting an event at the same time
		namespaceMutex::CSingleLock lckEventSource(&m_csEventSource, TRUE);

		// Get a handle to use with ReportEvent(). 
		hEventSource = RegisterEventSource(NULL, m_szServiceName);
		if (hEventSource != NULL)
		{
			// Write to event log.
			ReportEvent(hEventSource, wEventType, 0, dwEventId, NULL, 1, 0, (LPCTSTR*) &lpszStrings[0], NULL);
			DeregisterEventSource(hEventSource);
		}
    }
    else
    {
        // As we are not running as a service, just write the error to the console.
        _putts(chMsg);
    }

	LogDebug(_T("%s\n"), chMsg);
}

void CServiceModule::LogDebug(LPCTSTR pszFormat, ...)
{
	TCHAR chMsg[1024] = _T("");

#ifndef _DEBUG
	// Don't bother with thread id in DEBUG build.
	_sntprintf(chMsg, _countof(chMsg) - 1, _T("%s,"), GetThreadID(chMsg, 16));
#endif

	va_list pArg;

	va_start(pArg, pszFormat);
	_vsntprintf(chMsg + _tcslen(chMsg), _countof(chMsg) - (_tcslen(chMsg) + 1), pszFormat, pArg);
	va_end(pArg);

	::OutputDebugString(chMsg);
}

_TCHAR* CServiceModule::GetThreadID(_TCHAR* pszThreadID, int nMaxLenThreadID) const
{
	::ZeroMemory(pszThreadID, nMaxLenThreadID);

	DWORD dwCurrentThreadID	= ::GetCurrentThreadId();
	int nThreadPriority		= ::GetThreadPriority(::GetCurrentThread());

	VERIFY(
		_sntprintf(
			pszThreadID, 
			nMaxLenThreadID - 1, 
			_T("0x%x(%d)"),
			dwCurrentThreadID, 
			nThreadPriority) > -1);

	return pszThreadID;
}


//////////////////////////////////////////////////////////////////////////////////////////////
// Service startup and registration
inline void CServiceModule::Start()
{
    SERVICE_TABLE_ENTRY st[] =
    {
        { m_szServiceName, _ServiceMain },
        { NULL, NULL }
    };
    if (m_bService && !::StartServiceCtrlDispatcher(st))
    {
        m_bService = FALSE;
    }
    if (m_bService == FALSE)
        Run();
}

inline void CServiceModule::ServiceMain(DWORD /* dwArgc */, LPTSTR* /* lpszArgv */)
{
    // Register the control request handler
    m_status.dwCurrentState = SERVICE_START_PENDING;
    m_hServiceStatus = RegisterServiceCtrlHandler(m_szServiceName, _Handler);
    if (m_hServiceStatus == NULL)
    {
        LogEventError(_T("Handler not installed"));
        return;
    }
    SetServiceStatus(SERVICE_START_PENDING);

    m_status.dwWin32ExitCode = S_OK;
    m_status.dwCheckPoint = 0;
    m_status.dwWaitHint = 0;

    // When the Run function returns, the service has stopped.
    Run();

	CWSASocket::WSACleanup();

    SetServiceStatus(SERVICE_STOPPED);
    LogEventInfo(_T("Service stopped"));
}

inline void CServiceModule::Handler(DWORD dwOpcode)
{
    switch (dwOpcode)
    {
    case SERVICE_CONTROL_STOP:
        SetServiceStatus(SERVICE_STOP_PENDING);
        PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
        break;
    case SERVICE_CONTROL_PAUSE:
        break;
    case SERVICE_CONTROL_CONTINUE:
        break;
    case SERVICE_CONTROL_INTERROGATE:
        break;
    case SERVICE_CONTROL_SHUTDOWN:
        break;
    default:
        LogEventError(_T("Bad service request"));
    }
}

void WINAPI CServiceModule::_ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
    _Module.ServiceMain(dwArgc, lpszArgv);
}
void WINAPI CServiceModule::_Handler(DWORD dwOpcode)
{
    _Module.Handler(dwOpcode); 
}

void CServiceModule::SetServiceStatus(DWORD dwState)
{
    m_status.dwCurrentState = dwState;
    ::SetServiceStatus(m_hServiceStatus, &m_status);
}

void CServiceModule::Run()
{
	bool bSuccess = true;

    _Module.dwThreadID = GetCurrentThreadId();

    HRESULT hr = CoInitialize(NULL);
//  If you are running on NT 4.0 or higher you can use the following call
//  instead to make the EXE free threaded.
//  This means that calls come in on a random RPC thread
//  HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

    _ASSERTE(SUCCEEDED(hr));

    // This provides a NULL DACL which will allow access to everyone.
    CSecurityDescriptor sd;
    sd.InitializeFromThreadToken();
    hr = CoInitializeSecurity(sd, -1, NULL, NULL,
        RPC_C_AUTHN_LEVEL_PKT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    _ASSERTE(SUCCEEDED(hr));

    hr = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, REGCLS_MULTIPLEUSE);
    _ASSERTE(SUCCEEDED(hr));

	bSuccess = false;
	try
	{
		CWSASocket::WSAStartup();
		CMetaData::Init();
		bSuccess = true;
	}
	catch(_com_error& e)
	{
		CException Exception(E_GENERICERROR, e, __FILE__, __LINE__);
		CODIConverter1::LogError(Exception);
	}
	catch(CException& e)
	{
		CODIConverter1::LogError(e);
	}
	catch(...)
	{
		CException Exception(E_GENERICERROR, __FILE__, __LINE__, _T("Unknown error"));
		CODIConverter1::LogError(Exception);
	}

	if (bSuccess)
	{
		DisplayAboutInfo();

		LogEventInfo(_T("Service started"));
		if (m_bService)
			SetServiceStatus(SERVICE_RUNNING);

		// helper to debug the service	
#ifdef _DEBUG
		if (m_bService)
		{
			::DebugBreak();
		}
#endif

#ifdef _DEBUG
//		_CrtMemState MemCheckpoint;
//		_CrtMemCheckpoint(&MemCheckpoint);
#endif

		MSG msg;
		BOOL bResult = 0;
		while ((bResult = GetMessage(&msg, 0, 0, 0)) != 0)
		{
			if (bResult == -1)
			{
				// Error.
			}
			else
			{
				DispatchMessage(&msg);
			}
		}

#ifdef _DEBUG
//		_CrtMemDumpAllObjectsSince(&MemCheckpoint);
#endif
	}

    _Module.RevokeClassObjects();

    CoUninitialize();
}

/////////////////////////////////////////////////////////////////////////////
//
extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, 
    HINSTANCE /*hPrevInstance*/, LPTSTR lpCmdLine, int /*nShowCmd*/)
{
    lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT
    _Module.Init(ObjectMap, hInstance, IDS_SERVICENAME, &LIBID_ODICONVERTER);
    _Module.m_bService = TRUE;

    TCHAR szTokens[] = _T("-/");

    LPCTSTR lpszToken = FindOneOf(lpCmdLine, szTokens);
    while (lpszToken != NULL)
    {
        if (lstrcmpi(lpszToken, _T("?"))==0 || lstrcmpi(lpszToken, _T("help"))==0)
            return _Module.Help();

        if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
            return _Module.UnregisterServer();

        // Register as Local Server
        if (lstrcmpi(lpszToken, _T("RegServer"))==0)
            return _Module.RegisterServer(TRUE, FALSE, FALSE);
        
        // Register as Service
        if (lstrcmpi(lpszToken, _T("Service"))==0)
            return _Module.RegisterServer(TRUE, TRUE, FALSE);

        // Register as auto start Service
        if (lstrcmpi(lpszToken, _T("ServiceAuto"))==0)
            return _Module.RegisterServer(TRUE, TRUE, TRUE);
        
        lpszToken = FindOneOf(lpszToken, szTokens);
    }

    // Are we Service or Local Server
    CRegKey keyAppID;
    LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_READ);
    if (lRes != ERROR_SUCCESS)
        return lRes;

    CRegKey key;
    lRes = key.Open(keyAppID, _T("{8CBEA9F1-106D-4B90-A84A-E47FB2788E71}"), KEY_READ);
    if (lRes != ERROR_SUCCESS)
        return lRes;

    TCHAR szValue[_MAX_PATH];
    DWORD dwLen = _MAX_PATH;
    lRes = key.QueryValue(szValue, _T("LocalService"), &dwLen);

    _Module.m_bService = FALSE;
    if (lRes == ERROR_SUCCESS)
        _Module.m_bService = TRUE;

    _Module.Start();

    // When we get here, the service has been stopped
    return _Module.m_status.dwWin32ExitCode;
}

_bstr_t CServiceModule::MakeModulePath(LPCWSTR szOldPath)
{
	wchar_t szNewPath[_MAX_PATH]	= L"";
	wchar_t szDrive[_MAX_DRIVE]		= L"";
	wchar_t szDir[_MAX_DIR]			= L"";
	wchar_t szFName[_MAX_FNAME]		= L"";
	wchar_t szExt[_MAX_EXT]			= L"";

	_wsplitpath(szOldPath, szDrive, szDir, szFName, szExt);

	// Fill in the missing bits using the module path.
	if (wcslen(szDrive) == 0)
	{
		wcscpy(szDrive, m_szDrive);
	}
	if (wcslen(szDir) == 0)
	{
		wcscpy(szDir, m_szDir);

		wchar_t* pPos = NULL;
#ifdef _DEBUG
		if ((pPos = wcsstr(szDir, L"\\DebugU\\\0")) != NULL)
		{
			*(pPos + 1) = L'\0';
		}
#else
		if ((pPos = wcsstr(szDir, L"\\ReleaseUMinDependency\\\0")) != NULL)
		{
			*(pPos + 1) = L'\0';
		}
#endif
	}
	if (wcslen(szFName) == 0)
	{
		wcscpy(szFName, m_szFName);
	}
	if (wcslen(szExt) == 0)
	{
		wcscpy(szExt, m_szExt);
	}

	swprintf(szNewPath, L"%s%s%s%s", szDrive, szDir, szFName, szExt);

	return szNewPath;
}

_bstr_t CServiceModule::RemoveModulePath(LPCWSTR szOldPath)
{
	wchar_t szNewPath[_MAX_PATH]	= L"";
	wchar_t szDrive[_MAX_DRIVE]		= L"";
	wchar_t szDir[_MAX_DIR]			= L"";
	wchar_t szFName[_MAX_FNAME]		= L"";
	wchar_t szExt[_MAX_EXT]			= L"";

	_wsplitpath(szOldPath, szDrive, szDir, szFName, szExt);

	// Remove bits that match the module path.
	if (wcsicmp(szDrive, m_szDrive) == 0)
	{
		wcscpy(szDrive, L"");
	}

	wchar_t szModuleDir[_MAX_DIR]	= L"";
	
	wcscpy(szModuleDir, m_szDir);
	if (wcsicmp(szDir, szModuleDir) == 0)
	{
		wcscpy(szDir, L"");
	}

	wcscpy(szModuleDir, m_szDir);
	wchar_t* pPos = NULL;
#ifdef _DEBUG
	if ((pPos = wcsstr(szModuleDir, L"\\DebugU\\\0")) != NULL)
	{
		*(pPos + 1) = L'\0';
	}
#else
	if ((pPos = wcsstr(szModuleDir, L"\\ReleaseUMinDependency\\\0")) != NULL)
	{
		*(pPos + 1) = L'\0';
	}
#endif
	if (wcsicmp(szDir, szModuleDir) == 0)
	{
		wcscpy(szDir, L"");
	}
	
	if (wcsicmp(szFName, m_szFName) == 0)
	{
		wcscpy(szFName, L"");
	}
	if (wcsicmp(szExt, m_szExt) == 0)
	{
		wcscpy(szExt, L"");
	}

	swprintf(szNewPath, L"%s%s%s%s", szDrive, szDir, szFName, szExt);

	return szNewPath;
}


///////////////////////////////////////////////////////////////////////////////
//	From C++ Programming Language, Bjarne Stroustrup.
//	Used for case declaring case insensitive maps of strings to other objects, 
//	e.g., map<string, int, Nocase>
//	Returns true if x is lexicographically < y, ignoring case.
///////////////////////////////////////////////////////////////////////////////
bool Nocase::operator()(const _bstr_t& x, const _bstr_t& y) const
{
	return operator()(std::wstring(static_cast<LPCWSTR>(x)), std::wstring(static_cast<LPCWSTR>(y)));
}

bool Nocase::operator()(const std::wstring& x, const std::wstring& y) const
{
	bool bLessThan = false;

	std::wstring::const_iterator p = x.begin();
	std::wstring::const_iterator q = y.begin();

	_ASSERTE(p != NULL);
	_ASSERTE(q != NULL);

	while (p != x.end() && q != y.end() && towupper(*p) == towupper(*q))
	{
		++p;
		++q;
	}

	if (p == x.end())
	{
		bLessThan = q != y.end();
	}
	else
	{
		bLessThan = towupper(*p) < towupper(*q);
	}

	return bLessThan;
}

bool Nocase::operator()(const std::string& x, const std::string& y) const
{
	bool bLessThan = false;

	std::string::const_iterator p = x.begin();
	std::string::const_iterator q = y.begin();

	_ASSERTE(p != NULL);
	_ASSERTE(q != NULL);

	while (p != x.end() && q != y.end() && toupper(*p) == toupper(*q))
	{
		++p;
		++q;
	}

	if (p == x.end())
	{
		bLessThan = q != y.end();
	}
	else
	{
		bLessThan = toupper(*p) < toupper(*q);
	}

	return bLessThan;
}

///////////////////////////////////////////////////////////////////////////////
// Hexadecimal string conversion functions.

int hatoi(const char* string)
{
	int nHex = 0;
	sscanf(string, "%x", &nHex);
	return nHex;
}

int hwtoi(const wchar_t* string)
{
	int nHex = 0;
	swscanf(string, L"%x", &nHex);
	return nHex;
}

long hatol(const char* string)
{
	long lHex = 0L;
	sscanf(string, "%x", &lHex);
	return lHex;
}

long hwtol(const wchar_t* string)
{
	long lHex = 0L;
	swscanf(string, L"%x", &lHex);
	return lHex;
}

bool IsTrueString(LPCWSTR szValue)
{
	return 
		wcslen(szValue) > 0 &&
		(
			wcsicmp(szValue, L"TRUE") == 0 ||
			towupper(szValue[0]) == L'Y' ||
			wcscmp(szValue, L"1") == 0);
}

