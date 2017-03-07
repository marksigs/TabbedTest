///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListener.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//	LD		06/04/01	SYS2248 - Add performance counters
//	LD		03/05/01	SYS2296 - Support for Windows Installer.
//							RegServer now installs MQL as a service rather than LocalServer
//							and performance counters are only installed at run-time
//							LocalServer or Console now installs as LocalServer
//	AD		24/05/01	SYS2621 - Add exception handler to ReadDependantServices function
//	LD		03/10/01	SYS2770 - Improve error messaging
//	LD		05/10/01	SYS2779 - write exe version to event log on service start up
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//  LD		06/02/03	OPC0006008 - Prevent calling logging service during registration and
//						unregistration
//  LD		20/03/03	Make MessageQueueListenerLOG a dependent service
//  RF      04/02/04	Use MSXML4 (included as part of BMIDS727)
///////////////////////////////////////////////////////////////////////////////

// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL, 
//      run nmake -f MessageQueueListenerps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "Log.h"
#include "MessageQueueListener.h"
#include "PrfMonMessageQueueListener.h"

#include "MessageQueueListener_i.c"


#include <stdio.h>
#include "MessageQueueListener1.h"
#include "ThreadPoolManagerMSMQ1.h"
#include "ThreadPoolManagerOMMQ1.h"

#import "..\MessageQueueListenerMTS\MessageQueueListenerMTS.tlb" no_namespace

///////////////////////////////////////////////////////////////////////////////

DWORD s_dwWaitHintStartPending = 120000; // ms (wait hint whilst stopping the service)
DWORD s_dwWaitHintStopPending = 120000; // ms (wait hint whilst starting the service)

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); if (!_Module.m_bExecutingRegistrationCode) LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); if (!_Module.m_bExecutingRegistrationCode) LogInOut.AnyThreadInitialize

///////////////////////////////////////////////////////////////////////////////

CServiceModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_MessageQueueListener1, CMessageQueueListener1)
OBJECT_ENTRY(CLSID_ThreadPoolManagerMSMQ1, CThreadPoolManagerMSMQ1)
OBJECT_ENTRY(CLSID_ThreadPoolManagerOMMQ1, CThreadPoolManagerOMMQ1)
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

///////////////////////////////////////////////////////////////////////////////

CServiceModule::CSCMCheckPoint::CSCMCheckPoint(CServiceModule& ServiceModule) : 
	m_dwNextWaitIntervalms(500),
	m_ServiceModule(ServiceModule),
	m_dwWaitHintPending(0)
{
}

CServiceModule::CSCMCheckPoint::~CSCMCheckPoint()
{
}

bool CServiceModule::CSCMCheckPoint::StartUp()
{
	ANYTHREADLOG_THIS_INOUT(_T("CSCMCheckPoint::StartUp\n"));
	// Keep the SCM happy for a pre-defined amount of time
	// set the repeat count of calling SetServiceStatus
	switch (m_ServiceModule.m_status.dwCurrentState)
	{
		case SERVICE_START_PENDING:
			m_dwWaitHintPending = s_dwWaitHintStartPending;
			SetRepeatCount(m_dwWaitHintPending / m_dwNextWaitIntervalms);
			break;
		case SERVICE_STOP_PENDING:
			m_dwWaitHintPending = s_dwWaitHintStopPending;
			SetRepeatCount(m_dwWaitHintPending / m_dwNextWaitIntervalms);
			break;
		default:
			_ASSERTE(0); 
	}
    m_ServiceModule.m_status.dwCheckPoint = 0;
	return CScheduleManager::StartUp();
}

void CServiceModule::CSCMCheckPoint::CloseDown()
{
	ANYTHREADLOG_THIS_INOUT(_T("CSCMCheckPoint::CloseDown\n"));
    m_ServiceModule.m_status.dwCheckPoint = 0;
	CScheduleManager::CloseDown();
}

void CServiceModule::CSCMCheckPoint::OnEventSchedule()
{
	ANYTHREADLOG_THIS(_T("CSCMCheckPoint::OnEventSchedule - checkpoint %d, wait hint %d\n"), m_ServiceModule.m_status.dwCheckPoint, m_ServiceModule.m_status.dwWaitHint);
    m_ServiceModule.m_status.dwCheckPoint++;
	if (m_dwWaitHintPending > m_dwNextWaitIntervalms)
	{
		m_dwWaitHintPending -= m_dwNextWaitIntervalms;
	}
	else
	{
		m_dwWaitHintPending = 0;
	}
	m_ServiceModule.m_status.dwWaitHint = m_dwWaitHintPending;
	VERIFY(::SetServiceStatus(m_ServiceModule.m_hServiceStatus, &m_ServiceModule.m_status));
}

///////////////////////////////////////////////////////////////////////////////

CServiceModule::CServiceModule() :
	m_SCMCheckPoint(*this),
	m_bExecutingRegistrationCode(FALSE)
{
}

CServiceModule::~CServiceModule()
{
}

// Although some of these functions are big they are declared inline since they are only used once

inline HRESULT CServiceModule::RegisterServer(BOOL bRegTypeLib, BOOL bService)
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
        return hr;

    // Remove any previous service since it may point to
    // the incorrect file
    Uninstall();

    // Add service entries
    UpdateRegistryFromResource(IDR_MessageQueueListener, TRUE);

    // Adjust the AppID for Local Server or Service
    CRegKey keyAppID;
    LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_WRITE);
    if (lRes != ERROR_SUCCESS)
        return lRes;

    CRegKey key;
    lRes = key.Open(keyAppID, _T("{2B0E56B2-4B55-11D4-8237-005004E8D1A7}"), KEY_WRITE);
    if (lRes != ERROR_SUCCESS)
        return lRes;
    key.DeleteValue(_T("LocalService"));
    
    if (bService)
    {
        key.SetValue(_T("MessageQueueListener"), _T("LocalService"));
        key.SetValue(_T("-Service"), _T("ServiceParameters"));
        // Create service
        Install();
    }

    // Add object entries
    hr = CComModule::RegisterServer(bRegTypeLib);

    CoUninitialize();
    return hr;
}

inline HRESULT CServiceModule::UnregisterServer()
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
        return hr;

    // Remove service entries
    UpdateRegistryFromResource(IDR_MessageQueueListener, FALSE);
    // Remove service
    Uninstall();
    // Remove object entries
    CComModule::UnregisterServer(TRUE);
    CoUninitialize();
    return S_OK;
}

inline void CServiceModule::Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID, const GUID* plibid)
{
    CComModule::Init(p, h, plibid);

    m_bService = TRUE;

    LoadString(h, nServiceNameID, m_szServiceName, sizeof(m_szServiceName) / sizeof(TCHAR));

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
    LONG l = CComModule::Unlock();
    if (l == 0 && !m_bService)
        PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
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


// AD 28/08/01 - put in exception handler
inline BOOL CServiceModule::Install()
{

	enum
	{
		eInitialise,
		eOpenSC,
		eReadDependants,
		eCreateService,
		eNULL,
	} eAction = eNULL;

    try
	{

		eAction = eInitialise;

		if (IsInstalled())
		    return TRUE;

		eAction = eOpenSC;
	    SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
		if (hSCM == NULL)
	    {
		    MessageBox(NULL, _T("Couldn't open service manager"), m_szServiceName, MB_OK);
			return FALSE;
	    }

	    // Get the executable file path
		TCHAR szFilePath[_MAX_PATH];
	    ::GetModuleFileName(NULL, szFilePath, _MAX_PATH);


		// AD add in the reading of the startup xml so we can pass in a list of dependant services
		bstr_t bstrDependants;
		LPTSTR lpszDependants;


		eAction = eReadDependants;
		bstrDependants = ReadDependantServices();

#ifdef _UNICODE	
	USES_CONVERSION; 

		lpszDependants = OLE2T(bstrDependants);
#else
		lpszDependants = T2A(bstrDependants);
#endif

	    LogEventInfo(_T("Service installed with the following dependencies : %s"),lpszDependants);


		LPTSTR lpD, lpTemp;
#ifdef _UNICODE	
	    lpTemp = lpD = lpszDependants;
#else
	    lpTemp = lpD = (char*)lpszDependants;
#endif
		while (*lpTemp)
		{ 
			if (*lpTemp == _T(','))
			{ 
				*lpD++ = _T('\0');
			} 
			else 
			{
				*lpD++ = *lpTemp;
			} 
			lpTemp++;
		} 

		eAction = eCreateService;
		
		SC_HANDLE hService = ::CreateService(
	        hSCM, m_szServiceName, m_szServiceName,
		    SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS | SERVICE_INTERACTIVE_PROCESS ,
			SERVICE_AUTO_START, SERVICE_ERROR_NORMAL,
	        szFilePath, NULL, NULL, lpszDependants, NULL, NULL);

		if (hService == NULL)
	    {
		    ::CloseServiceHandle(hSCM);
			MessageBox(NULL, _T("Couldn't create service"), m_szServiceName, MB_OK);
	        return FALSE;
		}

	    ::CloseServiceHandle(hService);
		::CloseServiceHandle(hSCM);
	}

	
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eInitialise:
		        LogEventError(_T("Install - Unable to initialise. "));
				break;

			case eOpenSC:
		        LogEventError(_T("Install - Unable to open the Service Control Manager"));
				break;

			case eReadDependants:
		        LogEventError(_T("Install - Error processing the dependants list in MessageQueueListener.xml"));
				break;

			case eCreateService:
		        LogEventError(_T("Install - Unable to create the Service"));
				break;

			default :
		        LogEventError(_T("Install - Exception caught."));
				break;
		}

	
	}
	catch(...)
	{
        LogEventError(_T("Install - Exception caught."));
	}


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

    if (bDelete == FALSE)
	{
		MessageBox(NULL, _T("Service could not be deleted"), m_szServiceName, MB_OK);
        return FALSE;
	}

    return TRUE;
}

///////////////////////////////////////////////////////////////////////////////////////
// Logging functions

void CServiceModule::LogEvent(WORD wEventType, DWORD dwEventId, LPCTSTR* ppszFormatplusArgs)
{
    try
	{
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		HANDLE  hEventSource;
		LPTSTR  lpszStrings[1];
		va_list pArg;

		va_start(pArg, *ppszFormatplusArgs);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, *ppszFormatplusArgs, pArg) > -1);
		va_end(pArg);

		// write to event log
		switch (wEventType)
		{
			case EVENTLOG_ERROR_TYPE:
			{
				ANYTHREADLOG_THIS(_T("LogEvent EVENTLOG_ERROR_TYPE - %s\n"), chMsg);
				break;
			}
			case EVENTLOG_INFORMATION_TYPE:
			{
				ANYTHREADLOG_THIS(_T("LogEvent EVENTLOG_INFORMATION_TYPE - %s\n"), chMsg);
				break;
			}
			case EVENTLOG_WARNING_TYPE:
			{
				ANYTHREADLOG_THIS(_T("LogEvent EVENTLOG_WARNING_TYPE - %s\n"), chMsg);
				break;
			}
			default :
				break;
		}
		
		lpszStrings[0] = chMsg;

		// prevent more than one thread reporting an event at the same time
		namespaceMutex::CSingleLock lckEventSource(&m_csEventSource, TRUE);

		/* Get a handle to use with ReportEvent(). */
		hEventSource = RegisterEventSource(NULL, m_szServiceName);
		if (hEventSource != NULL)
		{
			/* Write to event log. */
			ReportEvent(hEventSource, wEventType, 0, dwEventId, NULL, 1, 0, (LPCTSTR*) &lpszStrings[0], NULL);
			DeregisterEventSource(hEventSource);
		}
	}
	catch(...)
	{
	}
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
    VERIFY(SetServiceStatus(SERVICE_START_PENDING, s_dwWaitHintStartPending));

    m_status.dwWin32ExitCode = S_OK;
    m_status.dwCheckPoint = 0;
    m_status.dwWaitHint = 0;

    // When the Run function returns, the service has stopped.
    Run();

	VERIFY(SetServiceStatus(SERVICE_STOPPED));
}

inline void CServiceModule::Handler(DWORD dwOpcode)
{
	ANYTHREADLOG_THIS_INOUT(_T("Service Handler Opcode %d\n"), dwOpcode);	

    switch (dwOpcode)
    {
    case SERVICE_CONTROL_STOP:
        VERIFY(SetServiceStatus(SERVICE_STOP_PENDING, s_dwWaitHintStopPending));
        PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
        break;
    case SERVICE_CONTROL_PAUSE:
        break;
    case SERVICE_CONTROL_CONTINUE:
        break;
    case SERVICE_CONTROL_INTERROGATE:
		VERIFY(::SetServiceStatus(m_hServiceStatus, &m_status));
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

BOOL CServiceModule::SetServiceStatus(DWORD dwState, DWORD dwWaitHint)
{
    m_status.dwCurrentState = dwState;
	m_status.dwWaitHint = dwWaitHint;
	BOOL bResult = ::SetServiceStatus(m_hServiceStatus, &m_status);
    VERIFY(bResult);
	return bResult;
}

void CServiceModule::Run()
{
	_Module.dwThreadID = GetCurrentThreadId();

    HRESULT hr = CoInitialize(NULL);
//  If you are running on NT 4.0 or higher you can use the following call
//  instead to make the EXE free threaded.
//  This means that calls come in on a random RPC thread
//  HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (!SUCCEEDED(hr) )
		LogEventError(_T("Service failed to start - CoInitialize failed"));

	_ASSERTE(SUCCEEDED(hr));

    // This provides a NULL DACL which will allow access to everyone.
    CSecurityDescriptor sd;
    sd.InitializeFromThreadToken();
    hr = CoInitializeSecurity(sd, -1, NULL, NULL,
        RPC_C_AUTHN_LEVEL_PKT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    _ASSERTE(SUCCEEDED(hr));

	if (!SUCCEEDED(hr) )
		LogEventError(_T("Service failed to start - CoInitializeSecurity failed"));

	hr = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, REGCLS_MULTIPLEUSE);
    _ASSERTE(SUCCEEDED(hr));

	if (!SUCCEEDED(hr) )
		LogEventError(_T("Service failed to start - RegisterClassObjects failed"));


//	// helper to debug the service	
//	#ifdef _DEBUG
//		if (m_bService)
//		{
//			DebugBreak();
//		}
//	#endif

	try
	{
		if (m_bService)
		{
			// Read any configurable wait hints from the configuration file
			ReadSCMWaitHints();

			// start a SCM checkpoint thread to keep the SCM happy during a lengthy startup
			m_SCMCheckPoint.StartUp();
		}					

//		delete g_pMpHeap;
//		g_pMpHeap = new namespaceMemMan::CMpHeap;
//		g_pMpHeap->Initialise();

		_ASSERTE(g_pPrfMonMessageQueueListener == NULL);
		g_pPrfMonMessageQueueListener = new CPrfMonMessageQueueListener;

		BOOL bInstalled;
		{
			ANYTHREADLOG_THIS_INOUT(_T("Install performance counters\n"));	
			bInstalled = g_pPrfMonMessageQueueListener->Install();
		}
		if (bInstalled == FALSE)
		{
			LogEventError(NULL, _T("Couldn't install performance counters"), m_szServiceName, MB_OK);
		}
		else
		{
			const BOOL bPrfMonEnabled = TRUE;
			LPCSTR sModuleFile = "MessageQueueListener";
			const BOOL bResetCountersOnIdle = TRUE;

			if (g_pPrfMonMessageQueueListener->Initialise(bPrfMonEnabled, sModuleFile, bResetCountersOnIdle))
			{
				// create and start the message queuelistener
				CComPtr<IInternalMessageQueueListener1> MessageQueueListener1Ptr;
				HRESULT hr = CoCreateInstance(CLSID_MessageQueueListener1, NULL, CLSCTX_LOCAL_SERVER, IID_IInternalMessageQueueListener1, (void**)&MessageQueueListener1Ptr);

				if (MessageQueueListener1Ptr && SUCCEEDED(MessageQueueListener1Ptr->Start()))
				{
					_bstr_t bstrVersion = GetVersion();
					LogEventInfo(_T("Service started (Version %s)"), (LPCTSTR)bstrVersion);
					if (m_bService)
					{
						// done with the SCM checkpoint
						m_SCMCheckPoint.CloseDown();
						// service is now running
						VERIFY(SetServiceStatus(SERVICE_RUNNING));
					}

					// test transactional status of supporting MTS DLL
					TestMessageQueueListenerMTS1();

					// AD - check for xml file and load up parameters
					MessageQueueListener1Ptr->LoadConfigurationDetails();

					MSG msg;
					while (GetMessage(&msg, 0, 0, 0))
						DispatchMessage(&msg);

					if (m_bService)
					{
						// start a SCM checkpoint thread to keep the SCM happy during a lengthy shutdown 
						m_SCMCheckPoint.StartUp();
					}					
					LogEventInfo(_T("Service stopping"));
					MessageQueueListener1Ptr->Stop();
				}
				else
				{
					LogEventError(_T("Service failed to start"));
				}
			}
			else
			{
				LogEventError(_T("Failed to initialize performance counters"));
			}
		}
	}
	catch(...)
	{
		LogEventError(_T("Service failed to start - caught exception"));
	}

	
	if (g_pPrfMonMessageQueueListener)
	{
		ANYTHREADLOG_THIS_INOUT(_T("Uninstall performance counters\n"));	
		g_pPrfMonMessageQueueListener->Exit();
		BOOL bUnInstalled = g_pPrfMonMessageQueueListener->UnInstall();
		if (bUnInstalled == FALSE)
		{
			MessageBox(NULL, _T("Couldn't un-install performance counters"), m_szServiceName, MB_OK);
		}
		delete g_pPrfMonMessageQueueListener;
		g_pPrfMonMessageQueueListener = NULL;
	}

//	if (g_pMpHeap)
//	{
//		delete g_pMpHeap;
//		g_pMpHeap = NULL;
//	}

    if (m_bService)
	{
		// done with the SCM checkpoint
		m_SCMCheckPoint.CloseDown();
	}

    LogEventInfo(_T("Service stopped"));
	_Module.RevokeClassObjects();
    CoUninitialize();
}

void CServiceModule::TestMessageQueueListenerMTS1()
{
	ANYTHREADLOG_THIS_INOUT(_T("TestMessageQueueListenerMTS1\n"));
	const LPCTSTR pszFunctionName = _T("TestMessageQueueListenerMTS1");

	enum
	{
		eCreate,
		eIsInsideMTS,
		eIsInMTSTransaction,
		eNULL,
	} eAction = eNULL;

	try
	{
		eAction = eCreate;
		IInternalMessageQueueListenerMTS1Ptr InternalMessageQueueListenerMTS1Ptr(__uuidof(MessageQueueListenerMTS1));
		if (InternalMessageQueueListenerMTS1Ptr)
		{
			eAction = eIsInsideMTS;
			BOOL bIsInsideMTS = InternalMessageQueueListenerMTS1Ptr->GetIsInsideMTS(::GetCurrentThreadId());
			if (bIsInsideMTS)
			{
				eAction = eIsInMTSTransaction;
				BOOL bIsInMTSTransaction = InternalMessageQueueListenerMTS1Ptr->GetIsInMTSTransaction(::GetCurrentThreadId());
				if (bIsInMTSTransaction)
				{
					LogEventInfo(_T("MessageQueueListenerMTS1 is running inside of MTS, and its MTS properties are set to support distributed transactions"));				
				}
				else
				{
					LogEventWarning(_T("MessageQueueListenerMTS1 is running inside of MTS, and its MTS properties are set to NOT support distributed transactions"));				
				}
			}
			else
			{
				LogEventWarning(_T("MessageQueueListenerMTS1 is not running inside of MTS, and does NOT support distributed transactions"));				
			}
		}
		else
		{
			LogEventError(_T("Service failed to create component or query interface IMessageQueueListenerMTS1"));
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eCreate:
				LogEventError(_T("Error creating MessageQueueListenerMTS1 for MTS/COMPLUS transaction test (ThreadId = %d, HResult = %d, Source = '%s', Description = '%s')"), GetCurrentThreadId(), comerr.Error(), (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
				break;
			case eIsInsideMTS:
				LogEventError(_T("Error testing whether MessageQueueListenerMTS1 is running inside of MTS/COMPLUS (ThreadId = %d, HResult = %d, Source = '%s', Description = '%s')"), GetCurrentThreadId(), comerr.Error(), (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
				break;
			case eIsInMTSTransaction:
				LogEventError(_T("Error testing whether MessageQueueListenerMTS1 is running inside a MTS/COMPLUS transaction (ThreadId = %d, HResult = %d, Source = '%s', Description = '%s')"), GetCurrentThreadId(), comerr.Error(), (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
				break;
			default:
				LogEventError(_T("Unknown error (ThreadId = %d, HResult = %d, Source = '%s', Description = '%s')"), GetCurrentThreadId(), comerr.Error(), (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
				break;
		}
	}
	catch(...)
	{
		LogEventError(_T("Unknown error"));
	}
}

_bstr_t CServiceModule::GetXmlConfigurationFileName()
{
	TCHAR szModuleFileName[_MAX_PATH]; 
	GetModuleFileName(NULL, szModuleFileName, _MAX_PATH); 

	USES_CONVERSION;
	string sModuleFileName = T2A(szModuleFileName);
	string sModulePath, sXmlFileName;
	string::size_type nPos = 0;
	if ((nPos = sModuleFileName.rfind('\\')) < sModuleFileName.npos)
	{
		sModulePath = sModuleFileName.substr(0, nPos);

		sXmlFileName = sModulePath + '\\' + "MessageQueueListener.xml";
	}
	return sXmlFileName.c_str();
}

bool CServiceModule::XmlConfigurationFileExists()
{
	bool bExists = false;

	FILE *fTest;
	fTest = fopen(GetXmlConfigurationFileName(), "r");
	if (fTest != NULL)
	{
		bExists = true;
		fclose(fTest);
	}
	return bExists;
}

_bstr_t CServiceModule::ReadDependantServices()
{
	enum
	{
		eInitialise,
		eOpenFile,
		eLoadXml,
		eGetRootElement,
		eParseDom,
		eNULL,
	} eAction = eNULL;
	
	bstr_t bstrDepends("RPCSS,MessageQueueListenerLOG"); // always have to have this
	
	try
	{
		eAction = eInitialise;
		bstrDepends += ",";

		// RF 04/02/04
		//XmlNS::IXMLDOMDocumentPtr	ptrDomDocument(__uuidof(XmlNS::DOMDocument));	
		XmlNS::IXMLDOMDocument2Ptr	ptrDomDocument(__uuidof(XmlNS::DOMDocument40));	
		//XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));	
		XmlNS::IXMLDOMDocument2Ptr	ptrDomOut(__uuidof(XmlNS::DOMDocument40));	
		XmlNS::IXMLDOMElementPtr	ptrElement;
		XmlNS::IXMLDOMNodeListPtr	ptrElementList;
		XmlNS::IXMLDOMNodePtr		ptrNode, ptrNode2;
		XmlNS::IXMLDOMNodeListPtr	ptrQueueElementList;

		eAction = eOpenFile;		
		if (XmlConfigurationFileExists())
		{
			eAction = eLoadXml;
			
			// RF 04/02/04 Start
			bstr_t bstrPropertyName("NewParser");
			ptrDomDocument->setProperty(bstrPropertyName, VARIANT_TRUE); 
			// RF 04/02/04 End

			if (ptrDomDocument->load(GetXmlConfigurationFileName()) == VARIANT_TRUE)
			{
				eAction = eGetRootElement;
				ptrElement = ptrDomDocument->GetdocumentElement();

				// root element is CONFIGURATION
				bstr_t bstrElement1("INSTALL");
				ptrElementList = ptrDomDocument->getElementsByTagName(bstrElement1);
				long lInList = ptrElementList->Getlength();

				if (lInList == 1)
				{
					eAction = eParseDom;
					
					ptrNode=ptrElementList->Getitem(0);
	
					ptrDomOut->PutRefdocumentElement((XmlNS::IXMLDOMElementPtr)ptrNode);
		
					ptrElement = ptrDomOut->GetdocumentElement();

					bstr_t bstrElement2("DEPENDANTSERVICES");
					ptrElementList = ptrDomOut->getElementsByTagName(bstrElement2);
					lInList = ptrElementList->Getlength();

					// hopefully should have at least 1
					if(lInList!=0) 
					{
						for(int i = 0; i<lInList; i++)
						{
							ptrNode=ptrElementList->Getitem(i);
							ptrNode->get_childNodes(&ptrQueueElementList);
							long lInList2 = ptrQueueElementList->Getlength();
							for(int i2 = 0; i2<lInList2; i2++)
							{
								ptrNode2 = ptrQueueElementList->Getitem(i2);
								bstr_t bstrnodestring = ptrNode2->Gettext();

								// add string to list
								bstrDepends += bstrnodestring;
								bstrDepends += ",";
							}
						}
					}
				}
			}
			else
			{
				// error loading xml
				LogEventError(_T("ReadConfigFile unable to process xml file."));
			}
		}

		// convert onto a LPCSTR
		bstrDepends += ",";
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eInitialise:
				// RF 04/02/04
		        //LogEventError(_T("ReadDependentServices - Unable to initialise. Check MSXML3 is installed correctly"));
				LogEventError(_T("ReadDependentServices - Unable to initialise. Check MSXML4 is installed correctly"));
				break;
			case eOpenFile:
		        LogEventError(_T("ReadDependentServices - Unable to open MessageQueueListener.xml"));
				break;
			case eLoadXml:
		        LogEventError(_T("ReadDependentServices - Unable to load MessageQueueListener.xml into the DOM"));
				break;
			case eGetRootElement:
		        LogEventError(_T("ReadDependentServices - Unable to locate root element"));
				break;
			case eParseDom:
		        LogEventError(_T("ReadDependentServices - Unable to unpack the service information"));
				break;
			default :
		        LogEventError(_T("ReadDependentServices - Error, unable to process xml file "));
				break;
		}
	}
	catch(...)
	{
        LogEventError(_T("ReadDependentServices - Error, unable to process xml file "));
	}

	return bstrDepends;
}

void CServiceModule::ReadSCMWaitHints()
{
	try
	{
		if (XmlConfigurationFileExists())
		{
			// extract data from request
			XmlNS::IXMLDOMDocumentPtr ConfigurationDocumentPtr(_uuidof(XmlNS::DOMDocument));
			ConfigurationDocumentPtr->Putasync(VARIANT_FALSE);
			if (ConfigurationDocumentPtr->load(GetXmlConfigurationFileName()) == VARIANT_TRUE)
			{
				XmlNS::IXMLDOMElementPtr ElementPtr;
				ElementPtr = ConfigurationDocumentPtr->selectSingleNode(L"/CONFIGURATION/SCM/WAITHINTSTARTPENDING");
				if (ElementPtr)
				{
					s_dwWaitHintStartPending = _wtoi(ElementPtr->text);
				}
				ElementPtr = ConfigurationDocumentPtr->selectSingleNode(L"/CONFIGURATION/SCM/WAITHINTSTOPPENDING");
				if (ElementPtr)
				{
					s_dwWaitHintStopPending = _wtoi(ElementPtr->text);
				}
			}
			else
			{
		        LogEventError(_T("ReadSCMWaitHints - Error, unable to process xml file "));
			}
		}
	}
	catch(_com_error comerr)
	{
		LogEventError(_T("ReadSCMWaitHints - Error, unable to process xml file "));
	}
	catch(...)
	{
		LogEventError(_T("ReadSCMWaitHints - Error, Unknown error"));
	}
}

_bstr_t CServiceModule::GetVersion()
{
	_bstr_t bstrVersion;
	
	TCHAR szFileName[_MAX_PATH];
	GetModuleFileName(NULL, szFileName, _MAX_PATH);
  
	DWORD dw;
	DWORD size = GetFileVersionInfoSize(szFileName, &dw);
	TCHAR* pVer = new TCHAR[size];
	if(GetFileVersionInfo(szFileName, 0, size, pVer))
	{
		LPVOID pBuff;
		UINT cbBuff;
		VerQueryValue(pVer, _T("\\StringFileInfo\\040904B0\\FileVersion"), &pBuff, &cbBuff); 
		bstrVersion = (LPCWSTR)pBuff;
		
	}
	delete[] pVer;

	return bstrVersion;
}

/////////////////////////////////////////////////////////////////////////////
//
extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, 
    HINSTANCE /*hPrevInstance*/, LPTSTR lpCmdLine, int /*nShowCmd*/)
{
    lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT
    _Module.Init(ObjectMap, hInstance, IDS_SERVICENAME, &LIBID_MESSAGEQUEUELISTENERLib);
    _Module.m_bService = TRUE;

    TCHAR szTokens[] = _T("-/");

    LPCTSTR lpszToken = FindOneOf(lpCmdLine, szTokens);
    while (lpszToken != NULL)
    {
        // Unregister
        if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
		{
            _Module.m_bExecutingRegistrationCode = TRUE;
			return _Module.UnregisterServer();
		}

        // Register as Local Server
        if (lstrcmpi(lpszToken, _T("LocalServer"))==0 || lstrcmpi(lpszToken, _T("Console"))==0)
		{
            _Module.m_bExecutingRegistrationCode = TRUE;
            return _Module.RegisterServer(TRUE, FALSE);
		}
        
        // Register as Service
        if (lstrcmpi(lpszToken, _T("Service"))==0 || lstrcmpi(lpszToken, _T("RegServer"))==0)
		{
            _Module.m_bExecutingRegistrationCode = TRUE;
            return _Module.RegisterServer(TRUE, TRUE);
		}
        
        lpszToken = FindOneOf(lpszToken, szTokens);
    }
    
	// Are we Service or Local Server
    CRegKey keyAppID;
    LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_READ);
    if (lRes != ERROR_SUCCESS)
        return lRes;

    CRegKey key;
    lRes = key.Open(keyAppID, _T("{2B0E56B2-4B55-11D4-8237-005004E8D1A7}"), KEY_READ);
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


