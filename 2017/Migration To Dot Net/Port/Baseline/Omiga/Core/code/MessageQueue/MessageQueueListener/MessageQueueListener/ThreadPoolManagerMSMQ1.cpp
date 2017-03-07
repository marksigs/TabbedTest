///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerMSMQ1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		19/03/01	SYS2103 - Make MSMQ1 queue tolerant to database shutdown/startup.
//	LD		06/04/01	SYS2248 - Add performance counters
//	LD		09/04/01	SYS2249 - Corrections for two listeners on the same queue
//	LD		12/09/01    SYS2707 - Serialize access to public methods
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		07/12/01	SYS3421 - Create MSMQ objects on a new thread
//	LD		26/12/01	SYS3917	- Events not always reenabled after ArrivedError/RestartMSMQ
//	LD		21/03/02	SYS4414 - Increase peek time out and enhanced logging.
//								- m_lTotalMessagesSkippedBecauseStalled introduced to manage messages to scan
//								- Additional logging
//	LD		15/05/02	SYS4618 - Make logging more robust
//								- Stop queue amendments
//  LD		22/08/02	SYS5464 - Wait for current messages being processed to stop
//  LD		19/03/03	Updates to support queuenames specified as a format name
//  LD					Introduce m_lTotalMessagesRolledback
//						Only access MSMQ objects from main creator thread
//  RF		24/02/04	BMIDS727 - Allow for multiple queue tables
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <process.h>
#include "mutex.h"
using namespace namespaceMutex;
#include "ThreadPoolManagerMSMQ1.h"
#include "ThreadPoolThread.h"
#include "PrfMonMessageQueueListener.h"
#include "Log.h"

#import "..\MessageQueueListenerMTS\MessageQueueListenerMTS.tlb" no_namespace
#import "..\MessageQueueComponentVC\MessageQueueComponentVC.tlb" no_namespace

const long s_lPeekTimeOut = 60000L; // in ms
const long s_lEnableNotificationTimeOut = INFINITE; // in ms
const DWORD s_dwPollMSMQService = 5000; // in ms when waiting for MSMQ to become available

UINT CThreadPoolManagerMSMQ1::WM_MSMQ1_THREADAVAILABLEFORREUSE = RegisterWindowMessage(_T("WM_MSMQ1_THREADAVAILABLEFORREUSE"));
UINT CThreadPoolManagerMSMQ1::WM_MSMQ1_CLOSEDOWN = RegisterWindowMessage(_T("WM_MSMQ1_CLOSEDOWN"));
UINT CThreadPoolManagerMSMQ1::WM_MSMQ1_RESET = RegisterWindowMessage(_T("WM_MSMQ1_RESET"));
UINT CThreadPoolManagerMSMQ1::WM_MSMQ1_RESTART = RegisterWindowMessage(_T("WM_MSMQ1_RESTART"));


#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MSMQ1); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MSMQ1); LogInOut.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA CLogIn LogIn(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MSMQ1); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA_INOUT CLogInOut LogInOut(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MSMQ1); LogInOut.AnyThreadInitialize

/////////////////////////////////////////////////////////////////////////////
// CThreadPoolManagerMSMQ1

CThreadPoolManagerMSMQ1::CThreadPoolManagerMSMQ1() :
	m_dwCookie(0),
	m_ePeek(ePeekCurrentToDo),
	m_bNotificationEnabled(FALSE),
	m_bQueueStalled(FALSE),
	m_bQueueStarted(FALSE),
	m_pPrfInstanceQueueInfo(NULL),
	m_bstrQueueName(L"Notdefined"),
	m_dwCreatorThreadId(0),
	m_hCreatorThread(NULL),
	m_hEventCreatedOrError(NULL),
	m_hEventCloseDown(NULL),
	m_lTotalMessagesSkippedBecauseStalled(0),
	m_lTotalMessagesRolledback(0)
{

}

CThreadPoolManagerMSMQ1::~CThreadPoolManagerMSMQ1()
{

}

HRESULT CThreadPoolManagerMSMQ1::FinalConstruct()
{
	return CComObjectRootEx<CComMultiThreadModel>::FinalConstruct();
}

void CThreadPoolManagerMSMQ1::FinalRelease()
{
	CComObjectRootEx<CComMultiThreadModel>::FinalRelease();
}

///////////////////////////////////////////////////////////////////////////////
// IMSMQEventSink

STDMETHODIMP CThreadPoolManagerMSMQ1::ArrivedError(IDispatch* pQueue, long lErrorCode, long lCursor)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s ArrivedError\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("ArrivedError");

	if (m_bQueueStarted) // queue not stopping
	{
		try
		{
			m_bNotificationEnabled = FALSE;
			switch (lErrorCode)
			{
				case MQ_ERROR_SERVICE_NOT_AVAILABLE: // fall-through
				case MQ_ERROR_INVALID_HANDLE:
				{
					// close and reopen the queue if the service is unavailable
					LogEventError(pszFunctionName, _T("(ErrorCode = %d).  Waiting for MSMQ service to become available"), lErrorCode);
					bool bMSMQRestarted = false;
					while (m_bQueueStarted == TRUE && bMSMQRestarted == false)
					{
						try
						{
							CreatorThreadCloseDownMSMQ();
							const bool bReportErrors = false;
							bMSMQRestarted = CreatorThreadStartUpMSMQ(bReportErrors);
							if (!bMSMQRestarted)
							{
								::Sleep(s_dwPollMSMQService);
							}
						}
						catch(...)
						{
						}
					}
					if (bMSMQRestarted)
					{
						LogEventInfo(pszFunctionName, _T("MSMQ service is now available"));
					}
					break;
				}
				case MQ_ERROR_MESSAGE_ALREADY_RECEIVED:
				{
					CreatorThreadEnableEvents(MQMSG_NEXT);
					break;
				}
				default :
					LogEventError(pszFunctionName, _T("(ErrorCode = %d)"), lErrorCode);
					CreatorThreadEnableEvents(MQMSG_CURRENT);
					break;
			}
		}
		catch(_com_error comerr)
		{
			LogEventError(pszFunctionName, comerr, _T("Unknown error "));
		}
		catch(...)
		{
			LogEventError(pszFunctionName, _T("Unknown error"));
		}
	}
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::Arrived(IDispatch* pQueue, long lCursor)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s Arrived\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("Arrived");

	if (m_bQueueStarted) // queue not stopping
	{
		m_bNotificationEnabled = FALSE;
		CreatorThreadDispatchMessageToWorkerThread();
	}
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
// IThreadPoolManagerCommon1

STDMETHODIMP CThreadPoolManagerMSMQ1::put_Started(BOOL newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_Started\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_Started");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	HRESULT hRes = S_FALSE;

	m_bQueueStarted = newVal;
	if(m_bQueueStarted == TRUE)
	{
		LogEventInfo(pszFunctionName, _T("Starting to listen to queue '%s' with %d threads"), (LPWSTR)m_bstrQueueName, GetnRequestedNumberOfThreads());

		// create msmq objects on a new thread
		// ... Create an event, which indicates creation of queue objects
		_ASSERT(m_hEventCreatedOrError == NULL);
		m_hEventCreatedOrError = CreateEvent(NULL, FALSE, FALSE, NULL);
		if (m_hEventCreatedOrError)
		{
			// ... create a new thread on which to create MSMQ objects
			_ASSERT(m_hCreatorThread == NULL);
			m_dwCreatorThreadId = NULL;
			m_hCreatorThread = ::CreateThread(NULL, 0, StartUponThread, this, 0, &m_dwCreatorThreadId);

			// ... and wait for them to be created
			DWORD dwWait = WaitForSingleObject(m_hEventCreatedOrError, INFINITE);
			switch (dwWait)
			{
			case WAIT_OBJECT_0: // objects created or an error has occurred
				break;
				
			case WAIT_TIMEOUT: // timed out
				break;

			case WAIT_ABANDONED:
				break;
			}
			CloseHandle(m_hEventCreatedOrError);
			m_hEventCreatedOrError = NULL;

			if (m_bQueueStarted)
			{
				hRes = S_OK;
			}
		}
		else
		{
			m_bQueueStarted = FALSE;
		}
	}
	else
	{
		LogEventInfo(pszFunctionName, _T("Stopping to listen to queue '%s'"), (LPWSTR)m_bstrQueueName);

		_ASSERT(m_hEventCloseDown == NULL);
		m_hEventCloseDown = CreateEvent(NULL, FALSE, FALSE, NULL);
		if (m_hEventCloseDown)
		{
			_ASSERT(m_hCreatorThread != NULL);
			// let the thread on which the msmq objects were created quit
			PostThreadMessage(m_dwCreatorThreadId, WM_MSMQ1_CLOSEDOWN, 0, 0);

			// ... and wait for them to be destroyed
			DWORD dwWait = WaitForSingleObject(m_hEventCloseDown, INFINITE);
			switch (dwWait)
			{
			case WAIT_OBJECT_0: // objects destroyed an error has occurred
				///??
				break;
				
			case WAIT_TIMEOUT: // timed out
				break;

			case WAIT_ABANDONED:
				break;
			}
			CloseHandle(m_hEventCloseDown);
			m_hEventCloseDown = NULL;

			CloseHandle(m_hCreatorThread);
			m_hCreatorThread = NULL;
			m_dwCreatorThreadId = 0; 

			if (!m_bQueueStarted)
			{
				hRes = S_OK;
			}
		}
		else
		{
			// do nothing
		}
	}

	return hRes;
}


STDMETHODIMP CThreadPoolManagerMSMQ1::get_Started(BOOL *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_Started\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bQueueStarted;
	return S_OK;
}


STDMETHODIMP CThreadPoolManagerMSMQ1::get_NumberOfThreads(long *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_NumberOfThreads\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = GetnRequestedNumberOfThreads();
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_NumberOfThreads(long newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_NumberOfThreads (newVal = %ld)\n"), (LPWSTR)m_bstrQueueName, newVal);
	const LPCTSTR pszFunctionName = _T("put_NumberOfThreads");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	if (newVal != GetnRequestedNumberOfThreads())
	{
		if (m_bQueueStarted)
		{
			LogEventInfo(pszFunctionName, _T("Changing the number of threads to %d on queue '%s'"), newVal, (LPWSTR)m_bstrQueueName);
		}
		SetnRequestedNumberOfThreads(newVal);
	}
	return S_OK;
}

// RF 24/02/04
STDMETHODIMP CThreadPoolManagerMSMQ1::get_TableSuffix(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_TableSuffix\n"), (LPWSTR)m_bstrTableSuffix);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bstrTableSuffix.copy();
	return S_OK;
}

// RF 24/02/04
STDMETHODIMP CThreadPoolManagerMSMQ1::put_TableSuffix(BSTR newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_TableSuffix\n"), (LPWSTR)m_bstrTableSuffix);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bstrTableSuffix = newVal;
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::get_QueueName(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueName\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bstrQueueName.copy();
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_QueueName(BSTR newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_QueueName\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_QueueName");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bstrQueueName = newVal;
	m_bstrMoveQueueName = m_bstrQueueName + L"X";

	_bstr_t bstrQueueName = _tcsupr(m_bstrQueueName);
	if (_tcsstr((LPCWSTR)bstrQueueName, L"PUBLIC=") != NULL ||
		_tcsstr((LPCWSTR)bstrQueueName, L"PRIVATE=") != NULL ||
		_tcsstr((LPCWSTR)bstrQueueName, L"DIRECT=") != NULL ) 
	{
		// format name passed in
		// note that appending an X for the move format name will not work if queue guids or queue
		// numbers are used rather than queue names
		m_bstrFormatName = m_bstrQueueName;
		m_bstrMoveFormatName = m_bstrMoveQueueName;
	}
	else
	{
		// path name passed in, so look up the format names
		// ... m_bstrFormatName
		try
		{
			IMSMQQueueInfoPtr MSMQQueueInfoPtr(__uuidof(MSMQQueueInfo));
			MSMQQueueInfoPtr->put_PathName(m_bstrQueueName);
			MSMQQueueInfoPtr->Refresh();
			m_bstrFormatName = MSMQQueueInfoPtr->FormatName;
		}
		catch(_com_error comerr)
		{
			LogEventError(pszFunctionName, comerr, _T("Error converting pathname (%s) to formatname"), (LPCWSTR)m_bstrQueueName);
		}
		catch(...)
		{
			LogEventError(pszFunctionName, _T("Error converting pathname (%s) to formatname"), (LPCWSTR)m_bstrQueueName);
		}

		// ... m_bstrMoveFormatName
		try
		{
			IMSMQQueueInfoPtr MSMQQueueInfoPtr(__uuidof(MSMQQueueInfo));
			MSMQQueueInfoPtr->put_PathName(m_bstrMoveQueueName);
			MSMQQueueInfoPtr->Refresh();
			m_bstrMoveFormatName = MSMQQueueInfoPtr->FormatName;
		}
		catch(_com_error comerr)
		{
			LogEventError(pszFunctionName, comerr, _T("Error converting pathname (%s) to formatname"), (LPCWSTR)m_bstrMoveQueueName);
		}
		catch(...)
		{
			LogEventError(pszFunctionName, _T("Error converting pathname to (%s) formatname"), (LPCWSTR)m_bstrMoveQueueName);
		}
	}

	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::get_QueueStalled(BOOL *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueStalled\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bQueueStalled;
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_QueueStalled(BOOL newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_QueueStalled\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_QueueStalled");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bQueueStalled = newVal;
	if (m_bQueueStalled)
	{
		LogEventWarning(pszFunctionName, _T("Stalling queue as requested due to an external request"));
	}
	else
	{
		LogEventInfo(pszFunctionName, _T("Unstalling queue as requested due to an external request"));
	}

	if (!m_bQueueStalled)
	{
		PostThreadMessage(m_dwCreatorThreadId, WM_MSMQ1_RESTART, 0, 0);
	}
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::get_ComponentsStalled(VARIANT *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_ComponentsStalled\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	return CThreadPoolManagerMQ::get_ComponentsStalled(pVal);
}


STDMETHODIMP CThreadPoolManagerMSMQ1::put_AddStalledComponents(VARIANT newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_AddStalledComponents\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_AddStalledComponents");
	CSingleLock lck(&m_csPublicMethod, TRUE);

	LogEventWarning(pszFunctionName, _T("Stalling components as requested due to an external request"));
	
	HRESULT hr = CThreadPoolManagerMQ::put_AddStalledComponents(newVal);
	m_pPrfInstanceQueueInfo->SetnDifferentStalledComponents(GetnDifferentComponentsStalled());
	return hr;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_RestartComponents(VARIANT newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_RestartComponents\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_RestartComponents");
	CSingleLock lck(&m_csPublicMethod, TRUE);

	LogEventInfo(pszFunctionName, _T("Unstalling components as requested due to an external request"));

	HRESULT hr = CThreadPoolManagerMQ::put_RestartComponents(newVal);
	m_pPrfInstanceQueueInfo->SetnDifferentStalledComponents(GetnDifferentComponentsStalled());
	PostThreadMessage(m_dwCreatorThreadId, WM_MSMQ1_RESTART, 0, 0);
	return hr;
}


STDMETHODIMP CThreadPoolManagerMSMQ1::get_QueueType(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueType\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = ::SysAllocString(L"MSMQ1");
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_ConnectionString(BSTR newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_ConnectionString\n"), (LPWSTR)m_bstrQueueName);
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerMSMQ1::put_dwNextWaitIntervalms(DWORD dwPollInterval)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_dwNextWaitIntervalms\n"), (LPWSTR)m_bstrQueueName);
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////

DWORD WINAPI CThreadPoolManagerMSMQ1::StartUponThread(void* pThreadPoolManagerMSMQ1)
{
	return reinterpret_cast<CThreadPoolManagerMSMQ1*>(pThreadPoolManagerMSMQ1)->StartUponThread();
}

DWORD WINAPI CThreadPoolManagerMSMQ1::StartUponThread()
{
	const LPCTSTR pszFunctionName = _T("StartUponThread");

	HRESULT hr = CoInitialize(NULL);
	if SUCCEEDED(hr)
	{
		enum
		{
			eNULL,
			eStartUp,
			eMessageLoop,
			eCloseDown
		} eAction = eNULL;
	
		try
		{
			eAction = eStartUp;
			m_bQueueStarted = CreatorThreadStartUp();
			SetEvent(m_hEventCreatedOrError); // signal startup complete or error through m_bQueueStarted

			eAction = eMessageLoop;
			MSG msg;
			while (GetMessage(&msg, 0, 0, 0))
			{
				if (msg.message == WM_MSMQ1_THREADAVAILABLEFORREUSE)
				{
					CreatorThreadDispatchMessageToWorkerThread();
				}
				else if (msg.message == WM_MSMQ1_CLOSEDOWN)
				{
					ANYTHREADLOG_THIS_INOUT(_T("%s WM_MSMQ1_CLOSEDOWN\n"), (LPWSTR)m_bstrQueueName);
					eAction = eCloseDown;
					CreatorThreadCloseDown();
					_ASSERT(m_hCreatorThread != NULL);
					// let the thread on which the msmq objects were created quit in its own time
					PostThreadMessage(m_dwCreatorThreadId, WM_QUIT, 0, 0);
				}
				else if (msg.message == WM_MSMQ1_RESET || msg.message == WM_MSMQ1_RESTART)
				{
					if (m_MSMQQueuePtr != NULL)
					{
						{
							ANYTHREADLOG_THIS_INOUT(_T("%s Reset\n"), (LPWSTR)m_bstrQueueName);
							m_MSMQQueuePtr->Reset();
							m_lTotalMessagesSkippedBecauseStalled = 0;
							m_lTotalMessagesRolledback = 0;
							m_ePeek = ePeekCurrentToDo;
						}
						CreatorThreadDispatchMessageToWorkerThread();
					}
				}
				else
				{
					DispatchMessage(&msg);
				}
			}
			ANYTHREADLOG_THIS_INOUT(_T("%s WM_QUIT\n"), (LPWSTR)m_bstrQueueName);
			m_bQueueStarted = FALSE;
			SetEvent(m_hEventCloseDown); // signal closedown complete or error through m_bQueueStarted
		}
		catch(_com_error comerr)
		{
			switch (eAction)
			{
				case eStartUp:
					LogEventError(pszFunctionName, comerr, _T("Error during startup"));
					SetEvent(m_hEventCreatedOrError); // signal startup complete or error through m_bQueueStarted
					break;
				case eMessageLoop:
					LogEventError(pszFunctionName, comerr, _T("Error during processing"));
					break;
				case eCloseDown:
					LogEventError(pszFunctionName, comerr, _T("Error closing down"));
					SetEvent(m_hEventCloseDown); // signal closedown complete or error through m_bQueueStarted
					break;
				default :
					LogEventError(pszFunctionName, comerr, _T("Unknown error"));
					break;
			}
		}
		catch(...)
		{
			switch (eAction)
			{
				case eStartUp:
					LogEventError(pszFunctionName, _T("Unknown error during startup"));
					SetEvent(m_hEventCreatedOrError); // signal startup complete or error through m_bQueueStarted
					break;
				case eMessageLoop:
					LogEventError(pszFunctionName, _T("Unknown error during processing"));
					break;
				case eCloseDown:
					LogEventError(pszFunctionName, _T("Unknown error closing down"));
					SetEvent(m_hEventCloseDown); // signal closedown complete or error through m_bQueueStarted
					break;
				default :
					LogEventError(pszFunctionName, _T("Unknown error"));
					break;
			}
		}
		CoUninitialize();
	}
	return 0;
}


bool CThreadPoolManagerMSMQ1::CreatorThreadStartUp()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadStartUp\n"), (LPWSTR)m_bstrQueueName);
	bool bResult = true;
	bResult = CreatorThreadStartUpThreadPool();
	if (bResult)
	{
		// open the queue
		bResult = CreatorThreadStartUpMSMQ();
		_ASSERTE(m_pPrfInstanceQueueInfo == NULL);
		m_pPrfInstanceQueueInfo = g_pPrfMonMessageQueueListener->AddInstanceQueueInfo(m_bstrQueueName);
	}
	return bResult;
}

bool CThreadPoolManagerMSMQ1::CreatorThreadStartUpThreadPool()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadStartUpThreadPool\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("StartUpThreadPool");

	bool bResult = true;

	enum
	{
		eNULL,
		eStartThreadPool,
	} eAction = eNULL;
	
	try
	{
		eAction = eStartThreadPool;
		bResult = CThreadPoolManager::StartUp();
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eStartThreadPool:
				LogEventError(pszFunctionName, comerr, _T("Error starting thread pool"));
				break;
			default :
				LogEventError(pszFunctionName, comerr, _T("Unknown error"));
				break;
		}
		bResult = false;
	}
	catch(...)
	{
		LogEventError(pszFunctionName, _T("Unknown error"));
	}

	return bResult;
}

bool CThreadPoolManagerMSMQ1::CreatorThreadStartUpMSMQ(bool bReportErrors)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadStartUpMSMQ\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("StartUpMSMQ");

	bool bResult = true;

	enum
	{
		eNULL,
		eOpenQueue,
		eCreateEventObject,
		eAtlAdvise,
		eEnableEvents
	} eAction = eNULL;
	
	try
	{
		// open the queue
		eAction = eOpenQueue;
		IMSMQQueueInfoPtr MSMQQueueInfoPtr(__uuidof(MSMQQueueInfo));
		MSMQQueueInfoPtr->put_FormatName(m_bstrFormatName);

		m_MSMQQueuePtr = MSMQQueueInfoPtr->Open(MQ_PEEK_ACCESS, MQ_DENY_NONE);
		m_lTotalMessagesSkippedBecauseStalled = 0;
		m_lTotalMessagesRolledback = 0;
		m_ePeek = ePeekCurrentToDo; // must peek at the current before peeking at the next

		// register for events
		// ... create event object
		eAction = eCreateEventObject;
		m_MSMQEventPtr.CreateInstance(__uuidof(MSMQEvent));
		// ... attach to connection point of event object
		m_DispatchPtr = this;
		eAction = eAtlAdvise;
		HRESULT hr = AtlAdvise(m_MSMQEventPtr, m_DispatchPtr, __uuidof(_DMSMQEventEvents), &m_dwCookie);
		_ASSERTE(SUCCEEDED(hr));

		// enable notification for queue using event object
		eAction = eEnableEvents;
		m_bNotificationEnabled = FALSE;
		CreatorThreadEnableEvents(MQMSG_FIRST);
	}
	catch(_com_error comerr)
	{
		if (bReportErrors)
		{
			switch (eAction)
			{
				case eOpenQueue:
					LogEventError(pszFunctionName, comerr, _T("Error opening queue"));
					break;
				case eCreateEventObject:
					LogEventError(pszFunctionName, comerr, _T("Error creating event object"));
					break;
				case eAtlAdvise:
					LogEventError(pszFunctionName, comerr, _T("Error advising MSMQ"));
					break;
				case eEnableEvents:
					LogEventError(pszFunctionName, comerr, _T("Error enabling events"));
					break;
				default :
					LogEventError(pszFunctionName, comerr, _T("Unknown error"));
					break;
			}
		}
		bResult = false;
	}
	catch(...)
	{
		if (bReportErrors)
		{
			LogEventError(pszFunctionName, _T("Unknown error"));
		}
		bResult = false;
	}

	return bResult;
}

void CThreadPoolManagerMSMQ1::CreatorThreadCloseDown()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadCloseDown\n"), (LPWSTR)m_bstrQueueName);
	CreatorThreadCloseDownThreadPool();
	CreatorThreadCloseDownMSMQ();
	m_pPrfInstanceQueueInfo = g_pPrfMonMessageQueueListener->RemoveInstanceQueueInfo(m_bstrQueueName);;
}


void CThreadPoolManagerMSMQ1::CreatorThreadCloseDownThreadPool()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadCloseDownThreadPool\n"), (LPWSTR)m_bstrQueueName);
	CThreadPoolManager::CloseDown();
}

void CThreadPoolManagerMSMQ1::CreatorThreadCloseDownMSMQ()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadCloseDownMSMQ\n"), (LPWSTR)m_bstrQueueName);
	if (m_MSMQEventPtr)
	{
		AtlUnadvise(m_MSMQEventPtr, __uuidof(_DMSMQEventEvents), m_dwCookie);
		m_MSMQEventPtr = NULL;
	}
	if (m_MSMQQueuePtr)
	{
		m_MSMQQueuePtr->Close();
		m_MSMQQueuePtr = NULL;
	}
}

void CThreadPoolManagerMSMQ1::CreatorThreadEnableEvents(long lCursor)
{
	const LPCTSTR pszFunctionName = _T("CreatorThreadEnableEvents");

	// only enable events if there are threads in the pool to process the events
	// and the queue is not stopping
	if (m_pThreadPoolThreadHead && m_bQueueStarted)
	{
		try
		{
			// lCursor (MQMSG_FIRST = 0, MQMSG_CURRENT = 1, MQMSG_NEXT = 2)
			const LPCTSTR pszCursor[] = {_T("MQMSG_FIRST"), _T("MQMSG_CURRENT"), _T("MQMSG_NEXT")};
			_ASSERTE(lCursor >= 0 && lCursor <= 2);
			if ((reinterpret_cast<bool>(::InterlockedCompareExchange((PVOID*)&m_bNotificationEnabled, (PVOID)TRUE, (PVOID)FALSE))) == FALSE)
			{
				ANYTHREADLOG_THIS_INOUT(_T("%s EnableNotification (Cursor = %s)\n"), (LPWSTR)m_bstrQueueName, pszCursor[lCursor]);
				HRESULT hr = m_MSMQQueuePtr->EnableNotification(m_MSMQEventPtr, &_variant_t(lCursor), &_variant_t(s_lEnableNotificationTimeOut));
				if (hr == MQ_INFORMATION_OPERATION_PENDING)
				{
					hr = S_OK;
				}
				_ASSERTE(SUCCEEDED(hr));
			}
		}
		catch (_com_error comerr) 
		{
			LogEventError(pszFunctionName, comerr, _T("Error enabling events"));
		}
		catch(...)
		{
			LogEventError(pszFunctionName, _T("Unknown error"));
		}
	}
}

// called by threads from the threadpool
CThreadPoolMessage* CThreadPoolManagerMSMQ1::RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread) // returns NULL if there is no more work to be done 
{
	const LPCTSTR pszFunctionName = _T("RemoveMessageFromQueue");

	// worker thread cannot retrieve another message so send it to sleep (otherwise marshalling of interface
	// pointers would be necessary)
	// ... so thread is available for reuse
	pThreadPoolThread->SetThreadAvailableForReuse();
	// ... post a notification message if required
	if ((reinterpret_cast<bool>(::InterlockedCompareExchange((PVOID*)&m_bThreadAvailablePosted, (PVOID)TRUE, (PVOID)FALSE))) == FALSE)
	{
		ANYTHREADLOG_THIS_INOUT(_T("%s PostThreadAvailableMessage\n"), (LPWSTR)m_bstrQueueName);
		PostThreadMessage(m_dwCreatorThreadId, WM_MSMQ1_THREADAVAILABLEFORREUSE, 0, 0);
	}
	return NULL;
}

// called by the thread which creates the MSMQ COM objects
void CThreadPoolManagerMSMQ1::CreatorThreadDispatchMessageToWorkerThread()
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadDispatchMessageToWorkerThread\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("CreatorThreadDispatchMessageToWorkerThread");

	try
	{
		CReadLock lckResize(GetrsharedlockResizeThreads());
		if (m_pThreadPoolThreadHead)
		{
			BOOL bTryAgain = FALSE;
			BOOL bAvailableThread = FALSE;
			CThreadPoolMessage* pThreadPoolMessage = NULL;
			do
			{
				bTryAgain = FALSE;

				// ... threads will either be suspended or about to suspend
				// ... try giving it directly to a thread
				bAvailableThread = FALSE;
				CThreadPoolThread* pThreadPoolThreadHead = m_pThreadPoolThreadHead; // make copy for local use
				CThreadPoolThread* pThreadPoolThread = m_pThreadPoolThreadHead;
				do
				{
					pThreadPoolThread = pThreadPoolThread->GetNextInSequence();
					bAvailableThread = pThreadPoolThread->TryToReuseThread();
				} while ((bAvailableThread == FALSE) && (pThreadPoolThread != pThreadPoolThreadHead));

				if (bAvailableThread)
				{
					// thread is available, tell it to do some work
					pThreadPoolMessage = CreatorThreadRemoveMessageFromQueue();
					if (pThreadPoolMessage)
					{
						pThreadPoolThread->SetRuntimeFunction(pThreadPoolMessage);
						pThreadPoolThread->ResumeThread();
        
						// change the head of the circular queue to balance thread loading
						m_pThreadPoolThreadHead = pThreadPoolThread;
						IncHitCtr(TRUE);

						// attempt to dispatch more messages to other threads
						bTryAgain = TRUE;
					}
					else
					{
						// no more messages to read within the timeout
						// ... so thread is available for reuse
						pThreadPoolThread->SetThreadAvailableForReuse();
						// ... enable events
						if (!m_bQueueStalled)
						{
							CreatorThreadEnableEvents(MQMSG_CURRENT);
						}

					}
				}
				else
				{
					// no thread is available
					// ... notifications are now required to indicate when threads are available
					m_bThreadAvailablePosted = FALSE;
					// but we need to retest for thread safety
					BOOL bAnyThreadBusy = FALSE;
					do
					{
						pThreadPoolThread = pThreadPoolThread->GetNextInSequence();
						if (pThreadPoolThread->IsThreadAvailableForReuse() == FALSE)
						{
							bAnyThreadBusy = TRUE;
							// the busy thread will post a notification message
						}
					} while ((bAnyThreadBusy == FALSE) && (pThreadPoolThread != pThreadPoolThreadHead));

					if (bAnyThreadBusy == FALSE)
					{
						if ((reinterpret_cast<bool>(::InterlockedCompareExchange((PVOID*)&m_bThreadAvailablePosted, (PVOID)TRUE, (PVOID)FALSE))) == FALSE)
						{
							// no threads busy, and notification was still enabled 
							bTryAgain = TRUE;
						}
					}
				}
					
			} while (bTryAgain);
		}
	}
	catch(_com_error comerr)
	{
		LogEventError(pszFunctionName, comerr, _T("Unknown error "));
	}
	catch(...)
	{
		LogEventError(pszFunctionName, _T("Unknown error"));
	}
}

CThreadPoolMessage* CThreadPoolManagerMSMQ1::CreatorThreadRemoveMessageFromQueue() // returns NULL if there is no more work to be done 
{
	ANYTHREADLOG_THIS_INOUT(_T("%s CreatorThreadRemoveMessageFrom\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("CreatorThreadRemoveMessageFrom");

	CThreadPoolMessage* pThreadPoolMessage = NULL;
	if (!m_bQueueStalled && m_bQueueStarted)
	{
		enum
		{
			eNULL,
			ePeekNext,
			eGetId,
			eGiveMessageToThread,
		} eAction = eNULL;

		bool bResetQueueIfNoMessages = true; // can reset the queue only once
		bool bTryAgain;
		bool bComponentStalled;
		do
		{
			bTryAgain = false;
			bComponentStalled = false;
			try
			{
				eAction = ePeekNext;
				IMSMQMessagePtr MSMQMessagePtr;
				if ((reinterpret_cast<bool>(::InterlockedCompareExchange((PVOID*)&m_ePeek, (PVOID)ePeekCurrentInProgress, (PVOID)ePeekCurrentToDo))) == ePeekCurrentToDo)
				{
					ANYTHREADLOG_THIS_INOUT(_T("%s PeekCurrent\n"), (LPWSTR)m_bstrQueueName);
					MSMQMessagePtr = m_MSMQQueuePtr->PeekCurrent(&_variant_t(false), &_variant_t(false), &_variant_t(s_lPeekTimeOut));
					if (MSMQMessagePtr)
					{
						m_ePeek = ePeekCurrentDone;
					}
					else
					{
						m_ePeek = ePeekCurrentToDo;
					}
				}
				else
				{
					while (m_ePeek != ePeekCurrentDone)
					{
						::Sleep(0);
					}
					ANYTHREADLOG_THIS_INOUT(_T("%s PeekNext\n"), (LPWSTR)m_bstrQueueName);
					MSMQMessagePtr = m_MSMQQueuePtr->PeekNext(&_variant_t(false), &_variant_t(false), &_variant_t(s_lPeekTimeOut));
				}
				if (MSMQMessagePtr)
				{
					_bstr_t bstrProgID;
					_bstr_t bstrConfig;
					_variant_t vIdMessage;
					{
						ANYTHREADLOG_THIS_INOUT(_T("%s Message Properties\n"), (LPWSTR)m_bstrQueueName);					
						eAction = eGetId;
						_bstr_t bstrProgID(MSMQMessagePtr->GetLabel());
						
						bComponentStalled = IsComponentStalled(bstrProgID);
						if (bComponentStalled)
						{
							::InterlockedIncrement(&m_lTotalMessagesSkippedBecauseStalled);
							ANYTHREADLOG_THIS(_T("%s Component stalled. Try again\n"), (LPWSTR)m_bstrQueueName);
							bTryAgain = true;
						}
						else
						{
							vIdMessage = MSMQMessagePtr->GetId();
							_variant_t vIdCorrelation(MSMQMessagePtr->GetCorrelationId());

							_bstr_t bstrIdMessageASCII = MSMQIdToBstrt(vIdMessage);
							_bstr_t bstrIdCorrelationASCII = MSMQIdToBstrt(vIdCorrelation);

							bstrConfig = 
								_T("<CONFIG>")
									_T("<QUEUE>")
										_T("<TYPE>MSMQ1</TYPE>")
										_T("<NAME>");
							bstrConfig += m_bstrQueueName;
							bstrConfig +=
										_T("</NAME>")
									_T("</QUEUE>")
									_T("<MESSAGE>")
										_T("<MESSAGEID>");
							bstrConfig += bstrIdMessageASCII;
							bstrConfig +=
										_T("</MESSAGEID>")
										_T("<CORRELATIONID>");
							bstrConfig += bstrIdCorrelationASCII;
							bstrConfig += 
										_T("</CORRELATIONID>")
									_T("</MESSAGE>")
								_T("</CONFIG>");
						}
						if (!bComponentStalled && m_bQueueStarted)
						{
							eAction = eGiveMessageToThread;
							long lMessagesToScan = 3 * (GetnRequestedNumberOfThreads() + m_lTotalMessagesSkippedBecauseStalled + m_lTotalMessagesRolledback);
							pThreadPoolMessage = new CThreadPoolMessage(
								reinterpret_cast<CThreadPoolMessage::funcptr>(ReceiveExecute), 
								(void*)(new CThreadData(*this, bstrConfig, m_bstrQueueName, m_bstrFormatName, m_bstrMoveQueueName, m_bstrMoveFormatName, vIdMessage, bstrProgID, lMessagesToScan, m_pPrfInstanceQueueInfo))
								);
							// allocated object is returned by this function
						}
					}
				}
				else
				{
					// timed out, or message not available
				}
			}
			catch(_com_error comerr)
			{
				switch (eAction)
				{
					case ePeekNext:
						switch (comerr.Error())
						{
							case MQ_ERROR_ILLEGAL_CURSOR_ACTION:
								// this error is expected when there are no more messages
								break;
							case MQ_ERROR_SERVICE_NOT_AVAILABLE: // fall-through
							case MQ_ERROR_INVALID_HANDLE:
							{
								// close and reopen the queue if the service is unavailable
								LogEventError(pszFunctionName, comerr, _T("Error peeking at next message.  Waiting for MSMQ service to become available"));
								bool bMSMQRestarted = false;
								while (m_bQueueStarted == TRUE && bMSMQRestarted == false)
								{
									try
									{
										CreatorThreadCloseDownMSMQ();
										const bool bReportErrors = false;
										bMSMQRestarted = CreatorThreadStartUpMSMQ(bReportErrors);
										if (!bMSMQRestarted)
										{
											::Sleep(s_dwPollMSMQService);
										}
									}
									catch(...)
									{
									}
								}
								if (bMSMQRestarted)
								{
									LogEventInfo(pszFunctionName, _T("MSMQ service is now available"));
									bTryAgain = true;
								}
								break;
							}
							default :
								LogEventError(pszFunctionName, comerr, _T("Error peeking at next message"));
								break;
						}
						break;
					case eGetId:
						LogEventError(pszFunctionName, comerr, _T("Error getting message id"));
						break;
					case eGiveMessageToThread:
						LogEventError(pszFunctionName, comerr, _T("Error dispatching message to a thread"));
						break;
					default :
						LogEventError(pszFunctionName, comerr, _T("Unknown error"));
						break;
				}
			}
			catch(...)
			{
				LogEventError(pszFunctionName, _T("Unknown error"));
			}

			if (!bComponentStalled && pThreadPoolMessage == NULL)
			{
				// ... reset queue if only one thread is active and there are no stalled components
				if (bResetQueueIfNoMessages && GetnThreadsActive() == 1 && !IsAnyComponentStalled())
				{
					ANYTHREADLOG_THIS_INOUT(_T("%s Reset\n"), (LPWSTR)m_bstrQueueName);
					bResetQueueIfNoMessages = false;
					m_MSMQQueuePtr->Reset();
					m_lTotalMessagesSkippedBecauseStalled = 0;
					m_ePeek = ePeekCurrentToDo;
					bTryAgain = true;
				}
			}
			if (!m_bQueueStarted)
			{
				// shutting down so don't try again
				bTryAgain = false;
			}

		} while (bTryAgain);
	}

	return pThreadPoolMessage;
}

void CThreadPoolManagerMSMQ1::ReceiveExecute(CThreadData* pThreadData)
{
	ANYTHREADLOG_THREADDATA_INOUT(_T("%s ReceiveExecute\n"), (LPWSTR)pThreadData->m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("ReceiveExecute");

	BSTR bstrErrorMessage = NULL;

	enum
	{
		eNULL,
		eReceiveExecute,
		eHandleFailures,
	} eAction = eNULL;
	
	try
	{
		// try executing the message nMaxTries Times
		eAction = eReceiveExecute;
		const int nMaxTries = 3;
		int nTry = 0;
		long lMESSQ_RESP = MESSQ_RESP_RETRY_NOW;
		while (nTry < nMaxTries && lMESSQ_RESP == MESSQ_RESP_RETRY_NOW)
		{
			IInternalMessageQueueListenerMTS1Ptr InternalMessageQueueListenerMTS1Ptr(__uuidof(MessageQueueListenerMTS1));
			if (InternalMessageQueueListenerMTS1Ptr)
			{
				lMESSQ_RESP = InternalMessageQueueListenerMTS1Ptr->MSMQReceiveExecute(::GetCurrentThreadId(), pThreadData->m_bstrConfig, pThreadData->m_bstrQueueName, pThreadData->m_bstrFormatName, pThreadData->m_bstrMoveQueueName, pThreadData->m_bstrMoveFormatName, pThreadData->m_vIdMessage, pThreadData->m_lMessagesToScan, &bstrErrorMessage);
				if (bstrErrorMessage != NULL)
				{
					USES_CONVERSION;
					pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, OLE2T(bstrErrorMessage));
				}
			}
			else
			{
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, _T("Error creating component or querying interface IMessageQueueListenerMTS1"));
				break;
			}
			nTry++;
		}

		// handle failures to execute		
		if (lMESSQ_RESP != MESSQ_RESP_SUCCESS)
		{
			::InterlockedIncrement(&pThreadData->m_rThreadPoolManagerMSMQ1.m_lTotalMessagesRolledback);
		}
		eAction = eHandleFailures;
		switch (lMESSQ_RESP)
		{
			case MESSQ_RESP_SUCCESS:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_SUCCESS\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementSuccess();
				break;
			}
			case MESSQ_RESP_RETRY_NOW:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_RETRY_NOW\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryNow();
				break;
			}
			case MESSQ_RESP_RETRY_LATER:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_RETRY_LATER\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryLater();
				break;
			}
			case MESSQ_RESP_RETRY_MOVE_MESSAGE:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_RETRY_MOVE_MESSAGE\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				IInternalMessageQueueListenerMTS1Ptr InternalMessageQueueListenerMTS1Ptr(__uuidof(MessageQueueListenerMTS1));
				if (InternalMessageQueueListenerMTS1Ptr)
				{
					lMESSQ_RESP = InternalMessageQueueListenerMTS1Ptr->MSMQMoveMessage(::GetCurrentThreadId(), pThreadData->m_bstrQueueName, pThreadData->m_bstrFormatName, pThreadData->m_bstrMoveQueueName, pThreadData->m_bstrMoveFormatName, pThreadData->m_vIdMessage, pThreadData->m_lMessagesToScan, &bstrErrorMessage);
					if (bstrErrorMessage != NULL)
					{
						USES_CONVERSION;
						pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, OLE2T(bstrErrorMessage));
					}
					else
					{
						pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryLater();
					}
				}
				else
				{
					pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, _T("Error creating component or querying interface IMessageQueueListenerMTS1"));
				}
				break;
			}
			case MESSQ_RESP_STALL_COMPONENT:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_STALL_COMPONENT\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				BOOL bAlreadyStalled;
				pThreadData->m_rThreadPoolManagerMSMQ1.AddStallComponent(pThreadData->m_bstrProgID, &bAlreadyStalled);
				if (!bAlreadyStalled)
				{
					pThreadData->m_pPrfInstanceQueueInfo->IncrementStalledComponents();
				}
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventWarning(pszFunctionName, _T("Stalling component %ls as requested by the component just called"), (LPCWSTR)pThreadData->m_bstrProgID);
				break;
			}
			case MESSQ_RESP_STALL_QUEUE:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_STALL_QUEUE\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_rThreadPoolManagerMSMQ1.m_bQueueStalled = TRUE;
				pThreadData->m_pPrfInstanceQueueInfo->SetbStallQueue(TRUE);
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventWarning(pszFunctionName, _T("Stalling queue as requested by the component just called"));
				break;
			}
			case MESSQ_RESP_HANDLED_RETRY_NOW:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_HANDLED_RETRY_NOW\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryNow();
				break;
			}
			case MESSQ_RESP_HANDLED_RETRY_LATER:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_HANDLED_RETRY_LATER\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryLater();
				break;
			}
			case MESSQ_RESP_HANDLED_RETRY_MOVE_MESSAGE:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_HANDLED_RETRY_MOVE_MESSAGE\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryMoveMessage();
				break;
			}
			case MESSQ_RESP_HANDLED_STOLEN:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_HANDLED_STOLEN\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_pPrfInstanceQueueInfo->IncrementStolen();
				break;
			}
			case MESSQ_RESP_HANDLED_ERROR:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_HANDLED_ERROR\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				pThreadData->m_pPrfInstanceQueueInfo->IncrementError();
				break;
			}
			default :
			{
				ANYTHREADLOG_THREADDATA(_T("%ls Unknown MESSQ_RESP\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				_ASSERTE(0); // should not reach here 
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, _T("Unknown MESSQ_RESP returned from component"));
				break;
			}
		}
	}
	catch(_com_error comerr)
	{
		::InterlockedIncrement(&pThreadData->m_rThreadPoolManagerMSMQ1.m_lTotalMessagesRolledback);
		switch (eAction)
		{
			case eReceiveExecute:
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, comerr, _T("Error receiving/executing message"));
				break;
			case eHandleFailures:
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, comerr, _T("Error handling failures"));
				break;
			default :
				pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, comerr, _T("Unknown error"));
				break;
		}
	}
	catch(...)
	{
		::InterlockedIncrement(&pThreadData->m_rThreadPoolManagerMSMQ1.m_lTotalMessagesRolledback);
		pThreadData->m_rThreadPoolManagerMSMQ1.LogEventError(pszFunctionName, _T("Unknown error"));
	}

	::SysFreeString(bstrErrorMessage);

	delete pThreadData;
}


void CThreadPoolManagerMSMQ1::LogEventInfo(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		_Module.LogEventInfo(_T("MSMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerMSMQ1::LogEventWarning(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		_Module.LogEventWarning(_T("MSMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerMSMQ1::LogEventError(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		_Module.LogEventError(_T("MSMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerMSMQ1::LogEventError(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		if (comerr.Description().length() == 0)
		{
			// output error message instead of description
			_Module.LogEventError(_T("MSMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', ErrorMessage = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.ErrorMessage());
		}
		else
		{
			// output description
			_Module.LogEventError(_T("MSMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', Source = '%s', Description = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
		}
	}
	catch(...)
	{
	}
}

_bstr_t CThreadPoolManagerMSMQ1::MSMQIdToBstrt(_variant_t vId)
{
	WCHAR wszId[50] = {'\0'};
	
	if (vId.vt == (VT_ARRAY | VT_UI1))
	{
		SAFEARRAY* psa = ((VARIANT)vId).parray;

		unsigned char* p20Bytes = NULL;
		HRESULT hr = ::SafeArrayAccessData(psa, (void HUGEP**)&p20Bytes);
		if (SUCCEEDED(hr))
		{
			const int nBytesMax = 20;
			if (psa->rgsabound->cElements == nBytesMax)
			{
				for (int nBytes = 0; nBytes < nBytesMax; nBytes++)
				{
					swprintf(wszId + nBytes * 2, L"%02x", p20Bytes[nBytes]);
				}
			}
			::SafeArrayUnaccessData(psa);
		}
	}

	return _bstr_t(wszId);	
}

