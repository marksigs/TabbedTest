///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerOMMQ1.cpp
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
//	LD		10/09/01	SYS2705 - Support for SQL Server added
//	LD		12/09/01    SYS2707 - Serialize access to public methods
//	LD		03/10/01	SYS2766 - Unlock message on MESSQ_RESP_HANDLED_ERROR.
//								  Reenable the poller on errors if queue is 
//								  not stalled
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		09/02/02	SYS4414	- Add Oracle's OLEDB provider
//								- bstrConfig expanded to include message information
//								- Additional logging.  
//								- ResetQueue called on StartUp
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Create queues on a new thread
//								- Keep alive an ADO connection to enable pooling
//								- Add originating thread id
//								- Stop queue amendments
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//								- Compensate for errors due to Oracle not fully
//								  initialised on a reboot.
//								- Introduce m_bDBConnectivityGood to reduce number of
//								  errors to event log when database connectivity is bad
//  LD		06/02/03	OPC0006008 - Explicity close connections when an excpetion
//						occurs
//	LD		13/02/03	Do not retry or reopen keep alive connection if queue is shutting down
//  RF		24/02/04	BMIDS727 - Allow for multiple queue tables
//	RF		02/04/04	BMIDS727 - Allow for multiple queue tables - improved error handling
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "mutex.h"
using namespace namespaceMutex;
#include "MessageQueueListener.h"
#include "ThreadPoolManagerOMMQ1.h"
#include "ThreadPoolThread.h"
#include "PrfMonMessageQueueListener.h"
#include "DbHelperOMMQ.h"
#include <OLEDBERR.H> // for DB_E_ERRORSINCOMMAND

#import "..\MessageQueueListenerMTS\MessageQueueListenerMTS.tlb" no_namespace
#import "..\MessageQueueComponentVC\MessageQueueComponentVC.tlb" no_namespace

_bstr_t CThreadPoolManagerOMMQ1::s_bstrNULL(static_cast<LPWSTR>(NULL));
_variant_t CThreadPoolManagerOMMQ1::s_vZero(0L);
_variant_t CThreadPoolManagerOMMQ1::s_vOne(1L);
UINT CThreadPoolManagerOMMQ1::WM_OMMQ1_OPENKEEPALIVE = RegisterWindowMessage(_T("WM_OMMQ1_OPENKEEPALIVE"));

const DWORD s_dwPollConnection = 5000; // in ms when waiting for database to become available

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_OMMQ1); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_OMMQ1); LogInOut.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA CLogIn LogIn(pThreadData->m_rThreadPoolManagerOMMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_OMMQ1); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA_INOUT CLogInOut LogInOut(pThreadData->m_rThreadPoolManagerOMMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_OMMQ1); LogInOut.AnyThreadInitialize

///////////////////////////////////////////////////////////////////////////////

bool CThreadPoolManagerOMMQ1::CPoller::StartUp()
{
	ANYTHREADLOG_THIS_INOUT(_T("CPoller::StartUp\n"));
	return CScheduleManager::StartUp();
}

void CThreadPoolManagerOMMQ1::CPoller::CloseDown()
{
	ANYTHREADLOG_THIS_INOUT(_T("CPoller::CloseDown\n"));
	CScheduleManager::CloseDown();
}

void CThreadPoolManagerOMMQ1::CPoller::OnEventSchedule()
{
	m_ThreadPoolManagerOMMQ1.RemoveMessageFromQueue(
		/*CThreadPoolThread* pThreadPoolThread =*/NULL, 
		/*bDispatchToAvailableThreads =*/true);
}

///////////////////////////////////////////////////////////////////////////////

CThreadPoolManagerOMMQ1::CThreadPoolManagerOMMQ1() :
	m_Poller(*this),
	m_bQueueStalled(FALSE),
	m_bQueueStarted(FALSE),
	m_pPrfInstanceQueueInfo(NULL),
	m_lProvider(PROVIDER_UNKNOWN),
	m_bstrQueueName(L"Notdefined"),
	m_dwThreadId(0),
	m_hThread(NULL),
	m_hEventStartUp(NULL),
	m_hEventCloseDown(NULL),
	m_bDBConnectivityGood(false),
	m_bProcessingOpenKeepAlive(false)
{
	// set up the locked by name for this service which must be unique.
	// ... using computer name
	TCHAR szLockedBy[MAX_COMPUTERNAME_LENGTH + 1]; // using computername
	DWORD dwSizeComputerName = MAX_COMPUTERNAME_LENGTH;
	::GetComputerName(szLockedBy, &dwSizeComputerName);

	m_bstrLockedBy = szLockedBy;
}

CThreadPoolManagerOMMQ1::~CThreadPoolManagerOMMQ1()
{

}

HRESULT CThreadPoolManagerOMMQ1::FinalConstruct()
{
	return CComObjectRootEx<CComMultiThreadModel>::FinalConstruct();
}

void CThreadPoolManagerOMMQ1::FinalRelease()
{
	CComObjectRootEx<CComMultiThreadModel>::FinalRelease();
}

///////////////////////////////////////////////////////////////////////////////
// IThreadPoolManagerCommon1

STDMETHODIMP CThreadPoolManagerOMMQ1::put_Started(BOOL newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_Started\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_Started");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	HRESULT hRes = S_FALSE;

	m_bQueueStarted = newVal;
	if(m_bQueueStarted == TRUE)
	{
		LogEventInfo(pszFunctionName, 
			_T("Starting to listen to queue '%s' using table suffix '%s' with %d threads"), 
			(LPWSTR)m_bstrQueueName, (LPWSTR)m_bstrTableSuffix, GetnRequestedNumberOfThreads());

		// create msmq objects on a new thread
		// ... Create an event, which indicates completion of RasDial()
		_ASSERT(m_hEventStartUp == NULL);
		m_hEventStartUp = CreateEvent(NULL, FALSE, FALSE, NULL);
		if (m_hEventStartUp)
		{
			// ... create a new thread on which to create MSMQ objects
			_ASSERT(m_hThread == NULL);
			m_dwThreadId = NULL;
			m_hThread = ::CreateThread(NULL, 0, StartUponThread, this, 0, &m_dwThreadId);

			// ... and wait for them to be created
			DWORD dwWait = WaitForSingleObject(m_hEventStartUp, INFINITE);
			switch (dwWait)
			{
			case WAIT_OBJECT_0: // objects created or an error has occurred
				///??
				break;
			case WAIT_TIMEOUT: // timed out
				break;
			case WAIT_ABANDONED:
				break;
			}
			CloseHandle(m_hEventStartUp);
			m_hEventStartUp = NULL;

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
		LogEventInfo(pszFunctionName, 
			_T("Stopping to listen to queue '%s'"), (LPWSTR)m_bstrQueueName);

		_ASSERT(m_hEventCloseDown == NULL);
		m_hEventCloseDown = CreateEvent(NULL, FALSE, FALSE, NULL);
		if (m_hEventCloseDown)
		{
			_ASSERT(m_hThread != NULL);
			// let the thread on which the msmq objects were created quit
			PostThreadMessage(m_dwThreadId, WM_QUIT, 0, 0);

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

			CloseHandle(m_hThread);
			m_hThread = NULL;
			m_dwThreadId = 0; 

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


STDMETHODIMP CThreadPoolManagerOMMQ1::get_Started(BOOL *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_Started\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bQueueStarted;
	return S_OK;
}


STDMETHODIMP CThreadPoolManagerOMMQ1::get_NumberOfThreads(long *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_NumberOfThreads\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = GetnRequestedNumberOfThreads();
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_NumberOfThreads(long newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_NumberOfThreads (newVal = %ld)\n"), 
		(LPWSTR)m_bstrQueueName, newVal);
	const LPCTSTR pszFunctionName = _T("put_NumberOfThreads");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	if (newVal != GetnRequestedNumberOfThreads())
	{
		if (m_bQueueStarted)
		{
			LogEventInfo(pszFunctionName, 
				_T("Changing the number of threads to %d on queue '%s'"), 
				newVal, (LPWSTR)m_bstrQueueName);
		}
		SetnRequestedNumberOfThreads(newVal);
	}
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::get_QueueName(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueName\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bstrQueueName.copy();
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_QueueName(BSTR newVal)
{
	//ANYTHREADLOG_THIS_INOUT(_T("%s put_QueueName\n"), (LPWSTR)m_bstrQueueName);
	ANYTHREADLOG_THIS_INOUT(_T("%s put_QueueName (newVal = %s)\n"), 
		(LPWSTR)m_bstrQueueName,(LPWSTR)newVal);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bstrQueueName = newVal;
	return S_OK;
}

// RF 24/02/04
STDMETHODIMP CThreadPoolManagerOMMQ1::get_TableSuffix(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_TableSuffix\n"), (LPWSTR)m_bstrTableSuffix);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bstrTableSuffix.copy();
	return S_OK;
}

// RF 24/02/04
STDMETHODIMP CThreadPoolManagerOMMQ1::put_TableSuffix(BSTR newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_TableSuffix (newVal = %s)\n"), 
		(LPWSTR)m_bstrQueueName,(LPWSTR)newVal);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bstrTableSuffix = newVal;
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::get_QueueStalled(BOOL *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueStalled\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = m_bQueueStalled;
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_QueueStalled(BOOL newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_QueueStalled\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_QueueStalled");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	m_bQueueStalled = newVal;
	if (m_bQueueStalled)
	{
		LogEventWarning(pszFunctionName, _T("Stalling queue as requested by an external component"));
	}
	else
	{
		LogEventInfo(pszFunctionName, _T("Unstalling queue as requested by an external component"));
	}

	if (!m_bQueueStalled)
	{
		{
			CWriteLock lckWriteQueue(m_sharedlockResetQueue);
			ResetQueue();
		}
		m_pPrfInstanceQueueInfo->SetbStallQueue(newVal);
		// reenable single shot events
		m_Poller.ExternalEnableEvents();
	}
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::get_ComponentsStalled(VARIANT *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_ComponentsStalled\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	return CThreadPoolManagerMQ::get_ComponentsStalled(pVal);
}


STDMETHODIMP CThreadPoolManagerOMMQ1::put_AddStalledComponents(VARIANT newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_AddStalledComponents\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_AddStalledComponents");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	LogEventWarning(pszFunctionName, _T("Stalling components as requested due to an external request"));
	
	HRESULT hr = CThreadPoolManagerMQ::put_AddStalledComponents(newVal);
	m_pPrfInstanceQueueInfo->SetnDifferentStalledComponents(GetnDifferentComponentsStalled());
	return hr;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_RestartComponents(VARIANT newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_RestartComponents\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_RestartComponents");
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	LogEventInfo(pszFunctionName, _T("Unstalling components as requested due to an external request"));

	HRESULT hr = CThreadPoolManagerMQ::put_RestartComponents(newVal);
	m_pPrfInstanceQueueInfo->SetnDifferentStalledComponents(GetnDifferentComponentsStalled());
	{
		CWriteLock lckWriteQueue(m_sharedlockResetQueue);
		ResetQueue();
	}
	// reenable single shot events
	m_Poller.ExternalEnableEvents();
	return hr;
}


STDMETHODIMP CThreadPoolManagerOMMQ1::get_QueueType(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_QueueType\n"), (LPWSTR)m_bstrQueueName);
	CSingleLock lck(&m_csPublicMethod, TRUE);
	
	*pVal = ::SysAllocString(L"OMMQ1");
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
// IThreadPoolManagerOMMQ1

STDMETHODIMP CThreadPoolManagerOMMQ1::get_ConnectionString(BSTR *pVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s get_ConnectionString\n"), (LPWSTR)m_bstrQueueName);
	*pVal = m_bstrConnectionString.copy();
	return S_OK;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_ConnectionString(BSTR newVal)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_ConnectionString\n"), (LPWSTR)m_bstrQueueName);
	const LPCTSTR pszFunctionName = _T("put_ConnectionString");

	HRESULT hr = E_FAIL;;
	m_bstrConnectionString = newVal;

	_bstr_t bstrConnectionString = _tcsupr(m_bstrConnectionString);
	if (_tcsstr((LPCWSTR)bstrConnectionString, L"MSDAORA") != NULL)
	{
		m_lProvider = PROVIDER_MSDAORA;
		hr = S_OK;
	}
	else if (_tcsstr((LPCWSTR)bstrConnectionString, L"SQLOLEDB") != NULL)
	{
		m_lProvider = PROVIDER_SQLOLEDB;
		hr = S_OK;
	}
	else if (_tcsstr((LPCWSTR)bstrConnectionString, L"ORAOLEDB") != NULL)
	{
		m_lProvider = PROVIDER_ORAOLEDB;
		hr = S_OK;
	}
	else
	{
		LogEventError(pszFunctionName, _T("Unknown provider specified in database connection string"));
		m_lProvider = PROVIDER_UNKNOWN;
	}
	return hr;
}

STDMETHODIMP CThreadPoolManagerOMMQ1::put_dwNextWaitIntervalms(DWORD dwPollInterval)
{
	ANYTHREADLOG_THIS_INOUT(_T("%s put_dwNextWaitIntervalms\n"), (LPWSTR)m_bstrQueueName);
	m_Poller.put_dwNextWaitIntervalms(dwPollInterval);
	return S_OK;
}

/////////////////////////////////////////////////////////////////////////////

DWORD WINAPI CThreadPoolManagerOMMQ1::StartUponThread(void* pThreadPoolManagerOMMQ1)
{
	return reinterpret_cast<CThreadPoolManagerOMMQ1*>(pThreadPoolManagerOMMQ1)->StartUponThread();
}

DWORD WINAPI CThreadPoolManagerOMMQ1::StartUponThread()
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
			bool bUnlockMessagesCalled = false;

			// some ADO connections are not running under MTS/COMPLUS
			// there a connection must be kept alive to enable connection
			// pooling
			// Compensate for errors due to database not fully initialised on a reboot
			_ConnectionPtr KeepAliveConnectionPtr;
			try
			{
				KeepAliveConnectionPtr.CreateInstance(__uuidof(Connection));
				KeepAliveConnectionPtr->ConnectionString = m_bstrConnectionString;
				KeepAliveConnectionPtr->CursorLocation = adUseClient;
				KeepAliveConnectionPtr->Open("", "", "", 0);
				// clear down on start.  Using the connection also tests 
				// that it is functioning
				if (UnlockMessages())
				{
					bUnlockMessagesCalled = true;
					m_bDBConnectivityGood = true;
					CWriteLock lckWriteQueue(m_sharedlockResetQueue);
					ResetQueue();
				}
				else
				{
					CloseConnection(KeepAliveConnectionPtr);
				}
			}
			catch(_com_error comerr)
			{
				LogEventWarning(pszFunctionName, comerr, _T("Error opening keep-alive connection.  Attempting to establish database connectivity"));
				CloseConnection(KeepAliveConnectionPtr);
			}
			catch(...)
			{
				LogEventWarning(pszFunctionName, _T("Error opening keep-alive connection.  Attempting to establish database connectivity"));
				CloseConnection(KeepAliveConnectionPtr);
			}

			eAction = eStartUp;			
			m_bQueueStarted = StartUp();
			SetEvent(m_hEventStartUp); // signal startup complete or error through m_bQueueStarted

			eAction = eMessageLoop;			
			MSG msg;
			while (GetMessage(&msg, 0, 0, 0))
			{
				if (msg.message == WM_OMMQ1_OPENKEEPALIVE)
				{
					m_bProcessingOpenKeepAlive = true;
					if (!m_bDBConnectivityGood)
					{
						CloseConnection(KeepAliveConnectionPtr);
						try
						{
							KeepAliveConnectionPtr.CreateInstance(__uuidof(Connection));
							KeepAliveConnectionPtr->ConnectionString = m_bstrConnectionString;
							KeepAliveConnectionPtr->CursorLocation = adUseClient;
							KeepAliveConnectionPtr->Open("", "", "", 0);
							// clear down on start.  Using the connection also tests 
							// that it is functioning
							if (!bUnlockMessagesCalled)
							{
								if (UnlockMessages())
								{
									LogEventInfo(pszFunctionName, _T("Keep-alive connection opended.  Database connectivity established"));
									bUnlockMessagesCalled = true;
									m_bDBConnectivityGood = true;
									CWriteLock lckWriteQueue(m_sharedlockResetQueue);
									ResetQueue();
								}
								else
								{
									CloseConnection(KeepAliveConnectionPtr);
								}
							}
							else
							{
								CWriteLock lckWriteQueue(m_sharedlockResetQueue);
								if (ResetQueue())
								{
									LogEventInfo(pszFunctionName, _T("Keep-alive connection opended.  Database connectivity established"));
									m_bDBConnectivityGood = true;
								}
							}
						}
						catch(_com_error comerr)
						{
							CloseConnection(KeepAliveConnectionPtr);
						}
						catch(...)
						{
							CloseConnection(KeepAliveConnectionPtr);
						}
					}
					m_bProcessingOpenKeepAlive = false;
				}
				else
				{
					DispatchMessage(&msg);
				}
			}

			eAction = eCloseDown;			
			CloseDown();
			CloseConnection(KeepAliveConnectionPtr);
			m_bQueueStarted = false;
			SetEvent(m_hEventCloseDown); // signal closedown complete or error through m_bQueueStarted
		}
		catch(_com_error comerr)
		{
			switch (eAction)
			{
				case eStartUp:
					LogEventError(pszFunctionName, comerr, _T("Error during startup"));
					SetEvent(m_hEventStartUp); // signal startup complete or error through m_bQueueStarted
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
					SetEvent(m_hEventStartUp); // signal startup complete or error through m_bQueueStarted
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

bool CThreadPoolManagerOMMQ1::StartUp()
{
	const LPCTSTR pszFunctionName = _T("StartUp");

	bool bResult = true;

	enum
	{
		eNULL,
		eStartThreadPool,
		eStartPoller
	} eAction = eNULL;
	
	try
	{
		// start the thread pool and poller
		eAction = eStartThreadPool;
		bResult = CThreadPoolManager::StartUp();
		if (bResult)
		{
			eAction = eStartPoller;
			bResult = m_Poller.StartUp();
			m_pPrfInstanceQueueInfo = g_pPrfMonMessageQueueListener->AddInstanceQueueInfo(m_bstrQueueName);
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eStartThreadPool:
				LogEventError(pszFunctionName, comerr, _T("Error starting thread pool"));
				break;
			case eStartPoller:
				LogEventError(pszFunctionName, comerr, _T("Error starting poller"));
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
	return bResult;
}

void CThreadPoolManagerOMMQ1::CloseDown()
{
	m_Poller.CloseDown();
	CThreadPoolManager::CloseDown();
	m_pPrfInstanceQueueInfo = g_pPrfMonMessageQueueListener->RemoveInstanceQueueInfo(m_bstrQueueName);
}

///////////////////////////////////////////////////////////////////////////////

CThreadPoolMessage* CThreadPoolManagerOMMQ1::RemoveMessageFromQueue(
	CThreadPoolThread* pThreadPoolThread, 
	bool bDispatchToAvailableThreads)
{
	const LPCTSTR pszFunctionName = _T("RemoveMessageFromQueue");

	CThreadPoolMessage* pThreadPoolMessage = NULL;
	CScheduleManager::CEvents::TeOnDestruct eOnDestruct;
	CScheduleManager::CEvents PollerEvents(m_Poller, CScheduleManager::CEvents::eOnDestructNULL);

	if (!m_bQueueStalled && m_bQueueStarted)
	{
		// automatically reenable single shot events
		if (bDispatchToAvailableThreads)
		{
			eOnDestruct = CScheduleManager::CEvents::eOnDestructInternalEnable;
		}
		else
		{
			eOnDestruct = CScheduleManager::CEvents::eOnDestructExternalEnable;
		}
		PollerEvents.SeteOnDestruct(eOnDestruct);

		if (!m_bDBConnectivityGood)
		{
			if (!m_bProcessingOpenKeepAlive && m_bQueueStarted)
			{
				PostThreadMessage(m_dwThreadId, WM_OMMQ1_OPENKEEPALIVE, 0, 0);
				Sleep(s_dwPollConnection);
			}
		}
		else
		{
			// prevent the queue from being reset by acquiring read lock only
			CReadLock lckReadQueue(m_sharedlockResetQueue);
			CWriteLock lckWriteQueue(m_sharedlockResetQueue, false);

			bool bResetQueueIfNoMessages = !bDispatchToAvailableThreads; // can reset the queue only once if not called by poller
			bool bTryAgain;
			bool bComponentStalled;
			do
			{
				bTryAgain = false;
				bComponentStalled = false;

				// returning one message or 
				// dispatching to threads ...work out the number of threads not active, and then try and allocate these number of messages to this service
				int nMessagesToAllocate = bDispatchToAvailableThreads ? GetnThreadsNotActive() : 1;
				if (nMessagesToAllocate > 0)
				{
					_ConnectionPtr ConnectionPtr;
					_RecordsetPtr RecordsetPtr;
					_CommandPtr CommandPtr;
					enum
					{
						eNULL,
						eOpenConnection,
						eSetupCommand,
						eSetupParameters,
						eExecuteCommand,
						eRetrieveData,
					} eAction = eNULL;
					
					try
					{
						// get the message (ProgID and XML) defined by its message id
						// ... open the connection
						eAction = eOpenConnection;
						ConnectionPtr.CreateInstance(__uuidof(Connection));
						ConnectionPtr->ConnectionString = m_bstrConnectionString;
						ConnectionPtr->CursorLocation = adUseClient;
						ConnectionPtr->Open("", "", "", 0);

						// ... set up the ADO command
						eAction = eSetupCommand;
						CommandPtr.CreateInstance(__uuidof(Command));
						CommandPtr->ActiveConnection = ConnectionPtr;

						switch (m_lProvider)
						{
							case PROVIDER_MSDAORA:
								CommandPtr->CommandType = adCmdText;
								CommandPtr->CommandText = "{call sp_MQLOMMQ.LockMessage(?,?,?)}";
								break;
							case PROVIDER_SQLOLEDB:
								CommandPtr->CommandType = adCmdStoredProc;
								// RF BMIDS727 Start
								//CommandPtr->CommandText = "USP_MQLOMMQLOCKMESSAGE";
								CDbHelperOMMQ* ptrHelper;
								ptrHelper = new CDbHelperOMMQ(); 
								if (ptrHelper == NULL)
								{
									_ASSERTE(0); // should not reach here
									throw L"Unable to set command text";
								}
								CommandPtr->CommandText = 
									ptrHelper->GetSQLOLEDBCommandText(
									m_bstrTableSuffix,
									CDbHelperOMMQ::eLockMessage);
								ptrHelper = NULL;
								// RF BMIDS727 End								
								break;
							case PROVIDER_ORAOLEDB:
								CommandPtr->CommandType = adCmdText;
								CommandPtr->CommandText = "{call sp_MQLOMMQ.LockMessage(?,?,?)}";
								break;
							default :
								_ASSERTE(0); // should not reach here 
								throw L"Unrecognised provider";
								break;
						}
						
						// ... set up the parameters for the commmand
						eAction = eSetupParameters;
						_ParameterPtr ParameterPtr;
						
						ParameterPtr = CommandPtr->CreateParameter(
							s_bstrNULL, adBSTR, adParamInput, SysStringLen(m_bstrQueueName), m_bstrQueueName);
						CommandPtr->GetParameters()->Append(ParameterPtr);

						ParameterPtr = CommandPtr->CreateParameter(
							s_bstrNULL, adBSTR, adParamInput, SysStringLen(m_bstrLockedBy), m_bstrLockedBy);
						CommandPtr->GetParameters()->Append(ParameterPtr);

						_variant_t vMessagesToAllocate((long)nMessagesToAllocate);
						ParameterPtr = CommandPtr->CreateParameter(
							s_bstrNULL, adInteger, adParamInput, 0, vMessagesToAllocate);
						CommandPtr->GetParameters()->Append(ParameterPtr);

						// ... execute command
						eAction = eExecuteCommand;
						RecordsetPtr = CommandPtr->Execute(NULL, NULL, 0);
						int nMessagesToProcess = RecordsetPtr->GetRecordCount();
						if (nMessagesToProcess >= nMessagesToAllocate)
						{
							// ... no need to reenable the single shot poller on destruction (threads are busy)
							PollerEvents.SeteOnDestruct(CScheduleManager::CEvents::eOnDestructNULL);
						}
						else
						{
							// ... reenable the events on destruction
							PollerEvents.SeteOnDestruct(eOnDestruct);
						}

						// ... retrive the ProgID and message id
						eAction = eRetrieveData;
						_variant_t vMessageId;
						while (nMessagesToProcess > 0)
						{
							vMessageId = RecordsetPtr->GetFields()->GetItem(s_vZero)->GetValue();
							_bstr_t bstrMessageId = MessageIdToBstrt(vMessageId);
							ANYTHREADLOG_THIS(_T("MessageId = %s\n"), (LPCTSTR)bstrMessageId);
							_bstr_t bstrProgID = RecordsetPtr->GetFields()->GetItem(s_vOne)->GetValue();

							bComponentStalled = IsComponentStalled(bstrProgID);
							if (bComponentStalled)
							{
								UnlockMessage(vMessageId);
								bTryAgain = true;
							}
							else
							{
								_bstr_t bstrConfig = 
									_T("<CONFIG>")
										_T("<QUEUE>")
											_T("<TYPE>OMMQ1</TYPE>")
											_T("<NAME>");
								bstrConfig += m_bstrQueueName;
								bstrConfig +=
											_T("</NAME>")
										_T("</QUEUE>")
										_T("<MESSAGE>")
											_T("<MESSAGEID>");
								bstrConfig += bstrMessageId;
								bstrConfig +=
											_T("</MESSAGEID>")
										_T("</MESSAGE>")
									_T("</CONFIG>");

								// ownership of message is either given to a thread (either dispatching to a thread here or returning as a result of this function)
								pThreadPoolMessage = new CThreadPoolMessage(
									reinterpret_cast<CThreadPoolMessage::funcptr>(ReceiveExecute), 
									(void*)(new CThreadData(
										*this, 
										bstrConfig, 
										m_lProvider, 
										m_bstrConnectionString, 
										m_bstrQueueName, 
										m_bstrTableSuffix,	// RF 30/03/04
										vMessageId, 
										bstrProgID, 
										m_pPrfInstanceQueueInfo))
									);
								
								if (bDispatchToAvailableThreads)
								{
									bool bMessageGivenToThread = false; 
									do 
									{
										CThreadPoolThread* pThreadPoolThread = TryGetpAvailableThreadPoolThread();
										if (pThreadPoolThread)
										{
											// thread is available, tell it to do some work
											pThreadPoolThread->SetRuntimeFunction(pThreadPoolMessage); // thread gains ownership of the message
											bMessageGivenToThread = pThreadPoolThread->ResumeThread();
										}
										else
										{
											::Sleep(0);	
										}
									} while (!bMessageGivenToThread);
									pThreadPoolMessage = NULL;
								}
							}
							RecordsetPtr->MoveNext();
							nMessagesToProcess--;
						}
						
						CloseCommand(CommandPtr);
						CloseRecordset(RecordsetPtr);
						CloseConnection(ConnectionPtr);
					}
					catch(_com_error comerr)
					{
						if (m_bDBConnectivityGood)
						{
							LogEventErrorProvider(ConnectionPtr);
							switch (eAction)
							{
								case eOpenConnection:
									LogEventError(pszFunctionName, comerr, 
										_T("Error opening connection. Attempting to re-establish database connectivity"));
									OnBadConnection();
									break;
								case eSetupCommand:
									LogEventError(pszFunctionName, comerr, 
										_T("Error setting up command"));
									break;
								case eSetupParameters:
									LogEventError(pszFunctionName, comerr, 
										_T("Error setting up parameters"));
									break;
								case eExecuteCommand:
									LogEventError(pszFunctionName, comerr, 
										_T("Error executing command"));
									break;
								case eRetrieveData:
									LogEventError(pszFunctionName, comerr, 
										_T("Error retrieving data"));
									break;
								default :
									LogEventError(pszFunctionName, comerr, 
										_T("Unknown error"));
									break;
							}
						}
						CloseCommand(CommandPtr);
						CloseRecordset(RecordsetPtr);
						CloseConnection(ConnectionPtr);
					}
					catch(...)
					{
						if (m_bDBConnectivityGood)
						{
							if (eAction == eOpenConnection)
							{
								LogEventErrorProvider(ConnectionPtr);
								LogEventError(pszFunctionName, _T("Unknown error opening connection. Attempting to re-establish database connectivity"));
								OnBadConnection();
							}
							else
							{
								LogEventErrorProvider(ConnectionPtr);
								LogEventError(pszFunctionName, _T("Unknown error"));
							}
						}
						CloseCommand(CommandPtr);
						CloseRecordset(RecordsetPtr);
						CloseConnection(ConnectionPtr);
					}
				}

				if (!bComponentStalled && pThreadPoolMessage == NULL)
				{
					// ... reset queue if only one thread is active and there are no stalled components
					if (bResetQueueIfNoMessages && GetnThreadsActive() == 1 && !IsAnyComponentStalled() && m_bQueueStarted)
					{
						// ... promote read lock to temporary possible write lock in case new messages are placed on queue. Events would be enabled so
						// ... another thread could re-enter.  The write lock would then not acquired
						lckReadQueue.UnLock();
						if (lckWriteQueue.TryLock())
						{
							bResetQueueIfNoMessages = false;
							ResetQueue();
							lckWriteQueue.UnLock();
						}
						else
						{
							// ... write lock not acquired, therefore more messages must have been received, so try getting one of them	
						}
						lckReadQueue.Lock();
						bTryAgain = true;
					}
				}
			} while (bTryAgain && m_bQueueStarted);
		}
	}

	try
	{
		if (pThreadPoolMessage == NULL)
		{
			if (bDispatchToAvailableThreads == false)
			{
				// ... so thread is available for reuse
				pThreadPoolThread->SetThreadAvailableForReuse();
			}
			if (!m_bQueueStalled && m_bQueueStarted)
			{
				// ... reenable the events on destruction 
				PollerEvents.SeteOnDestruct(eOnDestruct);
			}
		}
	}
	catch(_com_error comerr)
	{
		LogEventError(pszFunctionName, comerr, _T("Error enabling events"));
	}
	catch(...)
	{
		LogEventError(pszFunctionName, _T("Unknown error"));
	}
	
	return pThreadPoolMessage;
}

CThreadPoolMessage* CThreadPoolManagerOMMQ1::RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread) // returns NULL if there is no more work to be done 
{
	// get a message for this thread to process
	CThreadPoolMessage* pThreadPoolMessage = RemoveMessageFromQueue(pThreadPoolThread, /*bDispatchToAvailableThreads =*/false);
	return pThreadPoolMessage;
}

void CThreadPoolManagerOMMQ1::ReceiveExecute(CThreadData* pThreadData)
{
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
		while (nTry < nMaxTries && lMESSQ_RESP == MESSQ_RESP_RETRY_NOW && pThreadData->m_rThreadPoolManagerOMMQ1.m_bDBConnectivityGood)
		{
			IInternalMessageQueueListenerMTS1Ptr InternalMessageQueueListenerMTS1Ptr(__uuidof(MessageQueueListenerMTS1));
			if (InternalMessageQueueListenerMTS1Ptr)
			{
				lMESSQ_RESP = InternalMessageQueueListenerMTS1Ptr->OMMQReceiveExecute(
					::GetCurrentThreadId(), 
					pThreadData->m_bstrConfig, 
					pThreadData->m_lProvider, 
					pThreadData->m_bstrConnectionString, 
					pThreadData->m_bstrQueueName, 
					pThreadData->m_vMessageId, 
					pThreadData->m_bstrTableSuffix,	// RF 24/02/04
					&bstrErrorMessage);
				if (bstrErrorMessage != NULL)
				{
					USES_CONVERSION;
					pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, OLE2T(bstrErrorMessage));
				}
			}
			else
			{
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, _T("Error creating component or querying interface IMessageQueueListenerMTS1"));
				break;
			}
			nTry++;
		}

		// handle failures to execute		
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
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryNow();
				break;
			}
			case MESSQ_RESP_RETRY_LATER:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_RETRY_LATER\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryLater();
				break;
			}
			case MESSQ_RESP_RETRY_MOVE_MESSAGE:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_RETRY_MOVE_MESSAGE\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				IInternalMessageQueueListenerMTS1Ptr InternalMessageQueueListenerMTS1Ptr(__uuidof(MessageQueueListenerMTS1));
				if (InternalMessageQueueListenerMTS1Ptr)
				{
					InternalMessageQueueListenerMTS1Ptr->OMMQMoveMessage(
						::GetCurrentThreadId(), 
						pThreadData->m_lProvider, 
						pThreadData->m_bstrConnectionString, 
						pThreadData->m_bstrQueueName, 
						pThreadData->m_bstrTableSuffix,
						pThreadData->m_vMessageId, 
						&bstrErrorMessage);
					if (bstrErrorMessage != NULL)
					{
						USES_CONVERSION;
						pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, OLE2T(bstrErrorMessage));
					}
					else
					{
						pThreadData->m_pPrfInstanceQueueInfo->IncrementRetryMoveMessage();
					}
				}
				else
				{
					pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, _T("Error creating component or querying interface IMessageQueueListenerMTS1"));
				}
				break;
			}
			case MESSQ_RESP_STALL_COMPONENT:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_STALL_COMPONENT\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				BOOL bAlreadyStalled;
				pThreadData->m_rThreadPoolManagerOMMQ1.AddStallComponent(pThreadData->m_bstrProgID, &bAlreadyStalled);
				if (!bAlreadyStalled)
				{
					pThreadData->m_pPrfInstanceQueueInfo->IncrementStalledComponents();
				}
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventWarning(pszFunctionName, _T("Stalling component %ls as requested by the component just called"), (LPCWSTR)pThreadData->m_bstrProgID);
				break;
			}
			case MESSQ_RESP_STALL_QUEUE:
			{
				ANYTHREADLOG_THREADDATA(_T("%ls MESSQ_RESP_STALL_QUEUE\n"), (LPCWSTR)pThreadData->m_bstrQueueName);
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				pThreadData->m_rThreadPoolManagerOMMQ1.m_bQueueStalled = TRUE;
				pThreadData->m_pPrfInstanceQueueInfo->SetbStallQueue(TRUE);
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventWarning(pszFunctionName, _T("Stalling queue as requested by the component just called"));
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
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				pThreadData->m_pPrfInstanceQueueInfo->IncrementError();
				break;
			}
			default :
			{
				ANYTHREADLOG_THREADDATA(_T("%ls Unknown MESSQ_RESP\n"), (LPCWSTR)pThreadData->m_bstrQueueName);				
				_ASSERTE(0); // should not reach here 
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, _T("Unknown MESSQ_RESP returned from component"));
				break;
			}
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eReceiveExecute:
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, comerr, _T("Error receiving/executing message"));
				pThreadData->m_rThreadPoolManagerOMMQ1.UnlockMessage(pThreadData->m_vMessageId);
				break;
			case eHandleFailures:
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, comerr, _T("Error handling failures"));
				break;
			default :
				pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, comerr, _T("Unknown error"));
				break;
		}
	}
	catch(...)
	{
		pThreadData->m_rThreadPoolManagerOMMQ1.LogEventError(pszFunctionName, _T("Unknown error"));
	}

	::SysFreeString(bstrErrorMessage);

	delete pThreadData;
}

void CThreadPoolManagerOMMQ1::LogEventErrorProvider(_ConnectionPtr& ConnectionPtr)
{
	const LPCTSTR pszFunctionName = _T("LogEventErrorProvider");

	try
	{
		// Print Provider Errors from Connection object.
		// pErr is a record object in the Connection's Error collection.
		ErrorPtr    pErr  = NULL;
		long      nCount  = 0;    
		long      i     = 0;

		if( (ConnectionPtr->Errors->Count) > 0)
		{
			nCount = ConnectionPtr->Errors->Count;
			// Collection ranges from 0 to nCount -1.
			for(i = 0; i < nCount; i++)
			{
				pErr = ConnectionPtr->Errors->GetItem(i);
				TCHAR buff[MAXLOGMESSAGESIZE];
				_bstr_t bstrDescription = pErr->Description;
				VERIFY(_sntprintf(buff, MAXLOGMESSAGESIZE - 1, _T("Error number: %x.  Error Description: %s"), 
					pErr->Number,(LPCTSTR) bstrDescription) > -1);
				LogEventError(pszFunctionName, buff);
			}
		}
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerOMMQ1::UnlockMessage(_variant_t vMessageId)
{
	const LPCTSTR pszFunctionName = _T("UnlockMessage");

	_ConnectionPtr ConnectionPtr;
	_CommandPtr CommandPtr;
	enum
	{
		eNULL,
		eOpenConnection,
		eSetupCommand,
		eSetupParameters,
		eExecuteCommand,
	} eAction = eNULL;
	
	try
	{
		// get the message (ProgID and XML) defined by its message id
		// ... open the connection
		eAction = eOpenConnection;
		ConnectionPtr.CreateInstance(__uuidof(Connection));
		ConnectionPtr->ConnectionString = m_bstrConnectionString;
		ConnectionPtr->CursorLocation = adUseClient;
		ConnectionPtr->Open("", "", "", 0);

		// ... set up the ADO command
		eAction = eSetupCommand;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = ConnectionPtr;
		
		// ... set up the parameters for the commmand
		eAction = eSetupParameters;
		_ParameterPtr ParameterPtr;

		switch (m_lProvider)
		{
			case PROVIDER_MSDAORA:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.UnlockMessage(?)}";
				ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMessageId);
				break;
			case PROVIDER_SQLOLEDB:
				CommandPtr->CommandType = adCmdStoredProc;
				// RF BMIDS727 Start
				//CommandPtr->CommandText = "USP_MQLOMMQUNLOCKMESSAGE";
				CDbHelperOMMQ* ptrHelper;
				ptrHelper = new CDbHelperOMMQ(); 
				if (ptrHelper == NULL)
				{
					_ASSERTE(0); // should not reach here
					throw L"Unable to set command text";
				}
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					m_bstrTableSuffix,
					CDbHelperOMMQ::eUnLockMessage);
				ptrHelper = NULL;
				// RF BMIDS727 End								
				ParameterPtr = CommandPtr->CreateParameter(
					s_bstrNULL, adBinary, adParamInput, 16, vMessageId);
				break;
			case PROVIDER_ORAOLEDB:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.UnlockMessage(?)}";
				ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMessageId);
				break;
			default :
				_ASSERTE(0); // should not reach here 
				throw L"Unrecognised provider";
				break;
		}
		CommandPtr->GetParameters()->Append(ParameterPtr);

		// ... execute command
		eAction = eExecuteCommand;
		CommandPtr->Execute(NULL, NULL, adExecuteNoRecords);

		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(_com_error comerr)
	{
		if (m_bDBConnectivityGood)
		{
			LogEventErrorProvider(ConnectionPtr);
			switch (eAction)
			{
				case eOpenConnection:
					LogEventError(pszFunctionName, comerr, _T("Error opening connection. Attempting to re-establish database connectivity"));
					OnBadConnection();
					break;
				case eSetupCommand:
					LogEventError(pszFunctionName, comerr,_T("Error setting up command"));
					break;
				case eSetupParameters:
					LogEventError(pszFunctionName, comerr,_T("Error setting up parameters"));
					break;
				case eExecuteCommand:
					LogEventError(pszFunctionName, comerr,_T("Error executing command"));
					break;
				default :
					LogEventError(pszFunctionName, comerr,_T("Unknown error"));
					break;
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(...)
	{
		if (m_bDBConnectivityGood)
		{
			if (eAction == eOpenConnection)
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, _T("Unknown error opening connection. Attempting to re-establish database connectivity"));
				OnBadConnection();
			}
			else
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, _T("Unknown error"));
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
}


bool CThreadPoolManagerOMMQ1::ResetQueue()
{
	const LPCTSTR pszFunctionName = _T("ResetQueue");

	bool bResult = false;
	_ConnectionPtr ConnectionPtr;
	_CommandPtr CommandPtr;
	enum
	{
		eNULL,
		eOpenConnection,
		eSetupCommand,
		eSetupParameters,
		eExecuteCommand,
	} eAction = eNULL;
	
	try
	{
		// get the message (ProgID and XML) defined by its message id
		// ... open the connection
		eAction = eOpenConnection;
		ConnectionPtr.CreateInstance(__uuidof(Connection));
		ConnectionPtr->ConnectionString = m_bstrConnectionString;
		ConnectionPtr->CursorLocation = adUseClient;
		ConnectionPtr->Open("", "", "", 0);

		// ... set up the ADO command
		eAction = eSetupCommand;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = ConnectionPtr;
		
		switch (m_lProvider)
		{
			case PROVIDER_MSDAORA:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.ResetQueue(?)}";
				break;
			case PROVIDER_SQLOLEDB:
				CommandPtr->CommandType = adCmdStoredProc;
				// RF BMIDS727 Start
				//CommandPtr->CommandText = "USP_MQLOMMQRESETQUEUE";
				CDbHelperOMMQ* ptrHelper;
				ptrHelper = new CDbHelperOMMQ(); 
				if (ptrHelper == NULL)
				{
					_ASSERTE(0); // should not reach here
					throw L"Unable to set command text";
				}
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					m_bstrTableSuffix,
					CDbHelperOMMQ::eResetQueue);
				ptrHelper = NULL;
				// RF BMIDS727 End								
				break;
			case PROVIDER_ORAOLEDB:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.ResetQueue(?)}";
				break;
			default :
				_ASSERTE(0); // should not reach here 
				throw L"Unrecognised provider";
				break;
		}
		
		// ... set up the parameters for the commmand
		eAction = eSetupParameters;
		_ParameterPtr ParameterPtr;

		ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, SysStringLen(m_bstrQueueName), m_bstrQueueName);
		CommandPtr->GetParameters()->Append(ParameterPtr);

		// ... execute command
		eAction = eExecuteCommand;
		CommandPtr->Execute(NULL, NULL, adExecuteNoRecords);

		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
		bResult = true;
	}
	catch(_com_error comerr)
	{
		if (m_bDBConnectivityGood)
		{
			LogEventErrorProvider(ConnectionPtr);
			switch (eAction)
			{
				case eOpenConnection:
					LogEventError(pszFunctionName, comerr, _T("Error opening connection. Attempting to re-establish database connectivity"));
					OnBadConnection();
					break;
				case eSetupCommand:
					LogEventError(pszFunctionName, comerr, _T("Error setting up command"));
					break;
				case eSetupParameters:
					LogEventError(pszFunctionName, comerr, _T("Error setting up parameters"));
					break;
				case eExecuteCommand:
					LogEventError(pszFunctionName, comerr, _T("Error executing command"));
					break;
				default :
					LogEventError(pszFunctionName, comerr, _T("Unknown error"));
					break;
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(...)
	{
		if (m_bDBConnectivityGood)
		{
			if (eAction == eOpenConnection)
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, _T("Unknown error opening connection. Attempting to re-establish database connectivity"));
				OnBadConnection();
			}
			else
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, _T("Unknown error"));
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	return bResult;
}


bool CThreadPoolManagerOMMQ1::UnlockMessages()
{
	const LPCTSTR pszFunctionName = _T("UnlockMessages");

	bool bResult = false;
	_ConnectionPtr ConnectionPtr;
	_CommandPtr CommandPtr;
	enum
	{
		eNULL,
		eOpenConnection,
		eSetupCommand,
		eSetupParameters,
		eExecuteCommand,
	} eAction = eNULL;
	
	try
	{
		// get the message (ProgID and XML) defined by its message id
		// ... open the connection
		eAction = eOpenConnection;
		ConnectionPtr.CreateInstance(__uuidof(Connection));
		ConnectionPtr->ConnectionString = m_bstrConnectionString;
		ConnectionPtr->CursorLocation = adUseClient;
		ConnectionPtr->Open("", "", "", 0);

		// ... set up the ADO command
		eAction = eSetupCommand;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = ConnectionPtr;
		
		switch (m_lProvider)
		{
			case PROVIDER_MSDAORA:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.UnlockMessages(?,?)}";
				break;
			case PROVIDER_SQLOLEDB:
				CommandPtr->CommandType = adCmdStoredProc;
				// RF BMIDS727 Start
				//CommandPtr->CommandText = "USP_MQLOMMQUNLOCKMESSAGES";
				CDbHelperOMMQ* ptrHelper;
				ptrHelper = new CDbHelperOMMQ(); 
				if (ptrHelper == NULL)
				{
					_ASSERTE(0); // should not reach here
					throw L"Unable to set command text";
				}
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					m_bstrTableSuffix,
					CDbHelperOMMQ::eUnlockMessages);
				ptrHelper = NULL;
				// RF BMIDS727 End								
				break;
			case PROVIDER_ORAOLEDB:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.UnlockMessages(?,?)}";
				break;
			default :
				_ASSERTE(0); // should not reach here 
				throw L"Unrecognised provider";
				break;
		}
		
		// ... set up the parameters for the commmand
		eAction = eSetupParameters;
		_ParameterPtr ParameterPtr;

		ParameterPtr = CommandPtr->CreateParameter(
			s_bstrNULL, adBSTR, adParamInput, SysStringLen(m_bstrQueueName), m_bstrQueueName);
		CommandPtr->GetParameters()->Append(ParameterPtr);
		ParameterPtr = CommandPtr->CreateParameter(
			s_bstrNULL, adBSTR, adParamInput, SysStringLen(m_bstrLockedBy), m_bstrLockedBy);
		CommandPtr->GetParameters()->Append(ParameterPtr);

		// ... execute command
		eAction = eExecuteCommand;
		CommandPtr->Execute(NULL, NULL, adExecuteNoRecords);

		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
		bResult = true;
	}
	catch(_com_error comerr)
	{
		// RF 02/04/04	BMIDS727 - Allow for multiple queue tables - improved error handling

		// check for error caused by unhandled table suffix
		
		if (m_lProvider == PROVIDER_SQLOLEDB && 
			eAction == eExecuteCommand &&
			comerr.Error() == DB_E_ERRORSINCOMMAND)
		{
			LogEventError(pszFunctionName, comerr, 
				_T("Error executing command on queue '%s' for table suffix '%s'"),
				(LPWSTR)m_bstrQueueName, (LPWSTR)m_bstrTableSuffix);
		}
		else
		{
			if (m_bDBConnectivityGood)
			{
				LogEventErrorProvider(ConnectionPtr);
				switch (eAction)
				{
					case eOpenConnection:
						LogEventError(pszFunctionName, comerr, 
							_T("Error opening connection. Attempting to re-establish database connectivity"));
						OnBadConnection();
						break;
					case eSetupCommand:
						LogEventError(pszFunctionName, comerr, _T("Error setting up command"));
						break;
					case eSetupParameters:
						LogEventError(pszFunctionName, comerr, _T("Error setting up parameters"));
						break;
					case eExecuteCommand:
						LogEventError(pszFunctionName, comerr, _T("Error executing command"));
						break;
					default :
						LogEventError(pszFunctionName, comerr, _T("Unknown error"));
						break;
				}
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(...)
	{
		if (m_bDBConnectivityGood)
		{
			if (eAction == eOpenConnection)
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, 
					_T("Unknown error opening connection. Attempting to re-establish database connectivity"));
				OnBadConnection();
			}
			else
			{
				LogEventErrorProvider(ConnectionPtr);
				LogEventError(pszFunctionName, _T("Unknown error"));
			}
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	return bResult;
}

void CThreadPoolManagerOMMQ1::OnBadConnection()
{
	m_bDBConnectivityGood = false;
	if (!m_bProcessingOpenKeepAlive && m_bQueueStarted)
	{
		PostThreadMessage(m_dwThreadId, WM_OMMQ1_OPENKEEPALIVE, 0, 0);
		Sleep(s_dwPollConnection);
	}
}

void CThreadPoolManagerOMMQ1::LogEventInfo(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
    try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		_Module.LogEventInfo(_T("OMMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerOMMQ1::LogEventWarning(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
    try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		_Module.LogEventWarning(_T("OMMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerOMMQ1::LogEventWarning(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...)
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
			_Module.LogEventWarning(_T("OMMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', ErrorMessage = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.ErrorMessage());
		}
		else
		{
			// output description
			_Module.LogEventWarning(_T("OMMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', Source = '%s', Description = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
		}
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerOMMQ1::LogEventError(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
    try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > - 1);
		va_end(pArg);

		_Module.LogEventError(_T("OMMQ1 - (%s) %s (ThreadId = %d, QueueName = '%ls')"), pszFunctionName, chMsg, GetCurrentThreadId(), (LPWSTR)m_bstrQueueName);
	}
	catch(...)
	{
	}
}

void CThreadPoolManagerOMMQ1::LogEventError(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...)
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
			_Module.LogEventError(_T("OMMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', ErrorMessage = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.ErrorMessage());
		}
		else
		{
			// output description
			_Module.LogEventError(_T("OMMQ1 - (%s) %s (ThreadId = %d, HResult = %d, QueueName = '%ls', Source = '%s', Description = '%s')"), pszFunctionName, chMsg, GetCurrentThreadId(), comerr.Error(), (LPWSTR)m_bstrQueueName, (LPCTSTR)comerr.Source(), (LPCTSTR)comerr.Description());
		}
	}
	catch(...)
	{
	}
}

_bstr_t CThreadPoolManagerOMMQ1::MessageIdToBstrt(_variant_t vId)
{
	if (vId.vt == (VT_ARRAY | VT_UI1))
	{
		WCHAR wszId[50] = {'\0'};
		SAFEARRAY* psa = ((VARIANT)vId).parray;

		unsigned char* p20Bytes = NULL;
		HRESULT hr = ::SafeArrayAccessData(psa, (void HUGEP**)&p20Bytes);
		if (SUCCEEDED(hr))
		{
			const int nBytesMax = 16;
			if (psa->rgsabound->cElements == nBytesMax)
			{
				for (int nBytes = 0; nBytes < nBytesMax; nBytes++)
				{
					swprintf(wszId + nBytes * 2, L"%02x", p20Bytes[nBytes]);
				}
			}
			::SafeArrayUnaccessData(psa);
		}
		return _bstr_t(wszId);	
	}
	else if (vId.vt == VT_BSTR)
	{
		return _bstr_t(vId);
	}
	else
	{
		_ASSERTE(0); // should not reach here
		return _bstr_t(L"");
	}
}

void CThreadPoolManagerOMMQ1::CloseConnection(_ConnectionPtr& ConnectionPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseConnection");

	try
	{
		if (ConnectionPtr != NULL && ConnectionPtr->State == adStateOpen)
		{
			ConnectionPtr->Close();
		}
		ConnectionPtr = NULL;
	}
	catch(...)
	{
		try
		{
			ConnectionPtr = NULL;
		}
		catch(...)
		{
		}
	}
}

void CThreadPoolManagerOMMQ1::CloseRecordset(_RecordsetPtr& RecordsetPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseRecordset");

	try
	{
		if (RecordsetPtr != NULL && RecordsetPtr->State == adStateOpen)
		{
			RecordsetPtr->Close();
		}
		RecordsetPtr = NULL;
	}
	catch(...)
	{
		try
		{
			RecordsetPtr = NULL;
		}
		catch(...)
		{
		}
	}
}

void CThreadPoolManagerOMMQ1::CloseCommand(_CommandPtr& CommandPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseCommand");

	try
	{
		CommandPtr = NULL;
	}
	catch(...)
	{
	}
}
