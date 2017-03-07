///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolThread.cpp
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

#include "stdafx.h"
//#include "Eserver.h"
#include "ThreadPoolThread.h"
#include "ThreadPoolManager.h"

#ifdef DEBUG_NEW
#define new DEBUG_NEW
#endif


///////////////////////////////////////////////////////////////////////////////

CThreadPoolThread::CThreadPoolThread(CThreadPoolManager* pThreadPoolManager) :
	m_pThreadPoolManager(pThreadPoolManager),
    m_dwThreadId(NULL),
    m_hThread(NULL),
    m_bThreadAvailableForReuse(TRUE), // thread is created suspended
    m_bKeepThreadAlive(true),
    m_dwSuspendCount(0),
    m_pThreadPoolMessage(NULL),
    m_pThreadPoolThreadNextInSequence(NULL)
{
    // keep a track of the number of instances
	m_pThreadPoolManager->OnThreadCreated();
}

CThreadPoolThread::~CThreadPoolThread()
{
    // reaches here when on closing down, after finishing all of its work
    _ASSERTE(m_pThreadPoolMessage == NULL);

    // keep a track of the number of instances
	m_pThreadPoolManager->OnThreadDestroyed(m_hThread);
}

///////////////////////////////////////////////////////////////////////////////

CThreadPoolThread* CThreadPoolThread::CreateSuspendedThread(CThreadPoolManager* pThreadPoolManager)
{
    // object is only created when a suspended thread is requested
    CThreadPoolThread* pThreadPoolThread = new CThreadPoolThread(pThreadPoolManager);

    DWORD dwThreadId = NULL;
    HANDLE hThread = ::CreateThread(NULL, 0, ThreadEntry, pThreadPoolThread, CREATE_SUSPENDED, &dwThreadId);
#pragma chMSG("beginthreadex???")
    if (hThread == NULL)
    {
        // failed to create a thread - self destruct
        delete pThreadPoolThread;
        pThreadPoolThread = NULL;
    }
    else
    {
        // need to set up ids/handles for use by the thread
        // (must be done with suspended thread because the thread needs these immediately and there would
        // be no guarantee that the this code would be reached first)
		::SetThreadPriority(hThread, pThreadPoolManager->GetWorkerThreadPriority());
        pThreadPoolThread->m_hThread = hThread;
        pThreadPoolThread->m_dwThreadId = dwThreadId;
        ::ResumeThread(hThread); // Note that this is the global routine (in order for m_dwSuspendCount to be correct)
        // the thread entry routine suspends itself, thus creating a suspended thread
    }
     
    return pThreadPoolThread;
}

DWORD WINAPI CThreadPoolThread::ThreadEntry(LPVOID p)
{
    // static function which forwards thead execution to the corresponding instance
    return reinterpret_cast<CThreadPoolThread*>(p)->ThreadEntry();
}

DWORD WINAPI CThreadPoolThread::ThreadEntry()
{
    // when thread is resumed after setting runtime functions this processing will occur
    // (m_pFunction is only changed when this thread is suspended so that it needs no other
    // protection)
    
    HRESULT hr = CoInitialize(NULL);
	if (SUCCEEDED(hr))
	{
		while (m_bKeepThreadAlive || (m_pThreadPoolMessage != NULL))
		{		
			try
			{
				while (m_bKeepThreadAlive || (m_pThreadPoolMessage != NULL))
				{
					// initial entry or work has been done, suspend this thread 
					// (note that the thread will not suspend when closing down
					// due to an extra call to the ResumeThread() i.e  = false
					SuspendThread();

					// will assert if the thread usage has not been allocated with TryToReuseThread
					_ASSERTE(!m_bThreadAvailableForReuse || !m_bKeepThreadAlive);

					// will reach this point when the thread is resumed by the call to
					// ResumeThread in CThreadPoolManager::AddFunctionToQueue().
					// Start the thread performing some work.
					if (m_pThreadPoolMessage)
					{
						m_pThreadPoolMessage->ExecuteAndDelete();
						m_pThreadPoolMessage = NULL;
					}
        
					// while there is more work available, keep executing
					while (m_bKeepThreadAlive && (m_pThreadPoolMessage = m_pThreadPoolManager->RemoveMessageFromQueue(this)) != NULL)
					{
						// will assert if the thread usage has not been allocated with TryToReuseThread
						_ASSERTE(!m_bThreadAvailableForReuse || !m_bKeepThreadAlive);

						m_pThreadPoolMessage->ExecuteAndDelete();
						m_pThreadPoolMessage = NULL;
					};

					// will fire is the thread has not been set for reuse when there is no more work to be done
					//_ASSERTE(!m_bKeepThreadAlive || m_bThreadAvailableForReuse);
				}
			}
			catch(...)
			{
				_Module.LogEventError(_T("ThreadPool - Caught Exception"));
			}
		}
		CoUninitialize();
	}
    
    // self distruct
    delete this;

    return 0;
}

bool CThreadPoolThread::ResumeThread() 
{
	bool bSuccess = false;

	_ASSERTE(IsThreadCreated()); 

	if (IsThreadCreated())
	{
		bSuccess = true;

		if (::InterlockedDecrement((LPLONG)&m_dwSuspendCount) == 0) 
		{
			DWORD dwNTPreviousSuspendCount;
			while ((dwNTPreviousSuspendCount = ::ResumeThread(m_hThread)) == 0)
			{
				// Attempting to resumed a thread which is not suspended,
				// (i.e guard against context switch between ::InterlockedIncrement((LPLONG)&m_dwSuspendCount
				// and ::SuspendThread(m_hThread) in the implementation of CThreadPoolThread::SuspendThread())
				// ...try again later.
				::Sleep(0);
			}

			bSuccess = dwNTPreviousSuspendCount != 0xFFFFFFFF;
		}
	}

	return bSuccess;
}

bool CThreadPoolThread::SuspendThread() 
{
	bool bSuccess = false;

	_ASSERTE(IsThreadCreated()); 

	if (IsThreadCreated())
	{
		bSuccess = true;

		if (::InterlockedIncrement((LPLONG)&m_dwSuspendCount) == 1) 
		{
			bSuccess = ::SuspendThread(m_hThread) != 0xFFFFFFFF;
		} 
	}

	return bSuccess;
}
///////////////////////////////////////////////////////////////////////////////
