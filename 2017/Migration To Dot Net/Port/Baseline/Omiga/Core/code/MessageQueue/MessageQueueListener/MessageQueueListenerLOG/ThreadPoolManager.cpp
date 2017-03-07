///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManager.cpp
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
#include "ThreadPoolThread.h"
#include "ThreadPoolManager.h"
//#include "EserverPrfMon.h"

using namespace namespaceMutex;

DWORD CThreadPoolManager::s_dwNumberOfProcessors = CThreadPoolManager::DetermineNumberOfProcessors();


///////////////////////////////////////////////////////////////////////////////

// Note that it is safe using EnterCriticalSection/LeaveCriticalSection as there is no chance of an exception being raised
// (otherwise must use a class   e.g CSingleLock)

// Note that generic containers are not used for the queue to increase efficiency (less blocking)


///////////////////////////////////////////////////////////////////////////////

CThreadPoolManager::CThreadPoolManager() :
	m_pThreadPoolThreadHead(NULL),
	m_lHitCtr(0),
	m_lHitCtrTotal(0),
	m_lHitCtrPercent(0),
	m_bThreadPoolManagerStarted(false),
	m_nActualNumberOfThreads(0),
	m_nWorkerThreadPriority(THREAD_PRIORITY_NORMAL),
	m_lThreadsCreated(0),
	m_hLastThreadDestroyedEvent(NULL),
	m_hThreadLastToExit(NULL)
{
	m_nRequestedNumberOfThreads = (s_dwNumberOfProcessors * 2 + 1);
}

CThreadPoolManager::~CThreadPoolManager()
{
}

///////////////////////////////////////////////////////////////////////////////

DWORD CThreadPoolManager::DetermineNumberOfProcessors()
{
    SYSTEM_INFO sysinfo;
    ::GetSystemInfo(&sysinfo);
    return sysinfo.dwNumberOfProcessors;
}

DWORD CThreadPoolManager::GetNextIdealProcessor()
{
    static DWORD dwNextIdealProcessor = 0;
    DWORD dwResult = dwNextIdealProcessor++;
    if (dwNextIdealProcessor > s_dwNumberOfProcessors)
    {
        dwNextIdealProcessor = 0;
    }
    return dwResult;
}

CThreadPoolThread* CThreadPoolManager::CreateSuspendedThread() // returns NULL if failed (object returned is valid until CloseDownAndDelete())
{
	return CThreadPoolThread::CreateSuspendedThread(this);
}

bool CThreadPoolManager::StartUp(int nNumberOfThreads)
{
	if (nNumberOfThreads != eDefaultNumberOfThreads)
	{
		m_nRequestedNumberOfThreads = nNumberOfThreads;
	}
    DWORD dwNumberOfThreadsToTryCreate = m_nRequestedNumberOfThreads;

    CThreadPoolThread* pThreadPoolThread = NULL;
    CThreadPoolThread* pThreadPoolThreadPrev = NULL;

	CWriteLock lckResize(m_sharedlockResizeThreads);
    
    // create a circular list of suspended threads
    while (dwNumberOfThreadsToTryCreate > 0)
    {
        pThreadPoolThread = CreateSuspendedThread();
        if (pThreadPoolThread)
        {
            m_nActualNumberOfThreads++;

			// succesfully created a suspended thread, so add to list
            if (pThreadPoolThreadPrev)
            {
                pThreadPoolThreadPrev->SetNextInSequence(pThreadPoolThread);
            }
            else
            {
                m_pThreadPoolThreadHead = pThreadPoolThread;
            }

            pThreadPoolThreadPrev = pThreadPoolThread;

            // allocate it to a processor (i.e for multi-processor machines)
            DWORD dw = pThreadPoolThread->SetThreadIdealProcessor(GetNextIdealProcessor());
            _ASSERTE(dw != -1);
        }
        
        dwNumberOfThreadsToTryCreate--;
    };

    if (pThreadPoolThread)
    {
        pThreadPoolThread->SetNextInSequence(m_pThreadPoolThreadHead);
    }
    
	// signal that the thread-pool manager has started
	m_bThreadPoolManagerStarted = true;

    // return true if there is atleast one thread created
    return (pThreadPoolThread != NULL);
};

void CThreadPoolManager::CloseDown()
{
	CWriteLock lckResize(m_sharedlockResizeThreads);
    
	// initialise queue
    if (m_pThreadPoolThreadHead && m_bThreadPoolManagerStarted)
    {
        // tell the threads to close down
        CThreadPoolThread* pNextThreadPoolThread = m_pThreadPoolThreadHead;
        CThreadPoolThread* pThisThreadPoolThread = NULL;
        do
        {
			pThisThreadPoolThread = pNextThreadPoolThread;
            pNextThreadPoolThread = pThisThreadPoolThread->GetNextInSequence();
            pThisThreadPoolThread->CloseDownAndDelete();
			m_nActualNumberOfThreads--;
        } while (pNextThreadPoolThread != m_pThreadPoolThreadHead);

        // wait for them
		lckResize.UnLock();
        WaitForAllThreadsToCloseDown();
		m_bThreadPoolManagerStarted = false;
    }
}


void CThreadPoolManager::IncreaseThreads(int nNumberOfThreads)
{
	if (m_bThreadPoolManagerStarted)
	{
		CReadLock lckResize(m_sharedlockResizeThreads); // a read lock is sufficient

		CThreadPoolThread* pThreadPoolThread = NULL;
		CThreadPoolThread* pThreadPoolThreadPrev = m_pThreadPoolThreadHead;

		// insert into a circular list of threads (threads are still active)
		while (nNumberOfThreads > 0)
		{
			pThreadPoolThread = CreateSuspendedThread();
			if (pThreadPoolThread)
			{
				m_nActualNumberOfThreads++;

				// succesfully created a suspended thread, so add to list
				if (pThreadPoolThreadPrev)
				{
					pThreadPoolThread->SetNextInSequence(pThreadPoolThreadPrev->GetNextInSequence());
					pThreadPoolThreadPrev->SetNextInSequence(pThreadPoolThread);
				}
				else
				{
					pThreadPoolThread->SetNextInSequence(pThreadPoolThread); // circular list of one thread
					m_pThreadPoolThreadHead = pThreadPoolThread;
				}

				pThreadPoolThreadPrev = pThreadPoolThread;

				// allocate it to a processor (i.e for multi-processor machines)
				DWORD dw = pThreadPoolThread->SetThreadIdealProcessor(GetNextIdealProcessor());
				_ASSERTE(dw != -1);

                // an extra thread is now available, tell it to do any queued work
				// ... thread will go back to sleep if the queue is empty
                pThreadPoolThread->SetRuntimeFunction(NULL); 
				pThreadPoolThread->SetThreadAvailableForReuse(FALSE);
                pThreadPoolThread->ResumeThread();
			}
        
			nNumberOfThreads--;
		};
	}
}

void CThreadPoolManager::DecreaseThreads(int nNumberOfThreads)
{
	if (m_bThreadPoolManagerStarted)
	{
		CWriteLock lckResize(m_sharedlockResizeThreads);

		// remove from a circular list of threads (threads are still active)
		while (nNumberOfThreads > 0)
		{
			CThreadPoolThread* pThreadPoolThread = m_pThreadPoolThreadHead;
			if (pThreadPoolThread) // guard against zero threads
			{
				CThreadPoolThread* pThreadPoolThreadNext = m_pThreadPoolThreadHead->GetNextInSequence();
				if (pThreadPoolThread == pThreadPoolThreadNext)
				{
					// only one thread
					pThreadPoolThread->CloseDownAndDelete();
					m_nActualNumberOfThreads--;
					m_pThreadPoolThreadHead = NULL;
				}
				else
				{
					// two or more threads
					pThreadPoolThread->SetNextInSequence(pThreadPoolThreadNext->GetNextInSequence());
					pThreadPoolThreadNext->CloseDownAndDelete();
					m_nActualNumberOfThreads--;
				}
			}
			
			nNumberOfThreads--;
		}
	}
}

int CThreadPoolManager::GetnRequestedNumberOfThreads()
{
	return m_nRequestedNumberOfThreads;
}

void CThreadPoolManager::SetnRequestedNumberOfThreads(int nNumberOfThreads)
{
	m_nRequestedNumberOfThreads	= nNumberOfThreads;
	if (m_bThreadPoolManagerStarted)
	{
		if (m_nRequestedNumberOfThreads > m_nActualNumberOfThreads)
		{
			IncreaseThreads(m_nRequestedNumberOfThreads - m_nActualNumberOfThreads);
		}
		else if (m_nRequestedNumberOfThreads < m_nActualNumberOfThreads)
		{
			DecreaseThreads(m_nActualNumberOfThreads - m_nRequestedNumberOfThreads);
		}
	}
}

int CThreadPoolManager::GetnThreadsActive()
{
	int nThreadsActive = 0;
	CReadLock lckResize(m_sharedlockResizeThreads); // a read lock is sufficient
	CThreadPoolThread* pThreadPoolThread = m_pThreadPoolThreadHead;
	if (pThreadPoolThread)
	{
		do
		{
			if (pThreadPoolThread->IsThreadAvailableForReuse() == FALSE)
			{
				nThreadsActive++;
			}
			pThreadPoolThread = pThreadPoolThread->GetNextInSequence();
		} while (pThreadPoolThread != m_pThreadPoolThreadHead);
	}
	return nThreadsActive;
}

int CThreadPoolManager::GetnThreadsNotActive()
{
	int nThreadsActive = 0;
	CReadLock lckResize(m_sharedlockResizeThreads); // a read lock is sufficient
	CThreadPoolThread* pThreadPoolThread = m_pThreadPoolThreadHead;
	if (pThreadPoolThread)
	{
		do
		{
			if (pThreadPoolThread->IsThreadAvailableForReuse() == TRUE)
			{
				nThreadsActive++;
			}
			pThreadPoolThread = pThreadPoolThread->GetNextInSequence();
		} while (pThreadPoolThread != m_pThreadPoolThreadHead);
	}
	return nThreadsActive;
}

void CThreadPoolManager::IncHitCtr(BOOL bHit)
{
//	if (g_pEServerPrfMon->GetEnabled())
//	{
//		if (m_lHitCtrTotal + 1 == INT_MAX)
//		{
//			// counter about to overflow, so reset
//			m_lHitCtr		= m_lHitCtrPercent;
//			m_lHitCtrTotal	= 100;
//		}
//		if (bHit)
//		{
//			m_lHitCtr++;
//		}
//		m_lHitCtrTotal++;
//		m_lHitCtrPercent = m_lHitCtr == 0 ? 0 : (LONG)(((double)m_lHitCtr / (double)m_lHitCtrTotal) * 100);
//		g_pEServerPrfMon->GetCtr32(THREADPOOL_HIT, g_pEServerPrfMon->GetInstIdThreadPool()) = m_lHitCtrPercent;
//	}
}


long CThreadPoolManager::OnThreadCreated()
{
    long lInstances = ::InterlockedIncrement(&m_lThreadsCreated);
    if (lInstances == 1)
    {
        m_hLastThreadDestroyedEvent = ::CreateEvent(NULL, TRUE, FALSE, NULL);
    }
	return lInstances;
}

long CThreadPoolManager::OnThreadDestroyed(HANDLE hThreadClosingDown)
{
    // keep a track of the number of instances
    long lInstances = ::InterlockedDecrement(&m_lThreadsCreated);
    if (lInstances == 0)
    {
        // need to know the last thread that exits
        // (note that this last handle is closed in WaitForAllThreadsToCloseDown)
        m_hThreadLastToExit = hThreadClosingDown;

        // signal that all threads have terminated
        VERIFY(::SetEvent(m_hLastThreadDestroyedEvent));
    }
    else
    {
        VERIFY(::CloseHandle(hThreadClosingDown));
    }
	return lInstances;
}


void CThreadPoolManager::WaitForAllThreadsToCloseDown()
{
    if (m_lThreadsCreated > 0)
    {
        DWORD dwWait = ::WaitForSingleObject(m_hLastThreadDestroyedEvent, INFINITE);
        _ASSERTE(dwWait == WAIT_OBJECT_0);
        VERIFY(::CloseHandle(m_hLastThreadDestroyedEvent));
        m_hLastThreadDestroyedEvent = NULL;

        // wait for last thread - there's a finite time between event set by last thread to when it exits
        DWORD dwExitCode;
        do
        {
            ::GetExitCodeThread(m_hThreadLastToExit, &dwExitCode);
            ::Sleep(0);
        } while (dwExitCode == STILL_ACTIVE);
		// handles from other threads are closed by the thread itself 
		VERIFY(::CloseHandle(m_hThreadLastToExit));
    }
}


CThreadPoolThread* CThreadPoolManager::TryGetpAvailableThreadPoolThread()
{
    BOOL bAvailableThread = FALSE;
	CReadLock lckResize(m_sharedlockResizeThreads); // a read lock is sufficient
	CThreadPoolThread* pThreadPoolThreadHead = m_pThreadPoolThreadHead; // make copy for local use
    CThreadPoolThread* pThreadPoolThread = m_pThreadPoolThreadHead;
    do
    {
        pThreadPoolThread = pThreadPoolThread->GetNextInSequence();
        bAvailableThread = pThreadPoolThread->TryToReuseThread();
    } while ((bAvailableThread == FALSE) && (pThreadPoolThread != pThreadPoolThreadHead));

	if (bAvailableThread)
	{
		// change the head of the circular queue to balance thread loading
        m_pThreadPoolThreadHead = pThreadPoolThread;
	}
	else
	{
		// return NULL
		pThreadPoolThread = NULL;
	}
	return pThreadPoolThread;
}
