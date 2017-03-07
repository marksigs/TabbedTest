///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerFunctionQueue.cpp
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
#include "mutex.h"
using namespace namespaceMutex;
#include "ThreadPoolThread.h"
#include "ThreadPoolManagerFunctionQueue.h"
#include "ThreadPoolMessageFunctionQueue.h"

using namespace namespaceMutex;

///////////////////////////////////////////////////////////////////////////////

CThreadPoolManagerFunctionQueue::CThreadPoolManagerFunctionQueue() :
	m_pMessageHead(NULL),
	m_pMessageTail(NULL),
	m_csQueue(eSpinCount)
{

}

CThreadPoolManagerFunctionQueue::~CThreadPoolManagerFunctionQueue()
{

}

bool CThreadPoolManagerFunctionQueue::AddFunctionToQueue(CThreadPoolMessageFunctionQueue* pNewThreadPoolMessage)
{
    bool bResult = false;

    CReadLock lckResize(GetrsharedlockResizeThreads());
    if (m_pThreadPoolThreadHead)
    {
        while (bResult == false)
        {
			if (m_pMessageHead == NULL)
            {
                // queue is empty
                // ... threads will either be suspended or about to suspend
                // ... try giving it directly to a thread without queuing
                CThreadPoolThread* pThreadPoolThread = TryGetpAvailableThreadPoolThread();
                if (pThreadPoolThread)
                {
                    // thread is available, tell it to do some work
                    pThreadPoolThread->SetRuntimeFunction(pNewThreadPoolMessage); // thread gains ownership of the message
                    bResult = pThreadPoolThread->ResumeThread();
                
                    // change the head of the circular queue to balance thread loading
                    m_pThreadPoolThreadHead = pThreadPoolThread;
					IncHitCtr(TRUE);
                }
                else
                {
                    // thread is not available, queue the function
                    CSingleLock lck(&m_csQueue, TRUE);

					// need to retest this as a thread may have just added a message in the queue
                    if (m_pMessageTail == NULL) 
                    {
                        // need to restest availability of threads, and try again if necessary
                        // (all threads now available would cause the thread pool to stall)
                        if (m_pThreadPoolThreadHead->IsThreadAvailableForReuse() == FALSE)
                        {
                            // ... queue is empty
                            m_pMessageTail = m_pMessageHead = pNewThreadPoolMessage;
        					IncHitCtr(FALSE);
                            bResult = true;
                        }
                    }
                    else
                    {
                        // ... queue is not empty
                        pNewThreadPoolMessage->SetNext(m_pMessageHead);
                        m_pMessageHead->SetPrev(pNewThreadPoolMessage);
                        m_pMessageHead = pNewThreadPoolMessage;
    					IncHitCtr(FALSE);
                        bResult = true;
                    }
					if (bResult)
					{
//						g_pEServerPrfMon->IncCtr32(THREADPOOL_FUNCTION_QUEUE_LENGTH, g_pEServerPrfMon->GetInstIdThreadPool());
					}
                }
            }
            else
            {
                // queue is not empty, queue the function
                CSingleLock lck(&m_csQueue, TRUE);

				// need to retest this as a thread may have just pinched the last message in the queue
                if (m_pMessageHead) 
                {
                    pNewThreadPoolMessage->SetNext(m_pMessageHead);
                    m_pMessageHead->SetPrev(pNewThreadPoolMessage);
                    m_pMessageHead = pNewThreadPoolMessage;
					IncHitCtr(FALSE);
//					g_pEServerPrfMon->IncCtr32(THREADPOOL_FUNCTION_QUEUE_LENGTH, g_pEServerPrfMon->GetInstIdThreadPool());
                    bResult = true;
                }
            }
        }
    }
	else
	{
		// queue the function for later use (when threads are dynamically created)	
		CSingleLock lck(&m_csQueue, TRUE);
        if (m_pMessageHead) 
        {
            // ... queue is not empty
            pNewThreadPoolMessage->SetNext(m_pMessageHead);
            m_pMessageHead->SetPrev(pNewThreadPoolMessage);
            m_pMessageHead = pNewThreadPoolMessage;
			IncHitCtr(FALSE);
//					g_pEServerPrfMon->IncCtr32(THREADPOOL_FUNCTION_QUEUE_LENGTH, g_pEServerPrfMon->GetInstIdThreadPool());
		}
		else
		{
            // ... queue is empty
            m_pMessageTail = m_pMessageHead = pNewThreadPoolMessage;
        	IncHitCtr(FALSE);
		}
	}

    return bResult;
}

// returns NULL if there is no more work to be done
CThreadPoolMessage* CThreadPoolManagerFunctionQueue::RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread) 
{
    CThreadPoolMessageFunctionQueue* pRetrievedThreadPoolMessage = NULL;

	CSingleLock lck(&m_csQueue, TRUE);

    if (m_pMessageTail)
    {
        // queue is not empty, retrieve the function
        if (m_pMessageHead == m_pMessageTail)
        {
            // one function in the queue
            pRetrievedThreadPoolMessage = m_pMessageHead;
            m_pMessageTail = m_pMessageHead = NULL;
        }
        else
        {
            // more than one function in the queue
            pRetrievedThreadPoolMessage = m_pMessageTail;
            m_pMessageTail = pRetrievedThreadPoolMessage->GetPrev();
        }
//		g_pEServerPrfMon->DecCtr32(THREADPOOL_FUNCTION_QUEUE_LENGTH, g_pEServerPrfMon->GetInstIdThreadPool());
    }
    else
    {
        // no more work, so thread is available for reuse
        pThreadPoolThread->SetThreadAvailableForReuse();
//		g_pEServerPrfMon->GetCtr32(THREADPOOL_FUNCTION_QUEUE_LENGTH, g_pEServerPrfMon->GetInstIdThreadPool()) = 0;
    }

    return pRetrievedThreadPoolMessage;
}

