///////////////////////////////////////////////////////////////////////////////
//	FILE:			ScheduleManager.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ScheduleManager.h"
#include "ThreadPoolThread.h"

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogInOut.AnyThreadInitialize

///////////////////////////////////////////////////////////////////////////////

CScheduleManager::CScheduleManager() :
	m_nRequestedRepeatCount(eInfiniteRepeatCount),
	m_nRepeatCountRemaining(eInfiniteRepeatCount),
	m_bInternalEventsCalled(false)
{
	m_hDisableEvent = ::CreateEvent(NULL, TRUE, FALSE, NULL);
}

CScheduleManager::~CScheduleManager()
{
	VERIFY(::CloseHandle(m_hDisableEvent));
}

bool CScheduleManager::StartUp()
{
	ANYTHREADLOG_THIS_INOUT(_T("CScheduleManager::StartUp\n"));

	// create a thread pool containing one thread
	const int nNumberOfThreads = 1;
	bool bResult = CThreadPoolManager::StartUp(nNumberOfThreads);
	if (bResult)
	{
		bResult = ExternalEnableEvents();
	}
	return bResult;
}

void CScheduleManager::CloseDown()
{
	ANYTHREADLOG_THIS_INOUT(_T("CScheduleManager::CloseDown\n"));

	const bool bWait = true;
	ExternalDisableEvents(bWait);
	CThreadPoolManager::CloseDown();
}


CThreadPoolMessage* CScheduleManager::RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread) // returns NULL if there is no more work to be done
{
	bool bRepeat = true;
	do
	{
		DWORD dwMilliseconds = GetdwNextWaitIntervalms();
		DWORD dwWait = ::WaitForSingleObject(m_hDisableEvent, dwMilliseconds);
		switch (dwWait)
		{
			case WAIT_TIMEOUT:
				// time to schedule something
				OnEventSchedule();
				if (m_nRepeatCountRemaining != eInfiniteRepeatCount && !m_bInternalEventsCalled && --m_nRepeatCountRemaining == 0)
				{
					bRepeat = false;
				}
				m_bInternalEventsCalled = false;
				break;
			default :
				// scheduling disabled
				m_nRepeatCountRemaining = 0;
				bRepeat = false;
				break;
		}
    } while (bRepeat);

	pThreadPoolThread->SetThreadAvailableForReuse();
	return NULL;
}

bool CScheduleManager::ExternalDisableEvents(bool bWait)
{
	bool bResult = false;
	if (::SetEvent(m_hDisableEvent))
	{
		if (bWait && m_pThreadPoolThreadHead)
		{
			while (m_pThreadPoolThreadHead->IsThreadAvailableForReuse() == FALSE)
			{
				::Sleep(0);
			}
		}
		bResult = true;
	}
	return bResult;
}
bool CScheduleManager::ExternalEnableEvents()
{
	bool bResult = false;	
	m_nRepeatCountRemaining = m_nRequestedRepeatCount;
	if (::ResetEvent(m_hDisableEvent) && m_pThreadPoolThreadHead && m_pThreadPoolThreadHead->TryToReuseThread())
	{
		bResult = m_pThreadPoolThreadHead->ResumeThread();
	}
	return bResult;
}

bool CScheduleManager::InternalEnableEvents()
{
	m_nRepeatCountRemaining = m_nRequestedRepeatCount;
	m_bInternalEventsCalled = true;
	return true;
}
