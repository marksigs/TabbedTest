///////////////////////////////////////////////////////////////////////////////
//	FILE:			ScheduleManager.h
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

#if !defined(AFX_SCHEDULEMANAGER_H__230CAB5C_6DE4_11D4_8247_005004E8D1A7__INCLUDED_)
#define AFX_SCHEDULEMANAGER_H__230CAB5C_6DE4_11D4_8247_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ThreadPoolManager.h"
#include "Log.h"

///////////////////////////////////////////////////////////////////////////////

class CScheduleManager : private CThreadPoolManager, public CLog
{
public:
	CScheduleManager();
	virtual ~CScheduleManager();

	class CEvents : public CObject
	{
	public:
		enum TeOnDestruct {eOnDestructNULL, eOnDestructInternalEnable, eOnDestructExternalEnable, eOnDestructExternalDisable};

		CEvents(CScheduleManager& rScheduleManager, TeOnDestruct eOnDestruct) :
			m_rScheduleManager(rScheduleManager),
			m_eOnDestruct(eOnDestruct)
		{
		}

		~CEvents()
		{
			switch (m_eOnDestruct)
			{
				case eOnDestructInternalEnable:
					m_rScheduleManager.InternalEnableEvents();
					break;
				case eOnDestructExternalEnable:
					m_rScheduleManager.ExternalEnableEvents();
					break;
				case eOnDestructExternalDisable:
					m_rScheduleManager.ExternalDisableEvents();
					break;
				default :
					break;
			}
		}

		void SeteOnDestruct(TeOnDestruct eOnDestruct) {m_eOnDestruct = eOnDestruct;}

	private:
		CScheduleManager& m_rScheduleManager;
		TeOnDestruct m_eOnDestruct;
	};

// CThreadPoolManager
public:
	// initialisation (creates/destroys scheduling thread, and starts the schedluling)
    virtual bool StartUp();
    virtual void CloseDown();

	enum { eInfiniteRepeatCount = -1};
	void SetRepeatCount(int nRepeatCount = eInfiniteRepeatCount) {m_nRequestedRepeatCount = nRepeatCount;}

	virtual bool ExternalDisableEvents(bool bWait = true); // temporarily disables scheduling (without killing scheduling thread) (to be called by another thread)
	virtual bool ExternalEnableEvents();  // renables the scheduling (to be called by another thread)
	virtual bool InternalEnableEvents();  // renables the scheduling (to be called by the scheduling thread)

	CThreadPoolThread* TryGetpAvailableThreadPoolThread(); // returns NULL if thread is not available

private:
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread); // returns NULL if there is no more work to be done 

protected:
	// overrideables
	virtual DWORD GetdwNextWaitIntervalms() = 0; // in milliseconds
	virtual void OnEventSchedule() = 0;

private:
    HANDLE m_hDisableEvent; // event raised to disable scheduling

	int m_nRequestedRepeatCount;
	int m_nRepeatCountRemaining; 
	bool m_bInternalEventsCalled;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_SCHEDULEMANAGER_H__230CAB5C_6DE4_11D4_8247_005004E8D1A7__INCLUDED_)
