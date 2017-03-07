///////////////////////////////////////////////////////////////////////////////
//	FILE:			ScheduleManagerDaily1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AD      30/08/00    Initial version
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
///////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////
//
//	AD 15/08/2000
//	This class is derived from CScheduleManager. Creates
//	and maintains a circular list of CScheduleEventDaily items. Event 
//	is executed when time and date criteria are met.
//
/////////////////////////////////////////////////////////////////////

#if !defined(AFX_SCHEDULEMANAGERDAILY1_H__D80652A3_728D_11D4_825D_000102125FBA__INCLUDED_)
#define AFX_SCHEDULEMANAGERDAILY1_H__D80652A3_728D_11D4_825D_000102125FBA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ScheduleManager.h"
#include "ScheduleEventDaily.h"
#include "stdafx.h"


class CScheduleManagerDaily : public CScheduleManager  
{
public:
	DWORD GetdwNextWaitIntervalms();
	CScheduleManagerDaily();
	virtual ~CScheduleManagerDaily();
 	void OnEventSchedule();


public:
    virtual bool StartUp();
    virtual void CloseDown();

	bool AddEventToList(CMQEvent* pNewMQEvent,int nHour, 
					   int nMinute, enum DayOfWeek eDay);
	bool RemoveAll(void);
	CScheduleEventDaily* m_pActiveEvent;
	CScheduleEventDaily* m_pPrevEventActioned;
    namespaceMutex::CCriticalSection m_csQueue;
	bool RemoveQueueEvents(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr);
	void GetEventInfo(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, XmlNS::IXMLDOMDocument *ptrDomOut);
	bool RemoveEvent(XmlNS::IXMLDOMDocument *ptrDomOut, GUID gKey);
};

#endif // !defined(AFX_SCHEDULEMANAGERDAILY1_H__D80652A3_728D_11D4_825D_000102125FBA__INCLUDED_)
