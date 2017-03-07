// ScheduleEventDailyIncreaseThreads.h: interface for the CScheduleEventDailyIncreaseThreads class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_SCHEDULEEVENTDAILYINCREASETHREADS_H__63113B46_83D2_11D4_8267_000102125FBA__INCLUDED_)
#define AFX_SCHEDULEEVENTDAILYINCREASETHREADS_H__63113B46_83D2_11D4_8267_000102125FBA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ScheduleEventDaily.h"

//class CEvent : public CObject
//{
//public:
//	CScheduleEventDailyIncreaseThreads();
//	virtual ~CScheduleEventDailyIncreaseThreads();


//	virtual OnSchedule() = 0;

//};

class CEventIncreaseThreads : public CScheduleEventDaily  
{
public:
	CScheduleEventDailyIncreaseThreads();
	virtual ~CScheduleEventDailyIncreaseThreads();


	virtual OnSchedule();

};
#endif // !defined(AFX_SCHEDULEEVENTDAILYINCREASETHREADS_H__63113B46_83D2_11D4_8267_000102125FBA__INCLUDED_)
