///////////////////////////////////////////////////////////////////////////////
//	FILE:			ScheduleEventDaily.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AD      30/08/00    Initial version
//	LD		03/05/01	SYS2296 - Upgrade to version 3 of the xml parser
//  RF      04/02/04	Use MSXML4 (included as part of BMIDS727)
///////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////
//
//	AD 15/08/2000
//	Each item is an event which will take place on a time / date basis
//	and will excute it's payload.
//
/////////////////////////////////////////////////////////////////////

//#import "msxml3.dll"  rename_namespace("XmlNS")
#import "msxml4.dll" rename_namespace("XmlNS")

#if !defined(AFX_SCHEDULEEVENTDAILY_H__D80652A4_728D_11D4_825D_000102125FBA__INCLUDED_)
#define AFX_SCHEDULEEVENTDAILY_H__D80652A4_728D_11D4_825D_000102125FBA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Stdafx.h"
#include <list>
#include "CMQEvent.h"

using namespace std; 

typedef list<CMQEvent *> EventPtrList;

class CScheduleEventDaily : public CObject  
{
public:
//	AD 07/09/00 replace the payload by what's in the list of CMQEvents		
//    typedef void* funcparam; 
//    typedef int (*funcptr)(funcparam);  

	
    CScheduleEventDaily()
    {
		::ZeroMemory(this, sizeof(CScheduleEventDaily));
	}

	// AD 07/09/00 change constructor to take in new information
	// CMQEvent, hour, minute and day
    CScheduleEventDaily(int nHour, 
					   int nMinute, enum DayOfWeek eDay);

    CScheduleEventDaily(const CScheduleEventDaily* pScheduleEventDaily)
    {
		::CopyMemory(this, pScheduleEventDaily, sizeof(CScheduleEventDaily));
	}
	virtual ~CScheduleEventDaily();

	bool OnEventSchedule(); // replaces the Execute function
//    void Execute() 
//	{
//		if (m_pFunction) 
//		{
//			(*m_pFunction)(m_pFunctionParameter);
//		} 
//	}

    void SetNext(CScheduleEventDaily* pScheduleEventDailyNext) 
	{
		m_pScheduleEventDailyNext = pScheduleEventDailyNext;
	}

    CScheduleEventDaily* GetNext() const 
	{
		return m_pScheduleEventDailyNext;
	}


    void SetPrev(CScheduleEventDaily* pScheduleEventDailyPrev) 
	{
		m_pScheduleEventDailyPrev = pScheduleEventDailyPrev;
	}

    CScheduleEventDaily* GetPrev() const 
	{
		return m_pScheduleEventDailyPrev;
	}


	void SetDayOfWeek (enum DayOfWeek eDayOfWeek)
	{
		m_eDayOfWeek = eDayOfWeek;
	}

	void SetTime (int nHour, int nMins)
	{
		m_nHour = nHour;
		m_nMinute = nMins;
	}

	enum DayOfWeek GetDayOfWeek(void)
	{
		return m_eDayOfWeek;
	}
	
	int Compare (CScheduleEventDaily*);
	
	int GetHour(void)
	{
		return m_nHour;
	}

	int GetMinutes(void)
	{
		return m_nMinute;
	}

	DWORD GetdwNextWaitIntervalms(void);


	bool InsertNewEvent(CMQEvent* pNewMQEvent);
	bool RemoveQueueEvents(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr);
	void GetEventsInfo(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, XmlNS::IXMLDOMDocument *ptrDomOut);
	void RemoveEvent(GUID gKey, XmlNS::IXMLDOMDocument *ptrDomOut);
	void CreateErrorInfo(XmlNS::IXMLDOMDocument *ptrDomOut, BSTR strType, BSTR strNumber, BSTR strSource, BSTR strDescription,_com_error& comerr);
	void LogEventError(_com_error& comerr, LPCTSTR pszFormat, ...);
private:

//    funcptr m_pFunction;
//    funcparam m_pFunctionParameter;

    CScheduleEventDaily* m_pScheduleEventDailyNext;
    CScheduleEventDaily* m_pScheduleEventDailyPrev;

	DayOfWeek m_eDayOfWeek;
	int m_nHour;
	int m_nMinute;


	// AD 07/09/00
	// member list of CMQEvents - PTR array
	
	EventPtrList EventsList;
};

#endif // !defined(AFX_SCHEDULEEVENTDAILY_H__D80652A4_728D_11D4_825D_000102125FBA__INCLUDED_)
