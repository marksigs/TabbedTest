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
//	LD		15/05/02	SYS4618 - Make logging more robust
///////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////
//
//	AD 15/08/2000
//	Each item is an event which will take place on a time / date basis
//	and will excute it's payload.
//
/////////////////////////////////////////////////////////////////////
#include "afx.h"
#include "stdafx.h"
#include "ScheduleEventDaily.h"
#include "limits.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

// AD 07/09/00
// add in new constructor that takes in the new CMQEvent data and add's 
// stuff into the appropriate list etc


CScheduleEventDaily::CScheduleEventDaily(int nHour, 
					   int nMinute, enum DayOfWeek eDay) 
	{

		m_pScheduleEventDailyNext = NULL; 
		m_nHour = nHour;
		m_nMinute = nMinute;
		m_eDayOfWeek = eDay;
		
	}



CScheduleEventDaily::~CScheduleEventDaily()
{


	unsigned int nSize;
	EventPtrList::iterator i;

	nSize = EventsList.size();

    for (i = EventsList.begin(); i != EventsList.end(); ++i)
	{
		
		delete (*i);
	}
	

	
	
	// remove the MQEvents that are associated with this time slot
	EventsList.clear();

	nSize = EventsList.size();


}

int CScheduleEventDaily::Compare (CScheduleEventDaily* QEvent)
{

	// comprison. -1 = QItem less than this , 0 == QItem same as this, 1 = QItem greater than this
	// take in our two dates and see how they compare day of week / 
	// time wise.

	// day of week 

	if(m_eDayOfWeek < QEvent->GetDayOfWeek())
		return 1;
	else if (m_eDayOfWeek > QEvent->GetDayOfWeek())
		return (-1);

	// made it to this point so day must be the same

	if (m_nHour < QEvent->GetHour())
		return 1;
	else if (m_nHour > QEvent->GetHour())
		return (-1);

	// hour the same as well - final check on minutes

	if (m_nMinute < QEvent->GetMinutes())
		return 1;
	else if (m_nMinute > QEvent->GetMinutes())
		return (-1);

	// the 2 items are the same


	
	return 0;
}


#define INT64_FROM_FILETIME(li) (*((__int64 *)&(li)))
#define FILETIME_FROM_INT64(li) (*((FILETIME *)&(li)))   

bool HrFileTimeToSeconds(FILETIME *pft,                   // file time to convert
						 DWORD *pdwSec)                  // receives time in seconds (0 if error) 
{
    __int64 lift    = 0;
    __int64 liSec   = 0;
	 

	 *pdwSec = 0;
	 lift = INT64_FROM_FILETIME(*pft);
	 liSec = lift / (DWORD) 10000000L;
	 *pdwSec = (DWORD)liSec;
	 return true;
} 


// take in 2 system times - convert them to filetimes then get the
// number of milliseconds seperating the 2 times.
DWORD SystemTimeDifference (SYSTEMTIME *pst1, SYSTEMTIME *pst2)    
{
    LONGLONG       li1, li2 ;
    DWORD          dwRet;

	DWORD dwSec1, dwSec2;

	li1 = (LONGLONG) 0 ;
    li2 = (LONGLONG) 0 ;
	SystemTimeToFileTime (pst1, (FILETIME *) &li1) ;
    SystemTimeToFileTime (pst2, (FILETIME *) &li2) ;

	
	
	HrFileTimeToSeconds((FILETIME*)&li1, &dwSec1);

	HrFileTimeToSeconds((FILETIME*)&li2, &dwSec2);

	dwRet = dwSec2 - dwSec1;

	dwRet = dwRet * 1000;

	return dwRet;
}    

DWORD CScheduleEventDaily::GetdwNextWaitIntervalms(void)
{
	DWORD dwMillisecs = 0;
	WORD nDay = 0;
	int nDayAdjust = 0;
	VARIANT pVar;
	
	// take the current time and pass back the the number of milliseconds 
	// between it and the queue items criteria

	// based on day, hours and minutes
	// did use COLEDateTime but cos we're in ATL we can't so have to work around.

	SYSTEMTIME stCurrentTime, stAdjustedTime;


	
	GetLocalTime(&stCurrentTime);


	VariantInit(&pVar);

    SystemTimeToVariantTime(&stCurrentTime, &pVar.date);
	pVar.vt = VT_DATE;               
	

	// use day of week to work out target time

	nDay = stCurrentTime.wDayOfWeek;

	if ((enum DayOfWeek)nDay != m_eDayOfWeek)
	{
		nDayAdjust = m_eDayOfWeek - nDay;
		if (nDayAdjust < 0)
			nDayAdjust += 7;
	}


	pVar.date += nDayAdjust; 


	VariantTimeToSystemTime(pVar.date, &stAdjustedTime);


	VariantClear(&pVar);
	
	stAdjustedTime.wHour = m_nHour;
	stAdjustedTime.wMinute = m_nMinute;
	stAdjustedTime.wSecond = 0;
	stAdjustedTime.wMilliseconds = 0;


	dwMillisecs = SystemTimeDifference(&stCurrentTime, &stAdjustedTime); 

	return dwMillisecs;
}

// OnScheduleEvent instead of just executing the payload that used to be a part of this 
// object, the object now contains a list of events which we have to cycle thru
bool CScheduleEventDaily::OnEventSchedule() // replaces the Execute function
{
	bool bRet = true;
	EventPtrList::iterator i;


    for (i = EventsList.begin(); i != EventsList.end(); ++i)
	{
		
		(*i)->OnEventSchedule();
	}
	

	return bRet;
}


bool CScheduleEventDaily::InsertNewEvent(CMQEvent* pNewMQEvent)
{
	bool bRet = TRUE;

	EventsList.insert(EventsList.end(), pNewMQEvent);	
	return bRet;
}

// take the pointer to the queue and iterate thru all the MQevents - getting
// rid off all the ones that correpond to 
bool CScheduleEventDaily::RemoveQueueEvents(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr)
{
	bool bRet = TRUE;
	EventPtrList::iterator i;

    for (i = EventsList.begin(); i != EventsList.end(); ++i)
	{

		if ((*i)->CompareQueue(ThreadPoolManagerCommon1Ptr))
		{
			delete (*i);
		}
	}



	return bRet;
}

void CScheduleEventDaily::GetEventsInfo(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, XmlNS::IXMLDOMDocument *ptrDomOut)
{
	// return the info for any events that belong to the requested queue


	EventPtrList::iterator i;
	WCHAR  wszGUID[128];
	
    for (i = EventsList.begin(); i != EventsList.end(); ++i)
	{

		if ((*i)->CompareQueue(ThreadPoolManagerCommon1Ptr))
		{
			// query i for info
			GUID gID;
			gID = (*i)->GetUniqueID();

			StringFromGUID2(gID,wszGUID , 128); 

			_bstr_t bstrOut(wszGUID);

			int iThreads = (*i)->GetThreads();

			// create dom elements if none exist

			XmlNS::IXMLDOMElementPtr	ptrReturnElement;
			XmlNS::IXMLDOMNodePtr		ptrNode;
			bstr_t bstrElement3("THREADSLIST");
			XmlNS::IXMLDOMNodeListPtr	ptrElementList;
			ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
			long lOutList = ptrElementList->Getlength();


			if (lOutList == 0)
			{
				ptrReturnElement = ptrDomOut->createElement("THREADSLIST");
		
				ptrNode = ptrDomOut->appendChild(ptrReturnElement);
			}
			else
			{
				ptrReturnElement = ptrDomOut->GetdocumentElement();
			}


			XmlNS::IXMLDOMElementPtr	ptrSubElement;
			XmlNS::IXMLDOMElementPtr	ptrSubSubElement;

			ptrSubElement = ptrDomOut->createElement("THREADS");
	
			ptrNode = ptrReturnElement->appendChild(ptrSubElement);

			ptrSubSubElement = ptrDomOut->createElement("NUMBER");

			char buffer[20];
			itoa(iThreads, buffer, 10);
	   		bstr_t bstrNumber(buffer);

			ptrSubSubElement->text = bstrNumber;
			ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

			ptrSubSubElement = ptrDomOut->createElement("KEY");
			ptrSubSubElement->text = bstrOut;
			ptrNode = ptrSubElement->appendChild(ptrSubSubElement);


			ptrSubSubElement = ptrDomOut->createElement("DAY");
			switch(m_eDayOfWeek)
			{
				case eSunday:
					ptrSubSubElement->text = L"SUNDAY";
					break;
				case eMonday:
					ptrSubSubElement->text = L"MONDAY";
					break;
				case eTuesday:
					ptrSubSubElement->text = L"TUESDAY";
					break;
				case eWednesday:
					ptrSubSubElement->text = L"WEDNESDAY";
					break;
				case eThursday:
					ptrSubSubElement->text = L"THURSDAY";
					break;
				case eFriday:
					ptrSubSubElement->text = L"FRIDAY";
					break;
				case eSaturday:
					ptrSubSubElement->text = L"SATURDAY";
					break;
				default:
					break;
			}

			
			
			ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

			ptrSubSubElement = ptrDomOut->createElement("TIME");

			sprintf(buffer, "%0.2d:%0.2d", m_nHour, m_nMinute);
	   		bstr_t bstrTime(buffer);

			ptrSubSubElement->text = bstrTime;
			ptrNode = ptrSubElement->appendChild(ptrSubSubElement);
		}
	}


}

// remove an event from a particular time slot
void CScheduleEventDaily::RemoveEvent(GUID gKey, XmlNS::IXMLDOMDocument *ptrDomOut)
{


	EventPtrList::iterator i;
	
	try
	{
		i = EventsList.begin();
		for (int iLoop = 0; iLoop < EventsList.size(); iLoop++)	
		{

			if(EventsList.size() != 0)
			{
				if (IsEqualGUID((*i)->GetUniqueID(), gKey))
				{
					// remove event from list
	
					EventsList.erase(i);
				}
			}
	
			i++;
		}
	}
	catch(_com_error comerr)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Update ", L"Undefined error ",comerr);
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Update ", L"Undefined error ",(_com_error)NULL);
	}

}

void CScheduleEventDaily::CreateErrorInfo(XmlNS::IXMLDOMDocument *ptrDomOut, BSTR strType, BSTR strNumber, BSTR strSource, BSTR strDescription,_com_error& comerr)
{
	XmlNS::IXMLDOMElementPtr	ptrReturnElement;
	XmlNS::IXMLDOMNodePtr		ptrNode;


	// create elements in the DOM

	//<RESPONSE TYPE="SYSERR|APPERR">
	//	<ERROR>
	//		<NUMBER></NUMBER>
	//		<SOURCE></SOURCE>
	//		<DESCRIPTION ></DESCRIPTION>
	//		<VERSION></VERSION>
	//	</ERROR>
	//</ RESPONSE>


	// check if dom already has root element - if it has attach new children
	// otherwise create it.


	bstr_t bstrElement3("ERROR");
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();


	if (lOutList == 0)
	{
		// new error
		ptrReturnElement = ptrDomOut->createElement("RESPONSE");
		ptrReturnElement->setAttribute("TYPE",strType);
		ptrNode = ptrDomOut->appendChild(ptrReturnElement);
	}
	else
	{
		ptrReturnElement = ptrDomOut->GetdocumentElement();

	}

	if (wcsicmp (strType, L"SUCCESS") != 0)
	{
		// it's an error
		
		XmlNS::IXMLDOMElementPtr	ptrSubElement;
		XmlNS::IXMLDOMElementPtr	ptrSubSubElement;

		ptrSubElement = ptrDomOut->createElement("ERROR");
	
		ptrNode = ptrReturnElement->appendChild(ptrSubElement);

		ptrSubSubElement = ptrDomOut->createElement("NUMBER");
		ptrSubSubElement->text = strNumber;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

		ptrSubSubElement = ptrDomOut->createElement("SOURCE");
		ptrSubSubElement->text = strSource;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

		ptrSubSubElement = ptrDomOut->createElement("DESCRIPTION");
		ptrSubSubElement->text = strDescription;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);
		

		// call log event error to put the information in to the event log
		
		char *psz, *psz2;
		int i, i2;
		
		i =  WideCharToMultiByte(CP_ACP, 0, (LPWSTR)strDescription, -1, NULL, 0, NULL, NULL);

#ifdef _UNICODE	
		psz = (LPSTR)LocalAlloc(LMEM_FIXED, i*sizeof(TCHAR)); 
#else
		psz = (LPTSTR)LocalAlloc(LMEM_FIXED, i*sizeof(TCHAR)); 
#endif
		
		WideCharToMultiByte(CP_ACP, 0, (LPWSTR)strDescription, -1, psz, i, NULL, NULL); 

		i2 =  WideCharToMultiByte(CP_ACP, 0, (LPWSTR)strSource, -1, NULL, 0, NULL, NULL);

#ifdef _UNICODE	
		psz2 = (LPSTR)LocalAlloc(LMEM_FIXED, i2*sizeof(TCHAR)); 
#else
		psz2 = (LPTSTR)LocalAlloc(LMEM_FIXED, i2*sizeof(TCHAR)); 
#endif

		
		WideCharToMultiByte(CP_ACP, 0, (LPWSTR)strSource, -1, psz2, i2, NULL, NULL); 

		// add the 2 strings together


#ifdef _UNICODE	
		LogEventError(comerr, LPCTSTR(strcat(psz2, psz)));
#else
		LogEventError(comerr, strcat(psz2, psz));
#endif

	
	
	}

}


void CScheduleEventDaily::LogEventError(_com_error& comerr, LPCTSTR pszFormat, ...)
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
		_Module.LogEventError(_T(" %s (HResult = %d')"), chMsg, comerr.ErrorMessage());
	}
	else
	{
		// output description
		_Module.LogEventError(_T(" %s (HResult = %d')"), chMsg, comerr.Description());
	}
}
