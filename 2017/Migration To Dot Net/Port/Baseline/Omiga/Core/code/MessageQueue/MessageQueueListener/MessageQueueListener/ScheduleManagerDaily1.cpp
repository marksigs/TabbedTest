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
//	This class is derived from CScheduleManager classs. Creates
//	and maintains a circular list of CScheduleEventDaily items. Event 
//	is executed when time and date criteria are met.
//
/////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ScheduleManagerDaily1.h"
#include "mutex.h"

using namespace namespaceMutex;

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogInOut.AnyThreadInitialize

//////////////////////////////////////////////////////////////////////

CScheduleManagerDaily::CScheduleManagerDaily()
{
	m_pActiveEvent = m_pPrevEventActioned = NULL;

}

CScheduleManagerDaily::~CScheduleManagerDaily()
{
	RemoveAll(); // tidy up after ourselves
}

bool CScheduleManagerDaily::StartUp()
{
	ANYTHREADLOG_THIS_INOUT(_T("CScheduleManagerDaily::StartUp\n"));
	return CScheduleManager::StartUp();
}

void CScheduleManagerDaily::CloseDown()
{
	ANYTHREADLOG_THIS_INOUT(_T("CScheduleManagerDaily::CloseDown\n"));

	CScheduleManager::CloseDown();
	RemoveAll(); // tidy up after ourselves
}

void CScheduleManagerDaily::OnEventSchedule()
{
	// ok - timeout has expired so it's obviously this events time to
	// execute it's payload.


	// AD read lock needed here.

	// error checking ??
	
	m_pActiveEvent->OnEventSchedule();
	// increment the active event

	m_pActiveEvent = m_pActiveEvent->GetNext();


}

DWORD CScheduleManagerDaily::GetdwNextWaitIntervalms()
{
	DWORD dwRetVal = 0;
	// get the wait time for the next in the queue

	
	// AD read lock needed here
	
	if(m_pActiveEvent != NULL)
	{
		dwRetVal = m_pActiveEvent->GetdwNextWaitIntervalms();
	}
	else
		dwRetVal = ULONG_MAX;

	return dwRetVal;
}



bool CScheduleManagerDaily::AddEventToList(CMQEvent* pNewMQEvent,int nHour, 
					   int nMinute, enum DayOfWeek eDay)
{
	//  items are added by a 3rd party into a circular list. The head is the
	// active event. Only next ptrs are kept as we don't travel backwards 
	// thru the list.


    // AD write lock needed here / try catch processing
	
	CSingleLock lck(&m_csQueue, TRUE);
	CScheduleEventDaily* pNewEvent; // create the new event here rather than passing it in

	// sort out adding the time slots
    
	if (m_pActiveEvent == NULL )
	{
		// the queue is empty so add in the initial item anyway
		ExternalDisableEvents(true);


		// add in the CMQEvent object

		pNewEvent = new CScheduleEventDaily(nHour, nMinute, eDay/*, pNewMQEvent */);
		pNewEvent->InsertNewEvent(pNewMQEvent); 
    
		// add the CMQEvent into the list of

		m_pActiveEvent = pNewEvent;
        pNewEvent->SetNext(m_pActiveEvent);
	    pNewEvent->SetPrev(m_pActiveEvent);

		ExternalEnableEvents();

	}
	else if (m_pActiveEvent == m_pActiveEvent->GetNext())
	{
			// only one element - do a compare to see where to slot it in (before or after)
	


		// check to see if the element all ready exists - if it does then just add in 
		// CMQEvent

		
		pNewEvent = new CScheduleEventDaily(nHour, nMinute, eDay/*, pNewMQEvent*/);

		// need to put in exeption handler stuff		
		
		if (pNewEvent->Compare(m_pActiveEvent) != 0)
		{
		
			pNewEvent->InsertNewEvent(pNewMQEvent); 
			pNewEvent->SetNext(m_pActiveEvent); // wrap around
			pNewEvent->SetPrev(m_pActiveEvent); 
			m_pActiveEvent->SetNext(pNewEvent);
			m_pActiveEvent->SetPrev(pNewEvent);


			if (pNewEvent->Compare(m_pActiveEvent) != 1)
			{
				// after - no probs here me laddo
			}
			else
			{
				// and this is the code before
			
				// if the item comes before the active one check the execute
				// time cos it may be less and have to be scheduled 1st

				if(pNewEvent->GetdwNextWaitIntervalms() < m_pActiveEvent->GetdwNextWaitIntervalms())
				{
					// suspend queue 
					// make new item the active one
					// restart the queue

					ExternalDisableEvents(true);

					m_pActiveEvent = pNewEvent;

					ExternalEnableEvents();
			
			
				}

			}	
		}
		else
		{ 
			// add in new CMQEvent to the existing 
			m_pActiveEvent->InsertNewEvent(pNewMQEvent); 
			delete pNewEvent;


		}

	}
	else
	{
		// there is more than 1 item in the queue so run through the queue to 
		// find where slot it in
		
		// start at the head 
		bool bFoundSlot = false;

	    CScheduleEventDaily* pRetrievedEvent = NULL;
	    CScheduleEventDaily* pOriginalActiveEvent = NULL;

		pOriginalActiveEvent = pRetrievedEvent = m_pActiveEvent;


		
		// allocate new event
		// do compare - if that time slot already exists then just add the MQEvent
	
		// AD - check this bit

		pNewEvent = new CScheduleEventDaily(nHour, nMinute, eDay/*, pNewMQEvent*/);


		while(!bFoundSlot)
		{

			// make a comparison

			if(pRetrievedEvent->Compare(pNewEvent) == 0)
			{

				// item is the same so slot it in the after the current record  

				// just add the MQEvent

				// add in new CMQEvent to the existing 
				delete pNewEvent;
				m_pActiveEvent->InsertNewEvent(pNewMQEvent); 



				bFoundSlot = true;
			}
			else if (pRetrievedEvent->Compare(pNewEvent) == (-1))
			{
				// item is less - slot in before

				// check to see if the one before prev is less than as well
				// if it is then we have to go further back

				if(pRetrievedEvent->GetPrev()->Compare(pNewEvent) == (-1))
				{
					// go back even further
				}
				else
				{
						
					pNewEvent->SetNext(pRetrievedEvent);
			        pNewEvent->SetPrev(pRetrievedEvent->GetPrev());

					pRetrievedEvent->GetPrev()->SetNext(pNewEvent);
					pRetrievedEvent->SetPrev(pNewEvent);
					

					if(pNewEvent->GetdwNextWaitIntervalms() < m_pActiveEvent->GetdwNextWaitIntervalms())
					{
						// suspend queue 
						// make new item the active one
						// restart the queue

						ExternalDisableEvents(true);

						m_pActiveEvent = pNewEvent;

						ExternalEnableEvents();
			
					}
					
					
					// if time is less than current active then reset the active event

					bFoundSlot = true;
				}
				
				
				
			}
			else
			{
				// item is greater - find out how the item following compares to the
				// next item in the queue

				if ((pRetrievedEvent->GetNext())->Compare(pNewEvent) == (-1))
				{
					// item is less than the next one in the queue so slot that in
					pNewEvent->SetNext(pRetrievedEvent->GetNext());
			        pNewEvent->SetPrev(pRetrievedEvent);
					pRetrievedEvent->GetNext()->SetPrev(pNewEvent);
					pRetrievedEvent->SetNext(pNewEvent);

					bFoundSlot = true;
				}

			}
							
			// cycle thru the events				
	        pRetrievedEvent = pRetrievedEvent->GetNext();
			if (pRetrievedEvent == pOriginalActiveEvent)
			{
				// safety valve - if get round to the head then we've gone full circle
				// event must be greater than any in the list then add it in


				pNewEvent->SetNext(pRetrievedEvent->GetNext());
		        pNewEvent->SetPrev(pRetrievedEvent);
				pRetrievedEvent->GetNext()->SetPrev(pNewEvent);
				pRetrievedEvent->SetNext(pNewEvent);


				bFoundSlot = true;
			}
		} // while slot not found
	}
	


	return true;
}




// a queue pointer is passed in. We then iterate thru all the event list entries.
// Then qo through each entry's MQEvent list.
bool CScheduleManagerDaily::RemoveQueueEvents(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr)
{
	bool bRet = false;
	bool bEmpty = false;

	CScheduleEventDaily* pRetrievedEvent = NULL;


	if (m_pActiveEvent != NULL && m_pActiveEvent->GetPrev() != NULL)
	{
	
		
		while(!bEmpty)  
		{

			pRetrievedEvent = m_pActiveEvent;

		
			pRetrievedEvent->RemoveQueueEvents(ThreadPoolManagerCommon1Ptr);
		
			if(m_pActiveEvent->GetNext() == m_pActiveEvent)
			{
				// end of the list
				bEmpty = true;
			}
			else
			{
				m_pActiveEvent = m_pActiveEvent->GetNext();
			}

		}
	}



	return bRet;
}

// iterate thru the event lists passing back information - via dom??
void CScheduleManagerDaily::GetEventInfo(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, XmlNS::IXMLDOMDocument *ptrDomOut)
{

	bool bEmpty = false;

	CScheduleEventDaily* pRetrievedEvent = NULL;

	
	if (m_pActiveEvent != NULL && m_pActiveEvent->GetPrev() != NULL)
	{
	
		
		pRetrievedEvent = m_pActiveEvent;

		while(!bEmpty)  
		{


			pRetrievedEvent->GetEventsInfo(ThreadPoolManagerCommon1Ptr, ptrDomOut);
		
			if(pRetrievedEvent->GetNext() == m_pActiveEvent)
			{
				// end of the list
				bEmpty = true;
			}
			else
			{
				pRetrievedEvent = pRetrievedEvent->GetNext();
			}

		}
	}


}


bool CScheduleManagerDaily::RemoveAll(void)
{
	// remove all the items in the list

	// Start at the active event and work through the list

	// AD would need write lock here
	
	bool bEmpty = false;
    
	CScheduleEventDaily* pRetrievedEvent = NULL;

	// break the circle
	// AD 20/09/00 but make sure that we have an item otherwise it'll go pop

	if (m_pActiveEvent != NULL && m_pActiveEvent->GetPrev() != NULL)
	{
	
		m_pActiveEvent->GetPrev()->SetNext(NULL);
		
		while(!bEmpty)  
		{

			pRetrievedEvent = m_pActiveEvent;

		
		
			if(m_pActiveEvent->GetNext() == NULL)
			{
				// end of the list
				bEmpty = true;
			}
			else
			{
				m_pActiveEvent = m_pActiveEvent->GetNext();
			}

			delete pRetrievedEvent;
		}
	}

	return true;
}

bool CScheduleManagerDaily::RemoveEvent(XmlNS::IXMLDOMDocument *ptrDomOut, GUID gKey)
{

	// pass in the queue and compare the guid with the events in the list

	bool bEmpty = false;

	CScheduleEventDaily* pRetrievedEvent = NULL;

	
	if (m_pActiveEvent != NULL && m_pActiveEvent->GetPrev() != NULL)
	{
	
		
		pRetrievedEvent = m_pActiveEvent;

		while(!bEmpty)  
		{


			pRetrievedEvent->RemoveEvent(gKey, ptrDomOut);
		
			if(pRetrievedEvent->GetNext() == m_pActiveEvent)
			{
				// end of the list
				bEmpty = true;
			}
			else
			{
				pRetrievedEvent = pRetrievedEvent->GetNext();
			}

		}
	}

	return true;
}
