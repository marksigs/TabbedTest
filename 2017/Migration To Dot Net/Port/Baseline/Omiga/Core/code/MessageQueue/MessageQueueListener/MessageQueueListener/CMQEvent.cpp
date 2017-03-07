///////////////////////////////////////////////////////////////////////////////
//	FILE:			CMQEvent.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AD      06/09/00    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CMQEvent.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMQEvent::CMQEvent()
{
	CoCreateGuid (&m_UniqueID);   

}

CMQEvent::~CMQEvent()
{

}

void CMQEvent::SetUniqueID(GUID GuidIn)
{
	m_UniqueID = GuidIn;
}


GUID CMQEvent::GetUniqueID(void)
{
	return m_UniqueID;
}

///////////////////////////////////////////////////////////////////////
// Derived class for increasing a Queue's number of active threads
///////////////////////////////////////////////////////////////////////

CMQEventChangeThreads::CMQEventChangeThreads(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, int nThreads)
{
	// set the thread the action is going to affect
	m_ThreadPoolManagerCommon1Ptr = ThreadPoolManagerCommon1Ptr;

	m_nNumberOfThreads = nThreads;
	// for test try and access queue info
}

CMQEventChangeThreads::~CMQEventChangeThreads()
{
	m_ThreadPoolManagerCommon1Ptr = NULL;

}

CMQEventChangeThreads::OnEventSchedule()
{
	// check hresult here
	if (m_ThreadPoolManagerCommon1Ptr)
	{
		m_ThreadPoolManagerCommon1Ptr->put_NumberOfThreads(m_nNumberOfThreads);
	}	
}

int CMQEventChangeThreads::GetThreads()
{
	return m_nNumberOfThreads;

}


bool CMQEventChangeThreads::CompareQueue(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr)
{
	// check to see if the queue pointer we pass in is the same as the 
	// queue this event is associated with
	bool bRet = false;

	if (m_ThreadPoolManagerCommon1Ptr = ThreadPoolManagerCommon1Ptr)
		bRet = true;

	return bRet;
}
