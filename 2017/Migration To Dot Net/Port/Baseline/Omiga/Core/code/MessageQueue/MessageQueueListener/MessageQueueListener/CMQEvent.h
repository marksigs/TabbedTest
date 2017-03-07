///////////////////////////////////////////////////////////////////////////////
//	FILE:			CMQEvent.H
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

///////////////////////////////////////////////////////////////////////////////
//
//	Instead on the CScheduleEvent... objects carrying all the time info and
//	payload for an event - the CScheduleEvent... just has a time and the 
//	CMQEvent (and derivatives) carry the payload as it were.
//	CScheduleEvent... has 1-n CMQEvents.
//
///////////////////////////////////////////////////////////////////////////////

#include "ThreadPoolManagerOMMQ1.h"
#include "ThreadPoolManagerMSMQ1.h"

_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerMSMQ1, __uuidof(IInternalThreadPoolManagerMSMQ1));
_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerOMMQ1, __uuidof(IInternalThreadPoolManagerOMMQ1));
_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerCommon1, __uuidof(IInternalThreadPoolManagerCommon1));


class CMQEvent : public CObject
{
public:
	CMQEvent();
	virtual ~CMQEvent();

	virtual OnEventSchedule() = 0;

	GUID m_UniqueID; // unique identifier

	void SetUniqueID(GUID GuidIn);
	GUID GetUniqueID(void);
	virtual bool CompareQueue(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr) = 0;
	virtual int GetThreads(void) = 0;
};

// derived class that increases a queue's number of active threads
// members are - the pointer to the queue and the no of threads
class CMQEventChangeThreads : public CMQEvent
{
public:
	CMQEventChangeThreads(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr, int nThreads);
	virtual ~CMQEventChangeThreads();
	OnEventSchedule();

	int m_nNumberOfThreads;
	IInternalThreadPoolManagerCommon1Ptr m_ThreadPoolManagerCommon1Ptr;
	
	bool CompareQueue(IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr);
	int GetThreads(void);
};

