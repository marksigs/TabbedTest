///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfDataMapMessageQueueListener.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	LD		06/04/01	SYS2248 - Initial version.  Add Performance counters
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#pragma warning(disable : 4201)	// nonstandard extension used: nameless struct/union
#include <Windows.h>
#pragma warning(default: 4201)	// nonstandard extension used: nameless struct/union
#include "PrfData.h"
#include "PrfDataMapMessageQueueListener.h"

///////////////////////////////////////////////////////////////////////////////

PRFDATA_MAP_BEGIN()


	PRFDATA_MAP_OBJ(
		PRFOBJ_QUEUEINFO, 
		L"MQL: Queue information", 
		L"The history of calls object type includes those counters "
		L"that apply to calls since the queue was created (or the service was started).", 
		PERF_DETAIL_NOVICE, QUEUEINFO_SUCCESS, MAX_PERF_NO_THREAD_INSTANCES, MAXLEN_INSTANCE_NAME)

	PRFDATA_MAP_CTR(
		QUEUEINFO_SUCCESS,  
		L"Success",  
		L"The number of successful calls; "
		L"this is a good indication of how busy this queue is.",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_SUCCESSPERSEC,  
		L"Success/sec",  
		L"The change in the number of successful calls over time.",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_COUNTER)
	PRFDATA_MAP_CTR(
		QUEUEINFO_RETRY_NOW,
		L"Retry Now",  
		L"The number of calls where a component has a requested to be retried now.",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_RETRY_LATER,
		L"Retry Later",  
		L"The number of calls where a component has a requested to be retried later.",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_RETRY_MOVE_MESSAGE,  
		L"Move Message",  
		L"The number of calls where a component has a requested to be moved to another queue.",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_STALL_COMPONENT,  
		L"Stall Component",  
		L"The number of calls where a component has a requested that all similar components are stalled on this queue",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_STALL_QUEUE,  
		L"Stall Queue",  
		L"The number of calls where a component has a requested that all components are stalled on this queue",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_STOLEN,  
		L"Stolen messages",  
		L"The number of messages not found (i.e stolen by another process)",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)
	PRFDATA_MAP_CTR(
		QUEUEINFO_ERROR,  
		L"Errors",  
		L"The number of errors encountered",  
		PERF_DETAIL_NOVICE, 0, PERF_COUNTER_RAWCOUNT)



PRFDATA_MAP_END("MessageQueueListener")



///////////////////////////////// End Of File /////////////////////////////////
