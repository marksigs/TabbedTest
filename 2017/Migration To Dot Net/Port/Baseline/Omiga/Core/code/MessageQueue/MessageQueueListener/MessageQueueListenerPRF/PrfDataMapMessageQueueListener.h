///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfDataMapMessageQueueListener.h
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

#include "PrfData.h"

///////////////////////////////////////////////////////////////////////////////

PRFDATA_DEFINE_OBJECT(PRFOBJ_QUEUEINFO,	   							    100);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_SUCCESS,							    101);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_SUCCESSPERSEC,							102);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_RETRY_NOW,								103);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_RETRY_LATER,						    104);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_RETRY_MOVE_MESSAGE,					105);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_STALL_COMPONENT,					    106);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_STALL_QUEUE,						    107);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_STOLEN,								108);
PRFDATA_DEFINE_COUNTER(QUEUEINFO_ERROR,									109);

///////////////////////////////////////////////////////////////////////////////

#define MAX_PERF_NO_INSTANCES											16
#define MAXLEN_INSTANCE_NAME											64
#define MAX_PERF_NO_THREAD_INSTANCES									64
#define MAXLEN_THREAD_INSTANCE_NAME										64

///////////////////////////////// End Of File /////////////////////////////////
