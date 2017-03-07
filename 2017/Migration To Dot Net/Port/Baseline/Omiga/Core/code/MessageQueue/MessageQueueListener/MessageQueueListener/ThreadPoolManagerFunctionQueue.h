///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerFunctionQueue.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_THREADPOOLMANAGERFUNCTIONQUEUE_H__A94D0FA5_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
#define AFX_THREADPOOLMANAGERFUNCTIONQUEUE_H__A94D0FA5_4B60_11D4_8237_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "mutex.h"
#include "ThreadPoolManager.h"
class CThreadPoolMessageFunctionQueue;

///////////////////////////////////////////////////////////////////////////////

class CThreadPoolManagerFunctionQueue : public CThreadPoolManager
{
public:
	CThreadPoolManagerFunctionQueue();
	virtual ~CThreadPoolManagerFunctionQueue();

	bool AddFunctionToQueue(CThreadPoolMessageFunctionQueue* pThreadPoolMessageFunctionQueue); // queue gains ownership of the message

private:
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread); // returns NULL if there is no more work to be done 

    CThreadPoolMessageFunctionQueue* volatile m_pMessageHead;
    CThreadPoolMessageFunctionQueue* volatile m_pMessageTail;

	enum { eSpinCount = 4000 };
    namespaceMutex::CCriticalSection m_csQueue;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_THREADPOOLMANAGERFUNCTIONQUEUE_H__A94D0FA5_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
