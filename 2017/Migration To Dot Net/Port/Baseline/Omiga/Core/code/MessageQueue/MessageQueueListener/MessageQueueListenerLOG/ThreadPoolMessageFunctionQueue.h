///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolMessageFunctionQueue.h
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

#if !defined(AFX_THREADPOOLMESSAGEFUNCTIONQUEUE_H__A94D0FA8_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
#define AFX_THREADPOOLMESSAGEFUNCTIONQUEUE_H__A94D0FA8_4B60_11D4_8237_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ThreadPoolMessage.h"

///////////////////////////////////////////////////////////////////////////////

class CThreadPoolMessageFunctionQueue : public CThreadPoolMessage
{
public:
	CThreadPoolMessageFunctionQueue() :
	  	m_pNext(NULL), 
		m_pPrev(NULL) 
	{
	}
	CThreadPoolMessageFunctionQueue(funcptr pFunction, funcparam pFunctionParameter) :
        CThreadPoolMessage(pFunction, pFunctionParameter),
	  	m_pNext(NULL), 
		m_pPrev(NULL) 
	{
	}
	virtual ~CThreadPoolMessageFunctionQueue()
	{
	}

public:
    void SetNext(CThreadPoolMessageFunctionQueue* pNext) 
	{
		m_pNext = pNext;
	}
    void SetPrev(CThreadPoolMessageFunctionQueue* pPrev) 
	{
		m_pPrev = pPrev;
	}
    CThreadPoolMessageFunctionQueue* GetNext() const 
	{
		return m_pNext;
	}
    CThreadPoolMessageFunctionQueue* GetPrev() const 
	{
		return m_pPrev;
	}

private:
	CThreadPoolMessageFunctionQueue* m_pNext;
    CThreadPoolMessageFunctionQueue* m_pPrev;


};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_THREADPOOLMESSAGEFUNCTIONQUEUE_H__A94D0FA8_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
