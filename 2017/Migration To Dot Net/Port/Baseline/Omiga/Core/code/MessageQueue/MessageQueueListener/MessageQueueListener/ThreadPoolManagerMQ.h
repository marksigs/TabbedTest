///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerMQ.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      06/09/00    Initial version
//	LD		06/04/01	SYS2248 - Add performance counters
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////


#if !defined(AFX_THREADPOOLMANAGERMQ_H__BB973168_83D7_11D4_8250_005004E8D1A7__INCLUDED_)
#define AFX_THREADPOOLMANAGERMQ_H__BB973168_83D7_11D4_8250_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "resource.h"       // main symbols
#include <string>
#include <list>
#include <comutil.h>
#include "MessageQueueListener.h"
#include "ThreadPoolManager.h"
#include "Log.h"

///////////////////////////////////////////////////////////////////////////////

class CThreadPoolManagerMQ : public CThreadPoolManager, public CLog
{
public:
	CThreadPoolManagerMQ();
	virtual ~CThreadPoolManagerMQ();

public:
	STDMETHOD(get_ComponentsStalled)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(put_AddStalledComponents)(/*[in]*/ VARIANT newVal);
	STDMETHOD(put_RestartComponents)(/*[in]*/ VARIANT newVal);

	void AddStallComponent(_bstr_t bstrProgID, BOOL* pbAlreadyStalled = NULL);
	void RemoveStallComponent(_bstr_t bstrProgID);
	bool IsComponentStalled(_bstr_t bstrProgID);
	bool IsAnyComponentStalled();
	int GetnDifferentComponentsStalled();
	
private:
	typedef std::list<_bstr_t> TListStalledComponents;
	typedef TListStalledComponents::iterator TIteratorStalledComponents;

	TListStalledComponents m_listStalledComponents;

    namespaceMutex::CSharedLockSmallOperations m_sharedlockStalledComponents;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_THREADPOOLMANAGERMQ_H__BB973168_83D7_11D4_8250_005004E8D1A7__INCLUDED_)
