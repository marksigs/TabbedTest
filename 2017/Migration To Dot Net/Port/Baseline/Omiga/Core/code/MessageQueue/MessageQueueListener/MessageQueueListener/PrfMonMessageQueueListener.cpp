///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfMonMessageQueueListener.cpp
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
#include <process.h>
#include "PrfMonMessageQueueListener.h"

///////////////////////////////////////////////////////////////////////////////

CPrfInstanceQueueInfo::CPrfInstanceQueueInfo(CPrfMonObject& rPrfMonObject, _bstr_t bstrQueueName) :
	m_instidQueueInfo(reinterpret_cast<CPrfData::INSTID>(-1)),
	m_rPrfMonObject(rPrfMonObject),
	m_bstrQueueName(bstrQueueName)
{
}

CPrfInstanceQueueInfo::~CPrfInstanceQueueInfo()
{
}

BOOL CPrfInstanceQueueInfo::AddInstances()
{
	BOOL bSuccess = TRUE;

	m_instidQueueInfo	= m_rPrfMonObject.AddInstance(PRFOBJ_QUEUEINFO, m_bstrQueueName);

	bSuccess = 
		m_instidQueueInfo != reinterpret_cast<CPrfData::INSTID>(-1);

	return bSuccess;
}

void CPrfInstanceQueueInfo::RemoveInstances()
{
	if (m_instidQueueInfo != reinterpret_cast<CPrfData::INSTID>(-1))
	{
		m_rPrfMonObject.RemoveInstance(PRFOBJ_QUEUEINFO, m_instidQueueInfo);
	}
}

void CPrfInstanceQueueInfo::ResetCounters()
{
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_SUCCESS, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_SUCCESSPERSEC, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_RETRY_NOW, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_RETRY_LATER, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_RETRY_MOVE_MESSAGE, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_STALL_COMPONENT, m_instidQueueInfo);
	m_rPrfMonObject.ResetCtr32(QUEUEINFO_STALL_QUEUE, m_instidQueueInfo);
}

void CPrfInstanceQueueInfo::ResetIdleCounters()
{
}


///////////////////////////////////////////////////////////////////////////////

CPrfMonMessageQueueListener::CPrfMonMessageQueueListener()
{
}

CPrfMonMessageQueueListener::~CPrfMonMessageQueueListener()
{
}


BOOL CPrfMonMessageQueueListener::AddInstances(const char* pszInstanceName)
{
	BOOL bSuccess = TRUE;

//	MapPrfInstanceQueue::iterator it;
//	for (it = m_mapPrfInstanceQueue.begin(); it != m_mapPrfInstanceQueue.end() && bSuccess; it++)
//	{
//		bSuccess = (*it).second->AddInstances(pszInstanceName);
//	}

	return bSuccess;
}

void CPrfMonMessageQueueListener::RemoveInstances()
{
//	MapPrfInstanceQueue::iterator it;
//	for (it = m_mapPrfInstanceQueue.begin(); it != m_mapPrfInstanceQueue.end(); it++)
//	{
//		(*it).second->RemoveInstances();
//	}
}

void CPrfMonMessageQueueListener::ResetCounters()
{
	g_PrfData.LockCtrs();

	MapPrfInstanceQueue::iterator it;
	for (it = m_mapPrfInstanceQueue.begin(); it != m_mapPrfInstanceQueue.end(); it++)
	{
		(*it).second->ResetCounters();
	}

	g_PrfData.UnlockCtrs();
}

void CPrfMonMessageQueueListener::ResetIdleCounters()
{
	g_PrfData.LockCtrs();

	MapPrfInstanceQueue::iterator it;
	for (it = m_mapPrfInstanceQueue.begin(); it != m_mapPrfInstanceQueue.end(); it++)
	{
		(*it).second->ResetIdleCounters();
	}
	
	g_PrfData.UnlockCtrs();
}

CPrfInstanceQueueInfo* CPrfMonMessageQueueListener::AddInstanceQueueInfo(_bstr_t bstrQueueName)
{
	CPrfInstanceQueueInfo* pPrfInstanceQueue = new CPrfInstanceQueueInfo(*this, bstrQueueName);
	m_mapPrfInstanceQueue[bstrQueueName] = pPrfInstanceQueue;
	pPrfInstanceQueue->AddInstances();
	return pPrfInstanceQueue;
}

CPrfInstanceQueueInfo* CPrfMonMessageQueueListener::RemoveInstanceQueueInfo(_bstr_t bstrQueueName)
{
	MapPrfInstanceQueue::iterator it = m_mapPrfInstanceQueue.find(bstrQueueName);
	if (it != m_mapPrfInstanceQueue.end())
	{
		(*it).second->RemoveInstances();
		delete (*it).second;
		m_mapPrfInstanceQueue.erase(it);
	}
	return NULL;
}

///////////////////////////////////////////////////////////////////////////////

// Singleton global performance monitor object - created on the heap 
CPrfMonMessageQueueListener* g_pPrfMonMessageQueueListener = NULL;
