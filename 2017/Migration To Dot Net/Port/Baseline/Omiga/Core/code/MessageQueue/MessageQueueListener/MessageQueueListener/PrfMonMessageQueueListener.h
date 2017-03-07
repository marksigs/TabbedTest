///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfMonMessageQueueListener.h
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

#if !defined(AFX_PRFMONMESSAGEQUEUELISTENER_H__1EE39A3E_16A5_4A16_9DE5_20297B928F07__INCLUDED_)
#define AFX_PRFMONMESSAGEQUEUELISTENER_H__1EE39A3E_16A5_4A16_9DE5_20297B928F07__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>
#include "PrfMonObject.h"

///////////////////////////////////////////////////////////////////////////////

class CPrfInstanceQueueInfo;

typedef std::map<DWORD, CPrfData::INSTID>		INSTIDS;
typedef std::map<_bstr_t, CPrfInstanceQueueInfo*>	MapPrfInstanceQueue;

///////////////////////////////////////////////////////////////////////////////

class CPrfInstanceQueueInfo : public CObject
{
public:
	CPrfInstanceQueueInfo(CPrfMonObject& rPrfMonObject, _bstr_t bstrQueueName);
	virtual ~CPrfInstanceQueueInfo();
	
public:
	virtual BOOL AddInstances();
	virtual void ResetCounters();
	virtual void RemoveInstances();
	virtual void ResetIdleCounters();

public:
	void IncrementSuccess() 
	{
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_SUCCESS, m_instidQueueInfo); 
			m_rPrfMonObject.IncCtr32(QUEUEINFO_SUCCESSPERSEC, m_instidQueueInfo); 
		}
		catch(...)
		{
		}
	}
	void IncrementRetryNow()
	{
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_RETRY_NOW, m_instidQueueInfo);
		}
		catch(...)
		{
		}
	}
	void IncrementRetryLater()
	{
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_RETRY_LATER, m_instidQueueInfo); 
		}
		catch(...)
		{
		}
	}
	void IncrementRetryMoveMessage() 
	{
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_RETRY_MOVE_MESSAGE, m_instidQueueInfo); 
		}
		catch(...)
		{
		}
	}
	void IncrementStalledComponents() 
	{ 
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_STALL_COMPONENT, m_instidQueueInfo);
		}
		catch(...)
		{
		}
	}
	void SetnDifferentStalledComponents(int nDifferentStalledComponents) 
	{ 
		try
		{
			m_rPrfMonObject.GetCtr32(QUEUEINFO_STALL_COMPONENT, m_instidQueueInfo) = nDifferentStalledComponents; 
		}
		catch(...)
		{
		}
	}
	void SetbStallQueue(BOOL bStalled) 
	{ 
		try
		{
			m_rPrfMonObject.GetCtr32(QUEUEINFO_STALL_QUEUE, m_instidQueueInfo) = bStalled;
		}
		catch(...)
		{
		}
	}
	void IncrementStolen() 
	{ 
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_STOLEN, m_instidQueueInfo);
		}
		catch(...)
		{
		}
	}
	void IncrementError() 
	{
		try
		{
			m_rPrfMonObject.IncCtr32(QUEUEINFO_ERROR, m_instidQueueInfo); 
		}
		catch(...)
		{
		}
	}
public:
	inline CPrfData::INSTID GetInstanceQueueInfo() const { return m_instidQueueInfo; }

private:	
	CPrfData::INSTID m_instidQueueInfo;
	CPrfMonObject& m_rPrfMonObject;
	_bstr_t m_bstrQueueName;
};

///////////////////////////////////////////////////////////////////////////////

class CPrfMonMessageQueueListener : public CPrfMonObject
{
public:
	CPrfMonMessageQueueListener();
	virtual ~CPrfMonMessageQueueListener();

	virtual BOOL AddInstances(const char* pszInstanceName);
	virtual void ResetCounters();
	virtual void RemoveInstances();
	virtual void ResetIdleCounters();

	// support for many queues
	CPrfInstanceQueueInfo* AddInstanceQueueInfo(_bstr_t bstrQueueName);
	CPrfInstanceQueueInfo* CPrfMonMessageQueueListener::RemoveInstanceQueueInfo(_bstr_t bstrQueueName);

private:
	MapPrfInstanceQueue m_mapPrfInstanceQueue;
};

///////////////////////////////////////////////////////////////////////////////

extern CPrfMonMessageQueueListener* g_pPrfMonMessageQueueListener;

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_PRFMONMESSAGEQUEUELISTENER_H__1EE39A3E_16A5_4A16_9DE5_20297B928F07__INCLUDED_)
