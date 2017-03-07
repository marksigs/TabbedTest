///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfMonObject.h
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

#if !defined(AFX_PRFMONOBJECT_H__4A6A69A9_5D9D_4AA2_ADDD_22F86B2968C5__INCLUDED_)
#define AFX_PRFMONOBJECT_H__4A6A69A9_5D9D_4AA2_ADDD_22F86B2968C5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "..\MessageQueueListenerPRF\PrfData.h"
#include "..\MessageQueueListenerPRF\PrfDataMapMessageQueueListener.h"

///////////////////////////////////////////////////////////////////////////////

class CPrfMonObject : public CObject
{
public:
	CPrfMonObject();
	virtual ~CPrfMonObject();
	
public:
	BOOL Install(const char* mbszAppName = NULL);
	BOOL UnInstall();

	BOOL Initialise(BOOL bEnabled, const char* pszInstanceName = NULL, BOOL bResetCountersOnIdle = TRUE);
	BOOL Exit();

public:
	virtual BOOL AddInstances(const char* pszInstanceName) = 0;
	virtual void ResetCounters() = 0;
	virtual void RemoveInstances() = 0;
	virtual void ResetIdleCounters() = 0;

public:
	CPrfData::INSTID AddInstance(
		CPrfData::OBJID ObjId, 
		const char* mbszInstName,
		CPrfData::OBJID ObjIdParent = 0,
		CPrfData::INSTID InstIdParent = 0);

	inline void RemoveInstance(CPrfData::OBJID ObjId, CPrfData::INSTID InstId)
	{
		if (m_bEnabled) g_PrfData.RemoveInstance(ObjId, InstId);
	}

public:
	// Returns 32-bit address of counter in shared data block
	inline LONG& GetCtr32(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{
		::InterlockedExchange((LPLONG)&m_bSet, TRUE);
		return m_bEnabled && m_bInitialised ? g_PrfData.GetCtr32(CtrId, InstId) : m_lDummyCtr; 
	}
	inline void IncCtr32(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ GetCtr32(CtrId, InstId)++; }
	inline void DecCtr32(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ LONG& lCtr = GetCtr32(CtrId, InstId); if (lCtr > 0) lCtr--; }
	inline __int64& GetCtr64(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{
		::InterlockedExchange((LPLONG)&m_bSet, TRUE);
		return m_bEnabled && m_bInitialised ? g_PrfData.GetCtr64(CtrId, InstId) : m_i64DummyCtr; 
	}
	inline void IncCtr64(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ GetCtr64(CtrId, InstId)++; }
	inline void DecCtr64(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ __int64& lCtr = GetCtr64(CtrId, InstId); if (lCtr > 0) lCtr--; }
	inline BOOL GetEnabled() const { return m_bEnabled; }

public:	
	inline void ResetCtr32(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ if (m_bEnabled) g_PrfData.GetCtr32(CtrId, InstId) = 0; }
	inline void ResetCtr64(CPrfData::CTRID CtrId, CPrfData::INSTID InstId = 0) const 
	{ if (m_bEnabled) g_PrfData.GetCtr64(CtrId, InstId) = 0; }
	inline wchar_t* mbstowcs(wchar_t* wcstr, const char* mbstr)
	{
		return (mbstr && ::mbstowcs(wcstr, mbstr, strlen(mbstr) + 1) != -1) ? wcstr : NULL;
	}

protected:
	static void ResetIdleCountersThread(void* pvArg);

private:
	WCHAR						m_wszAppName[_MAX_PATH];
	BOOL						m_bInitialised;
	BOOL						m_bEnabled;
	BOOL						m_bResetCountersOnIdle;
	BOOL						m_bSet;
	LONG						m_lDummyCtr;
	__int64						m_i64DummyCtr;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_PRFMONOBJECT_H__4A6A69A9_5D9D_4AA2_ADDD_22F86B2968C5__INCLUDED_)
