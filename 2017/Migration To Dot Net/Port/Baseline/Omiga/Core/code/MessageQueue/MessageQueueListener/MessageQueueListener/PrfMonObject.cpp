///////////////////////////////////////////////////////////////////////////////
//	FILE:			PrfMonObject.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	LD		06/04/01	SYS2248 - Initial version.  Add Performance counters
//  LD		29/05/01	SYS2329 - Perfmon access violation on NT4 
//							Use short filenames due to restriction of Performance
//							monitor.
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <process.h>
#include "PrfMonObject.h"

///////////////////////////////////////////////////////////////////////////////

CPrfMonObject::CPrfMonObject() :
	m_bInitialised(FALSE),
	m_bEnabled(FALSE),
	m_bSet(FALSE),
	m_bResetCountersOnIdle(TRUE)
{
	m_wszAppName[0] = L'\0';
}

CPrfMonObject::~CPrfMonObject()
{
}
	
BOOL CPrfMonObject::Install(const char* mbszAppName)
{
	BOOL bSuccess = TRUE;

	LPCWSTR pwszAppName = mbstowcs(m_wszAppName, mbszAppName);

	WCHAR szLongPath[_MAX_PATH] = L"";
	bSuccess = GetModuleFileNameW(NULL, szLongPath, (sizeof(szLongPath) / sizeof(szLongPath[0]))) != 0;

	if (bSuccess)
	{
#ifdef _DEBUG
		wcscpy(wcsrchr(szLongPath, L'\\') + 1, L"MessageQueueListenerPRF.dll");
#else
		wcscpy(wcsrchr(szLongPath, L'\\') + 1, L"MessageQueueListenerPRF.dll");
#endif
		WCHAR szShortPath[_MAX_PATH] = L"";
		DWORD dw = GetShortPathNameW(szLongPath, szShortPath, _MAX_PATH);
		if (dw > 0 && dw < _MAX_PATH)
		{
			g_PrfData.Install(szShortPath, pwszAppName);
		}
		else
		{
			bSuccess = FALSE;
		}
	}

	return bSuccess;
}

BOOL CPrfMonObject::UnInstall()
{
	BOOL bSuccess = TRUE;

	try
	{
		g_PrfData.Uninstall();
	}
	catch(...)
	{
		// fail safe to catch any exceptions during destruction
		bSuccess = FALSE;
	}

	return bSuccess;
}

BOOL CPrfMonObject::Initialise(BOOL bEnabled, const char* pszInstanceName, BOOL bResetCountersOnIdle)
{
	m_bEnabled = bEnabled;

	if (m_bEnabled)
	{
		DWORD dwErr = g_PrfData.Activate(NULL);
		if (dwErr == ERROR_SUCCESS || dwErr == ERROR_ALREADY_EXISTS)
		{
			m_bInitialised = AddInstances(pszInstanceName);
			if (m_bInitialised)
			{
				ResetCounters();
			}
			if (m_bInitialised && m_bResetCountersOnIdle)
			{
				_beginthread(ResetIdleCountersThread, 0, reinterpret_cast<void*>(this));
			}
		}
		else
		{
			m_bInitialised = m_bEnabled = FALSE;
		}
	}
	else
	{
		// still initialised even if not enabled
		m_bInitialised = TRUE;
	}

	return m_bInitialised;
}

BOOL CPrfMonObject::Exit()
{
	BOOL bSuccess = TRUE;

	try
	{
		if (m_bEnabled && m_bInitialised)
		{
			::InterlockedExchange((LPLONG)&m_bInitialised, FALSE);
			ResetCounters();
			RemoveInstances();
		}
		m_bEnabled = m_bInitialised = FALSE;
	}
	catch(...)
	{
		// fail safe to catch any exceptions during destruction
		bSuccess = FALSE;
	}

	return bSuccess;
}

CPrfData::INSTID CPrfMonObject::AddInstance(
	CPrfData::OBJID ObjId, 
	const char* mbszInstName,
	CPrfData::OBJID ObjIdParent,
	CPrfData::INSTID InstIdParent)
{
	CPrfData::INSTID InstId = reinterpret_cast<CPrfData::INSTID>(-1);

	if (m_bEnabled)
	{
		WCHAR wszInstName[_MAX_PATH] = L"";
		LPCWSTR pwszInstName = mbstowcs(wszInstName, mbszInstName);
		InstId = g_PrfData.AddInstance(ObjId, pwszInstName, ObjIdParent, InstIdParent);
	}

	return InstId;
}

///////////////////////////////////////////////////////////////////////////////
//	ResetIdleCounters()
//	Resets transient counters, e.g., raw counters that are not timed based, 
//	but represent a transient, absolute value.
///////////////////////////////////////////////////////////////////////////////
void CPrfMonObject::ResetIdleCountersThread(void* pvArg)
{
//	TRACE("->CPrfMonObject::ResetIdleCountersThread() (0x%X)\n", ::GetCurrentThreadId());

	::SetThreadPriority(::GetCurrentThread(), THREAD_PRIORITY_IDLE);

	CPrfMonObject* pObjectPRFMon = (CPrfMonObject*)pvArg;

	while (pObjectPRFMon && pObjectPRFMon->m_bInitialised)
	{
		::Sleep(1000);
		if (!pObjectPRFMon->m_bSet)
		{
			// A counter has not been set in the last second, so reset all idle counters.
			pObjectPRFMon->ResetIdleCounters();
		}
		else
		{
			::InterlockedExchange((LPLONG)&pObjectPRFMon->m_bSet, FALSE);
		}
	}

//	TRACE("<-CPrfMonObject::ResetIdleCountersThread() (0x%X)\n", ::GetCurrentThreadId());
}
