///////////////////////////////////////////////////////////////////////////////
//	FILE:			mutex.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version (based on ESERVER code)
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "mutex.h"

namespace namespaceMutex
{
using namespace std;

// Conditional expression is constant.
#pragma warning(disable: 4127)

#ifndef UNUSED
#ifdef _DEBUG
#define UNUSED(x)
#else
#define UNUSED(x) x
#endif
#endif

/////////////////////////////////////////////////////////////////////////////
// Basic synchronization object

CSyncObject::CSyncObject(LPCTSTR pstrName)
{
	UNUSED(pstrName);   // unused in release builds

	m_hObject		= NULL;

	memset(&m_sa, 0, sizeof(m_sa));
	memset(&m_sd, 0, sizeof(m_sd));

#ifdef _DEBUG
	if (pstrName != NULL)
	{
		m_strName = pstrName;
	}
#endif
}

CSyncObject::~CSyncObject()
{
	if (m_hObject != NULL)
	{
		::CloseHandle(m_hObject);
		m_hObject = NULL;
	}
}

BOOL CSyncObject::Lock(DWORD dwTimeout)
{
	if (::WaitForSingleObject(m_hObject, dwTimeout) == WAIT_OBJECT_0)
		return TRUE;
	else
		return FALSE;
}

BOOL CSyncObject::CreateAccessibleSecurityDescriptor()
{
	// Create a null DACL (discretionary access-control list).
	BOOL bSuccess = FALSE;
    
	bSuccess =
		InitializeSecurityDescriptor(&m_sd, SECURITY_DESCRIPTOR_REVISION) &&
		SetSecurityDescriptorDacl(&m_sd, TRUE, (PACL)NULL, FALSE);

	if (bSuccess)
	{
		m_sa.nLength				= sizeof(m_sa);
		m_sa.bInheritHandle			= FALSE;
		m_sa.lpSecurityDescriptor	= &m_sd;
	}
 
    return bSuccess;
}


/////////////////////////////////////////////////////////////////////////////
// CSemaphore

CSemaphore::CSemaphore(LONG lInitialCount, LONG lMaxCount,
	LPCTSTR pstrName, LPSECURITY_ATTRIBUTES lpsaAttributes)
	:  CSyncObject(pstrName)
{
	_ASSERTE(lMaxCount > 0);
	_ASSERTE(lInitialCount <= lMaxCount);

	m_hObject = ::CreateSemaphore(lpsaAttributes, lInitialCount, lMaxCount, pstrName);
	if (m_hObject == NULL)
	{
		char szWhat[255];
		wsprintfA(szWhat, "Unable to create semaphore: %d", ::GetLastError());
		throw exception(szWhat);
	}
}

CSemaphore::~CSemaphore()
{
}

BOOL CSemaphore::Unlock(LONG lCount, LPLONG lpPrevCount /* =NULL */)
{
	return ::ReleaseSemaphore(m_hObject, lCount, lpPrevCount);
}

/////////////////////////////////////////////////////////////////////////////
// CMutex

CMutex::CMutex(
	BOOL bInitiallyOwn, 
	LPCTSTR pstrName, 
	LPSECURITY_ATTRIBUTES lpsaAttribute,
	ESecurityAttribute SecurityAttribute) : 
		CSyncObject(pstrName)
{
	// To share this mutex when it is created by Eserver running as a service, we must place a null
	// DACL in the security descriptor field of ::CreateMutex().
	// See Knowledge Base article Q106387.
	if (lpsaAttribute == NULL && SecurityAttribute == saEveryone && CreateAccessibleSecurityDescriptor())
	{
		// Caller has not specified security attributes structure, but has asked for one to be 
		// created which grants everyone access rights to the mutex.
		lpsaAttribute = &m_sa;
	}

	m_hObject = ::CreateMutex(lpsaAttribute, bInitiallyOwn, pstrName);

	if (m_hObject == NULL)
	{
		if (::GetLastError() == 5)
		{
			m_hObject = ::OpenMutex(NULL, bInitiallyOwn, pstrName);
		}
		else
		{
			char szWhat[255];
			wsprintfA(szWhat, "Unable to create mutex: %d", ::GetLastError());
			throw exception(szWhat);
		}
	}
}

CMutex::~CMutex()
{
}

BOOL CMutex::Unlock()
{
	return ::ReleaseMutex(m_hObject);
}

/////////////////////////////////////////////////////////////////////////////
// CEvent

CEvent::CEvent(BOOL bInitiallyOwn, BOOL bManualReset, LPCTSTR pstrName,
	LPSECURITY_ATTRIBUTES lpsaAttribute)
	: CSyncObject(pstrName)
{
	m_hObject = ::CreateEvent(lpsaAttribute, bManualReset, bInitiallyOwn, pstrName);
	if (m_hObject == NULL)
	{
		char szWhat[255];
		wsprintfA(szWhat, "Unable to create event: %d", ::GetLastError());
		throw exception(szWhat);
	}
}

CEvent::~CEvent()
{
}

BOOL CEvent::Unlock()
{
	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// CSingleLock

CSingleLock::CSingleLock(CSyncObject* pObject, BOOL bInitialLock)
{
	_ASSERTE(pObject != NULL);
//	_ASSERTE(pObject->IsKindOf(RUNTIME_CLASS(CSyncObject)));

	m_pObject = pObject;
	m_hObject = pObject->m_hObject;
	m_bAcquired = FALSE;

	if (bInitialLock)
		Lock();
}

BOOL CSingleLock::Lock(DWORD dwTimeOut /* = INFINITE */)
{
	_ASSERTE(m_pObject != NULL || m_hObject != NULL);
	_ASSERTE(!m_bAcquired);

	m_bAcquired = m_pObject->Lock(dwTimeOut);
	return m_bAcquired;
}

BOOL CSingleLock::Unlock()
{
	_ASSERTE(m_pObject != NULL);
	if (m_bAcquired)
		m_bAcquired = !m_pObject->Unlock();

	// successfully unlocking means it isn't acquired
	return !m_bAcquired;
}

BOOL CSingleLock::Unlock(LONG lCount, LPLONG lpPrevCount /* = NULL */)
{
	_ASSERTE(m_pObject != NULL);
	if (m_bAcquired)
		m_bAcquired = !m_pObject->Unlock(lCount, lpPrevCount);

	// successfully unlocking means it isn't acquired
	return !m_bAcquired;
}

/////////////////////////////////////////////////////////////////////////////
// CMultiLock

CMultiLock::CMultiLock(CSyncObject* pObjects[], DWORD dwCount,
	BOOL bInitialLock)
{
	_ASSERTE(dwCount > 0 && dwCount <= MAXIMUM_WAIT_OBJECTS);
	_ASSERTE(pObjects != NULL);

	m_ppObjectArray = pObjects;
	m_dwCount = dwCount;

	// as an optimization, skip alloacating array if
	// we can use a small, predeallocated bunch of handles

	if (m_dwCount > (sizeof(m_hPreallocated) / sizeof(m_hPreallocated[0])))
	{
		m_pHandleArray = new HANDLE[m_dwCount];
		m_bLockedArray = new BOOL[m_dwCount];
	}
	else
	{
		m_pHandleArray = m_hPreallocated;
		m_bLockedArray = m_bPreallocated;
	}

	// get list of handles from array of objects passed
	for (DWORD i = 0; i <m_dwCount; i++)
	{
		// ASSERT_VALID(pObjects[i]);
		// _ASSERTE(pObjects[i]->IsKindOf(RUNTIME_CLASS(CSyncObject)));

		// can't wait for critical sections

		// _ASSERTE(!pObjects[i]->IsKindOf(RUNTIME_CLASS(CCriticalSection)));

		m_pHandleArray[i] = pObjects[i]->m_hObject;
		m_bLockedArray[i] = FALSE;
	}

	if (bInitialLock)
		Lock();
}

CMultiLock::~CMultiLock()
{
	Unlock();
	if (m_pHandleArray != m_hPreallocated)
	{
		delete[] m_bLockedArray;
		delete[] m_pHandleArray;
	}
}

DWORD CMultiLock::Lock(DWORD dwTimeOut /* = INFINITE */,
		BOOL bWaitForAll /* = TRUE */, DWORD dwWakeMask /* = 0 */)
{
	DWORD dwResult;
	if (dwWakeMask == 0)
		dwResult = ::WaitForMultipleObjects(m_dwCount,
			m_pHandleArray, bWaitForAll, dwTimeOut);
	else
		dwResult = ::MsgWaitForMultipleObjects(m_dwCount,
			m_pHandleArray, bWaitForAll, dwTimeOut, dwWakeMask);

	if (dwResult >= WAIT_OBJECT_0 && dwResult < (WAIT_OBJECT_0 + m_dwCount))
	{
		if (bWaitForAll)
		{
			for (DWORD i = 0; i < m_dwCount; i++)
				m_bLockedArray[i] = TRUE;
		}
		else
		{
			_ASSERTE((dwResult - WAIT_OBJECT_0) >= 0);
			m_bLockedArray[dwResult - WAIT_OBJECT_0] = TRUE;
		}
	}
	return dwResult;
}

BOOL CMultiLock::Unlock()
{
	for (DWORD i=0; i < m_dwCount; i++)
	{
		if (m_bLockedArray[i])
			m_bLockedArray[i] = !m_ppObjectArray[i]->Unlock();
	}
	return TRUE;
}

BOOL CMultiLock::Unlock(LONG lCount, LPLONG lpPrevCount /* =NULL */)
{
	BOOL bGotOne = FALSE;
	for (DWORD i=0; i < m_dwCount; i++)
	{
		if (m_bLockedArray[i])
		{
			CSemaphore* pSemaphore = (CSemaphore*)(m_ppObjectArray[i]);
			if (pSemaphore != NULL)
			{
				bGotOne = TRUE;
				m_bLockedArray[i] = !m_ppObjectArray[i]->Unlock(lCount, lpPrevCount);
			}
		}
	}

	return bGotOne;
}

/////////////////////////////////////////////////////////////////////////////
// CCriticalSection


/*
inline BOOL InitializeCriticalSectionAndSpinCountProxy(LPCRITICAL_SECTION lpCriticalSection, DWORD)
{
	::InitializeCriticalSection(lpCriticalSection);
	return TRUE;
}

inline DWORD SetCriticalSectionSpinCountProxy(LPCRITICAL_SECTION, DWORD)
{
	return 0;
}

FARPROC CCriticalSection::m_fpInitializeCriticalSectionAndSpinCount = NULL;
FARPROC CCriticalSection::m_fpSetCriticalSectionSpinCountProc		= NULL;

FARPROC CCriticalSection::GetInitializeCriticalSectionAndSpinCountProc()
{
	if (!m_fpInitializeCriticalSectionAndSpinCount)
	{
		// function pointer not already resolved
		m_fpInitializeCriticalSectionAndSpinCount = 
			::GetProcAddress(::GetModuleHandle("KERNEL32.DLL"), "InitializeCriticalSectionAndSpinCount");
	
		if (!m_fpInitializeCriticalSectionAndSpinCount)
		{
			// InitializeCriticalSectionAndSpinCount not exported by KERNEL32.DLL, i.e., not using
			// Windows NT 4 Service Pack 3 or greater, therefore use proxy function
			m_fpInitializeCriticalSectionAndSpinCount = (FARPROC)InitializeCriticalSectionAndSpinCountProxy;
		}
	}

	return m_fpInitializeCriticalSectionAndSpinCount;
}

FARPROC CCriticalSection::GetSetCriticalSectionSpinCountProc()
{
	if (!m_fpSetCriticalSectionSpinCountProc)
	{
		// function pointer not already resolved
		m_fpSetCriticalSectionSpinCountProc = 
			::GetProcAddress(::GetModuleHandle("KERNEL32.DLL"), "SetCriticalSectionSpinCount");
	
		if (!m_fpSetCriticalSectionSpinCountProc)
		{
			// SetCriticalSectionSpinCount not exported by KERNEL32.DLL, i.e., not using
			// Windows NT 4 Service Pack 3 or greater, therefore use proxy function
			m_fpSetCriticalSectionSpinCountProc = (FARPROC)SetCriticalSectionSpinCountProxy;
		}
	}

	return m_fpSetCriticalSectionSpinCountProc;
}
*/

}	// namespaceMutex
