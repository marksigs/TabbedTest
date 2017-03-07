/******************************************************************************
Module name: Optex.cpp
Written by:  Jeffrey Richter
Purpose:     Implements the COptex (optimized mutex) synchronization object
******************************************************************************/


#include "stdafx.h"
#include "CmdHdr.h"
#include <windows.h>
#include "Optex.h"


///////////////////////////////////////////////////////////////////////////////

BOOL COptex::CommonConstructor(PVOID pszName, BOOL fUnicode, DWORD dwSpinCount) 
{
	DWORD dwLastError = NULL;

	m_hevt = m_hfm = m_pSharedInfo = NULL;

	SYSTEM_INFO sinf;
	GetSystemInfo(&sinf);
	m_fUniprocessorHost = (sinf.dwNumberOfProcessors == 1);

	char szNameA[100];
	if (fUnicode) 
	{   
		// Convert Unicode name to ANSI
		wsprintfA(szNameA, "%S", pszName);
		pszName = (PVOID) szNameA;
	}
	char sz[100];
	wsprintfA(sz, "JMR_Optex_Event_%s", pszName);

	SECURITY_ATTRIBUTES sa;
	SECURITY_DESCRIPTOR sd;
	LPSECURITY_ATTRIBUTES lpsa = SetSecurityDescriptorNullDacl(&sa, &sd);

	m_hevt = CreateEventA(lpsa, FALSE, FALSE, sz);
	dwLastError = GetLastError();
	chASSERT(m_hevt != NULL);
	if (m_hevt != NULL) 
	{
		wsprintfA(sz, "JMR_Optex_MMF_%s", pszName);
		m_hfm = CreateFileMappingA((HANDLE)MMF_PAGING_FILE, lpsa, PAGE_READWRITE, 0, sizeof(*m_pSharedInfo), sz);
		dwLastError = GetLastError();
		chASSERT(m_hfm != NULL);
		if (m_hfm != NULL) 
		{
			m_pSharedInfo = (PSHAREDINFO) MapViewOfFile(m_hfm, FILE_MAP_ALL_ACCESS, 0, 0, 0);

			// Note: SHAREDINFO's m_lLockCount, m_dwThreadId, and m_lRecurseCount
			// members need to be initialized to 0. Fortunately, a new pagefile 
			// MMF sets all of its data to 0 when created. This saves use from 
			// some thread synchronization work.

			if (m_pSharedInfo != NULL)
			{
				SetSpinCount(dwSpinCount);
			}
		}
	}

	return IsValid();
}


///////////////////////////////////////////////////////////////////////////////


COptex::~COptex() 
{
#ifdef _DEBUG
	if (m_pSharedInfo && m_pSharedInfo->m_dwThreadId == GetCurrentThreadId()) 
	{
		DebugBreak();
	}
#endif
	if (m_pSharedInfo)
	{
		UnmapViewOfFile(m_pSharedInfo);
	}
	if (m_hfm)
	{
		CloseHandle(m_hfm);
	}
	if (m_hevt)
	{
		CloseHandle(m_hevt);
	}
}


///////////////////////////////////////////////////////////////////////////////


void COptex::SetSpinCount(DWORD dwSpinCount) 
{
	if (!m_fUniprocessorHost && m_pSharedInfo)
	{
		InterlockedExchange((PLONG)&m_pSharedInfo->m_dwSpinCount, dwSpinCount);
	}
}


///////////////////////////////////////////////////////////////////////////////


void COptex::Enter() 
{
	// Spin, trying to get the Optex
	if (TryEnter()) return;

	DWORD dwThreadId = GetCurrentThreadId();  // The calling thread's ID
	if (InterlockedIncrement(&m_pSharedInfo->m_lLockCount) == 1) 
	{
		  // Optex is unowned, let this thread own it once
		InterlockedExchange((PLONG) &m_pSharedInfo->m_dwThreadId, dwThreadId);
		m_pSharedInfo->m_lRecurseCount = 1;
	}
	else 
	{
		// Optex is owned by a thread
		if (m_pSharedInfo->m_dwThreadId == dwThreadId) 
		{
			// Optex is owned by this thread, own it again
			m_pSharedInfo->m_lRecurseCount++;
		}
		else 
		{
			// Optex is owned by another thread
			// Wait for the owning thread to release the Optex
			WaitForSingleObject(m_hevt, INFINITE);

			// We got ownership of the Optex
			InterlockedExchange((PLONG) &m_pSharedInfo->m_dwThreadId, dwThreadId); // We own it now
			m_pSharedInfo->m_lRecurseCount = 1;       // We own it once
		}
	}
}


///////////////////////////////////////////////////////////////////////////////


BOOL COptex::TryEnter() 
{
	DWORD dwThreadId = GetCurrentThreadId();  // The calling thread's ID

	// If the lock count is zero, the Optex is unowned and
	// this thread can become the owner of it now.
	BOOL fThisThreadOwnsTheOptex = FALSE;
	DWORD dwSpinCount = m_pSharedInfo->m_dwSpinCount;
	do 
	{
		fThisThreadOwnsTheOptex = (0 == (DWORD) 
		InterlockedCompareExchange((PVOID*) &m_pSharedInfo->m_lLockCount, (PVOID) 1, (PVOID) 0)); 

		if (fThisThreadOwnsTheOptex) 
		{
			// We now own the Optex
			InterlockedExchange((PLONG) &m_pSharedInfo->m_dwThreadId, 
			dwThreadId); // We own it
			m_pSharedInfo->m_lRecurseCount = 1;		// We own it once
		}
		else
		{
			// Some thread owns the Optex
			if (m_pSharedInfo->m_dwThreadId == dwThreadId) 
			{
				// We already own the Optex
				InterlockedIncrement(&m_pSharedInfo->m_lLockCount);
				m_pSharedInfo->m_lRecurseCount++;	// We own it again
				fThisThreadOwnsTheOptex = TRUE;		// Return that we own the Optex
			}
		}
	} while (!fThisThreadOwnsTheOptex && (dwSpinCount-- > 0));

	// Return whether or not this thread owns the Optex
	return fThisThreadOwnsTheOptex;
}


///////////////////////////////////////////////////////////////////////////////


void COptex::Leave()
{
#ifdef _DEBUG
	if (m_pSharedInfo->m_dwThreadId != GetCurrentThreadId()) 
	{
		DebugBreak();
	}
#endif

	if (--m_pSharedInfo->m_lRecurseCount > 0) 
	{
		// We still own the Optex
		InterlockedDecrement(&m_pSharedInfo->m_lLockCount);
	}
	else 
	{
		// We don't own the Optex
		InterlockedExchange((PLONG)&m_pSharedInfo->m_dwThreadId, 0);
		if (InterlockedDecrement(&m_pSharedInfo->m_lLockCount) > 0) 
		{
			// Other threads are waiting, wake one of them
			SetEvent(m_hevt);
		}
	}
}


///////////////////////////////////////////////////////////////////////////////
// AS 09/10/00
// Creates a security descriptor than can be shared when the application is
// running as a service.
// See Knowledge Base article Q106387.

LPSECURITY_ATTRIBUTES SetSecurityDescriptorNullDacl(LPSECURITY_ATTRIBUTES lpsa, PSECURITY_DESCRIPTOR psd)
{
	if	(
			lpsa != NULL && psd != NULL &&
			::InitializeSecurityDescriptor(psd, SECURITY_DESCRIPTOR_REVISION) &&
			// Add a NULL DACL to the security descriptor..
			::SetSecurityDescriptorDacl(psd, TRUE, (PACL)NULL, FALSE)
		)
	{
		lpsa->nLength = sizeof(SECURITY_ATTRIBUTES);
		lpsa->lpSecurityDescriptor = psd;
		lpsa->bInheritHandle = TRUE;
	}
	else
	{
		lpsa = NULL;
	}

	return lpsa;
}


///////////////////////////////// End of File /////////////////////////////////
