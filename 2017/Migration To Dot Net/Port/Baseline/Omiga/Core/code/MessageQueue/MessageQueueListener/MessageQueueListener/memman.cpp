///////////////////////////////////////////////////////////////////////////////
//	FILE:			memman.cpp
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
#include <new.h>
//#include "eserver.h"
//#include "EserverPrfMon.h"
//#include "log.h"
#include "memman.h"

namespace namespaceMemMan
{
//using namespace namespaceLog;

// need to undefine here so that we call the real CRT functions, and not the macros in memman.h
#undef malloc
#undef realloc
#undef calloc
#undef free
#undef _heapmin
#undef new
#undef delete


CMpHeap::CMpHeap() : 
	m_hMpHeap(NULL), 
	m_flOptions(NULL),
	m_dwInitialSizeBytes(NULL),
	m_dwParallelism(NULL),
	m_dwHeaps(NULL),
	m_bValidate(FALSE)
{
}

CMpHeap::CMpHeap(DWORD flOptions, DWORD dwInitialSize, DWORD dwParallelism) : 
	m_hMpHeap(NULL), 
	m_flOptions(NULL),
	m_dwInitialSizeBytes(dwInitialSize),
	m_dwParallelism(dwParallelism),
	m_dwHeaps(NULL),
	m_bValidate(FALSE)
{
	Create(flOptions, dwInitialSize, dwParallelism);
}

CMpHeap::~CMpHeap()
{ 
	Destroy();
}

BOOL CMpHeap::Destroy()
{
	BOOL bSuccess = TRUE;

	if (m_hMpHeap)
	{
		bSuccess = ::MpHeapDestroy(m_hMpHeap);
	}

	return bSuccess;
}

#ifdef _DEBUG
LPVOID CMpHeap::Alloc(DWORD flOptions, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine)
#else
LPVOID CMpHeap::Alloc(DWORD flOptions, DWORD dwBytes)
#endif
{
	LPVOID pMem = NULL;

	if (m_hMpHeap)
	{
		pMem = ::MpHeapAlloc(m_hMpHeap, flOptions, dwBytes);
	}
	else
	{
#ifdef _DEBUG
		pMem = ::_malloc_dbg(dwBytes, nBlockType, pszFileName, nLine);
#else
		pMem = ::malloc(dwBytes);
#endif
	}

	return pMem;
}

#ifdef _DEBUG
LPVOID CMpHeap::Calloc(DWORD dwNum, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine) 
#else
LPVOID CMpHeap::Calloc(DWORD dwNum, DWORD dwBytes) 
#endif
{
	LPVOID pMem = NULL;

	if (m_hMpHeap)
	{
		pMem = ::MpHeapAlloc(m_hMpHeap, MPHEAP_ZERO_MEMORY, dwBytes * dwNum);
	}
	else
	{
		// use malloc and not calloc as block size is 4 bytes greater than dwBytes * dwNum
		// (the extra byte is the flag byte)
#ifdef _DEBUG
		pMem = (LPDWORD)::_calloc_dbg(dwNum, dwBytes, nBlockType, pszFileName, nLine);
#else
		pMem = (LPDWORD)::calloc(dwNum, dwBytes);
#endif
	}

	return pMem;
}

#ifdef _DEBUG
LPVOID CMpHeap::ReAlloc(LPVOID pMem, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine)
#else
LPVOID CMpHeap::ReAlloc(LPVOID pMem, DWORD dwBytes)
#endif
{
	LPVOID pOldMem = pMem;
	LPVOID pNewMem = NULL;

	if (pOldMem)
	{
		// memory already allocated

		if (m_hMpHeap)
		{
			// memory was allocated on an MP heap, so reallocate there
			if (dwBytes == 0)
			{
				// ANSI: reallocated with 0 bytes frees memory
				::MpHeapFree(m_hMpHeap, pOldMem);
				pNewMem = NULL;
			}
			else 
			{
				pNewMem = ::MpHeapReAlloc(m_hMpHeap, pOldMem, dwBytes);
			}
		}
		else 
		{
			// memory was allocated on normal heap
#ifdef _DEBUG
			pNewMem = ::_realloc_dbg(pOldMem, dwBytes, nBlockType, pszFileName, nLine);
#else
			pNewMem = ::realloc(pOldMem, dwBytes);
#endif
		}
	}
	else if (m_hMpHeap)
	{
		// memory not already allocated, and MP heap exists, so allocate on MP heap
		pNewMem = ::MpHeapAlloc(m_hMpHeap, NULL, dwBytes);
	}
	else
	{
		// memory not already allocated, and MP Heap does not exist, so allocate on normal heap
#ifdef _DEBUG
		pNewMem = ::_realloc_dbg(pMem, dwBytes, nBlockType, pszFileName, nLine);
#else
		pNewMem = ::realloc(pMem, dwBytes);
#endif
	}

	return pNewMem;
}

#ifdef _DEBUG
LPVOID CMpHeap::New(size_t nSize, int nBlockType, const char* pszFileName, int nLine)
#else
LPVOID CMpHeap::New(size_t nSize)
#endif
{
	LPVOID pMem = NULL;

	if (nSize == 0)			// handle 0 bytes requests by treating them as 1-byte requests
	{
		nSize = 1;
	}
	while (pMem == NULL)
	{
		if (m_hMpHeap)
		{
			pMem = ::MpHeapAlloc(m_hMpHeap, NULL, nSize);
		}
		else
		{
#ifdef _DEBUG
			pMem = ::_malloc_dbg(nSize, nBlockType, pszFileName, nLine);
#else
			pMem = ::malloc(nSize);
#endif
		}

		if (pMem)
		{
			break;
		}
		else
		{
			// memory allocation failed

			_PNH globalHandler = _query_new_handler();
			if (!globalHandler || (*globalHandler)(nSize))
			{
				// no global new handler, or there is one and it failed, so raise exception
				// (otherwise loop round again)
				throw std::bad_alloc();
			}
		}
	}

	return pMem;
}

BOOL CMpHeap::Free(LPVOID pMem)
{
	BOOL bSuccess = TRUE;

	if (pMem)
	{
		if (m_hMpHeap)
		{
			// memory was allocated on MP heap
			bSuccess = ::MpHeapFree(m_hMpHeap, pMem);
		}
		else
		{
			// memory was allocated on standard heap
#ifdef _DEBUG
			::_free_dbg(pMem, _NORMAL_BLOCK);
#else
			::free(pMem);
#endif
		}
	}

	return bSuccess;
}

int CMpHeap::Min()
{
	Compact();
	// always compact normal heap
	return ::_heapmin();
}

UINT CMpHeap::Compact()
{
	UINT nLargestFreeSize = 0;

	if (m_hMpHeap)
	{
		nLargestFreeSize = ::MpHeapCompact(m_hMpHeap);
	}

	return nLargestFreeSize;
}

BOOL CMpHeap::GetStatistics(LPDWORD lpdwSize, MPHEAP_STATISTICS Statistics[])
{
	BOOL bSuccess = FALSE;		// returns FALSE by default (i.e., if not using MpHeap

	if (m_hMpHeap)
	{
		bSuccess = ::MpHeapGetStatistics(m_hMpHeap, lpdwSize, Statistics) == ERROR_SUCCESS;
	}

	return bSuccess;
}

BOOL CMpHeap::Initialise()
{
	BOOL bInitialised = TRUE;

	static const char* pszMpHeap						= "MpHeap";
	static const char* pszEnabled						= "Enabled";
	static const char* pszInitialSizeBytes				= "InitialSizeBytes";
	static const char* pszParallelism					= "Parallelism";
	static const char* pszGrowable						= "Growable";
	static const char* pszReallocInPlaceOnly			= "ReallocInPlaceOnly";
	static const char* pszTailCheckingEnabled			= "TailCheckingEnabled";
	static const char* pszFreeCheckingEnabled			= "FreeCheckingEnabled";
	static const char* pszDisableCoalesceOnFree			= "DisableCoalesceOnFree";
	static const char* pszZeroMemory					= "ZeroMemory";
	static const char* pszCollectStats					= "CollectStats";
	static const char* pszValidate						= "Validate";

	BOOL bMpHeap = EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszEnabled, "N"));

	if (bMpHeap)
	{
		IF_LOG(g_pLog->Report("Multiple Heaps: enabled\n"));

		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszGrowable, "N")))
		{
			m_flOptions |= MPHEAP_GROWABLE;
			IF_LOG(g_pLog->Report("Multiple Heaps: growable\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszReallocInPlaceOnly, "N")))
		{
			m_flOptions |= MPHEAP_REALLOC_IN_PLACE_ONLY;
			IF_LOG(g_pLog->Report("Multiple Heaps: realloc in place only\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszTailCheckingEnabled, "N")))
		{
			m_flOptions |= MPHEAP_TAIL_CHECKING_ENABLED;
			IF_LOG(g_pLog->Report("Multiple Heaps: tail checking enabled\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszFreeCheckingEnabled, "N")))
		{
			m_flOptions |= MPHEAP_FREE_CHECKING_ENABLED;
			IF_LOG(g_pLog->Report("Multiple Heaps: free checking enabled\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszDisableCoalesceOnFree, "N")))
		{
			m_flOptions |= MPHEAP_DISABLE_COALESCE_ON_FREE;
			IF_LOG(g_pLog->Report("Multiple Heaps: disable coalesce on free\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszZeroMemory, "N")))
		{
			m_flOptions |= MPHEAP_ZERO_MEMORY;
			IF_LOG(g_pLog->Report("Multiple Heaps: zero memory\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszCollectStats, "N")))
		{
			m_flOptions |= MPHEAP_COLLECT_STATS;
			IF_LOG(g_pLog->Report("Multiple Heaps: collect stats\n"));
		}
		if (EServerApp.IsTrueString(EServerApp.GetProfileString(pszMpHeap, pszValidate, "N")))
		{
			m_bValidate = TRUE;
			IF_LOG(g_pLog->Report("Multiple Heaps: validate\n"));
		}

		m_dwInitialSizeBytes = EServerApp.GetProfileInt(pszMpHeap, pszInitialSizeBytes, 0);
		IF_LOG(g_pLog->Report("Multiple Heaps: initial size = %ld bytes\n", m_dwInitialSizeBytes));
		m_dwParallelism	= EServerApp.GetProfileInt(pszMpHeap, pszParallelism, 0);
		if (m_dwParallelism == 0) 
		{ 
			SYSTEM_INFO SystemInfo; 
 			GetSystemInfo(&SystemInfo); 
			m_dwHeaps = 3 + SystemInfo.dwNumberOfProcessors; 
		} 
		else
		{
			m_dwHeaps = m_dwParallelism;
		}
		IF_LOG(g_pLog->Report("Multiple Heaps: parallelism = %ld\n", m_dwParallelism));
		IF_LOG(g_pLog->Report("Multiple Heaps: heaps = %ld\n", m_dwHeaps));

		bInitialised = Create(m_flOptions, m_dwInitialSizeBytes, m_dwParallelism) != NULL;
	}

	return bInitialised;
}

BOOL CMpHeap::LogStatistics()
{
	BOOL bSuccess = TRUE;

	if (g_pEServerPrfMon != NULL && EServerApp.GetPrfMonEnabled())
	{
		if ((m_flOptions & MPHEAP_COLLECT_STATS) == MPHEAP_COLLECT_STATS)
		{
			DWORD dwStatSize = m_dwHeaps * sizeof(MPHEAP_STATISTICS); 
			LPMPHEAP_STATISTICS lpHeapStats = (LPMPHEAP_STATISTICS)::LocalAlloc(LMEM_FIXED, dwStatSize);

			if (lpHeapStats) 
			{ 
				bSuccess = GetStatistics(&dwStatSize, lpHeapStats);

				if (bSuccess) 
				{ 
					for (DWORD dwStats = 0; dwStats < dwStatSize / sizeof(MPHEAP_STATISTICS); dwStats++)
					{ 
						g_pEServerPrfMon->GetCtr32(MPHEAP_ALLOCATIONS, g_pEServerPrfMon->GetInstIdMpHeap()) = lpHeapStats[dwStats].TotalAllocates;

						if (0 == lpHeapStats[dwStats].TotalAllocates) 
						{
							continue;  // avoid divide by 0 
						}
						g_pEServerPrfMon->GetCtr32(MPHEAP_CONTENTION, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							lpHeapStats[dwStats].Contention;
						g_pEServerPrfMon->GetCtr32(MPHEAP_CONTENTION_PERCENT, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							100 * lpHeapStats[dwStats].Contention / lpHeapStats[dwStats].TotalAllocates;
						g_pEServerPrfMon->GetCtr32(MPHEAP_LOOKASIDE_ALLOCS, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							lpHeapStats[dwStats].LookasideAllocates;
						g_pEServerPrfMon->GetCtr32(MPHEAP_LOOKASIDE_ALLOCS_PERCENT, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							100 * lpHeapStats[dwStats].LookasideAllocates / lpHeapStats[dwStats].TotalAllocates;

						g_pEServerPrfMon->GetCtr32(MPHEAP_FREES, g_pEServerPrfMon->GetInstIdMpHeap()) = lpHeapStats[dwStats].TotalFrees; 
						if (0 == lpHeapStats[dwStats].TotalFrees) 
						{
							continue;  // avoid divide by 0 
						}
						g_pEServerPrfMon->GetCtr32(MPHEAP_LOOKASIDE_FREES, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							lpHeapStats[dwStats].LookasideFrees;
						g_pEServerPrfMon->GetCtr32(MPHEAP_LOOKASIDE_FREES_PERCENT, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							100 * lpHeapStats[dwStats].LookasideFrees / lpHeapStats[dwStats].TotalFrees; 

						g_pEServerPrfMon->GetCtr32(MPHEAP_DELAYED_FREES, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							lpHeapStats[dwStats].DelayedFrees;
						g_pEServerPrfMon->GetCtr32(MPHEAP_DELAYED_FREES_PERCENT, g_pEServerPrfMon->GetInstIdMpHeap()) = 
							100 * lpHeapStats[dwStats].DelayedFrees / lpHeapStats[dwStats].TotalFrees; 
					}
				} 
				::LocalFree(lpHeapStats);
			}
		}
	}

	return bSuccess;
}

BOOL CMpHeap::Validate(LPVOID pMem)
{
	BOOL bSuccess = TRUE;

	if (m_bValidate && m_hMpHeap)
	{
		bSuccess = ::MpHeapValidate(m_hMpHeap, pMem);

		if (!bSuccess)
		{
			IF_LOG_FLAGS(g_pLog->Report(NULL, g_pLog->GetGenericErrorFlags(), "MpHeap validation error"), g_pLog->GetGenericErrorFlags());
		}
	}

	return bSuccess;
}

}	// namespaceMpHeap

// Singleton global MpHeap object - created on the heap by EServerApp object.
namespaceMemMan::CMpHeap* g_pMpHeap = NULL;
