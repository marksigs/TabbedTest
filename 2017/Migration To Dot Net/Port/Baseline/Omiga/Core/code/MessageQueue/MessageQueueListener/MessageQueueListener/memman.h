///////////////////////////////////////////////////////////////////////////////
//	FILE:			memman.h
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

#ifndef MEMMAN_H
#define MEMMAN_H

#pragma message(MESSAGE("->Compiling MEMMAN.H"))

#include "mpheap.h"

namespace namespaceMemMan
{

class CMpHeap
{
private:
	HANDLE				m_hMpHeap;
	DWORD				m_flOptions;
	DWORD				m_dwInitialSizeBytes;
	DWORD				m_dwParallelism;
	DWORD				m_dwHeaps;
	BOOL				m_bValidate;

public:
	CMpHeap();
	CMpHeap(DWORD flOptions, DWORD dwInitialSize, DWORD dwParallelism);
	~CMpHeap();
	inline HANDLE Create(DWORD flOptions, DWORD dwInitialSizeBytes, DWORD dwParallelism)
	{
		return m_hMpHeap = ::MpHeapCreate(flOptions, dwInitialSizeBytes, dwParallelism);
	}
	BOOL Destroy();
	BOOL Validate(LPVOID lpvMem = NULL);
	BOOL GetStatistics(LPDWORD lpdwSize, MPHEAP_STATISTICS Statistics[]);
	UINT Compact();
	int Min();
#ifdef _DEBUG
	LPVOID Alloc(DWORD flOptions, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine);
	LPVOID Calloc(DWORD dwNum, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine);
	LPVOID ReAlloc(LPVOID pvMem, DWORD dwBytes, int nBlockType, const char* pszFileName, int nLine);
	LPVOID New(size_t nSize, int nBlockType, const char* pszFileName, int nLine);
#else
	LPVOID Alloc(DWORD flOptions, DWORD dwBytes);
	LPVOID Calloc(DWORD dwNum, DWORD dwBytes);
	LPVOID ReAlloc(LPVOID pvMem, DWORD dwBytes);
	LPVOID New(size_t nSize);
#endif
	BOOL Free(LPVOID pvMem);
	DWORD GetOptions() const { return m_flOptions; }
	BOOL Initialise();
	BOOL LogStatistics();
	inline HANDLE GetHandle() const { return m_hMpHeap; }
};

}	// namespaceMpHeap

extern namespaceMemMan::CMpHeap* g_pMpHeap;

#ifdef malloc
#undef malloc
#endif
#ifdef realloc
#undef realloc
#endif
#ifdef calloc
#undef calloc
#endif
#ifdef free
#undef free
#endif
#ifdef new
#undef new
#endif
#ifdef delete
#undef delete
#endif


// map CRT memory functions to MP heap functions
#ifdef _DEBUG

// debug build: map CRT memory functions to MP heap functions debug versions
#define malloc(size)			(g_pMpHeap ? g_pMpHeap->Alloc(0, size, _NORMAL_BLOCK, __FILE__, __LINE__) : ::_malloc_dbg(size, _NORMAL_BLOCK, __FILE__, __LINE__))
#define calloc(num, size)		(g_pMpHeap ? g_pMpHeap->Calloc(num, size, _NORMAL_BLOCK, __FILE__, __LINE__) : ::_calloc_dbg(num, size, _NORMAL_BLOCK, __FILE__, __LINE__))
#define realloc(memory, size)	(g_pMpHeap ? g_pMpHeap->ReAlloc(memory, size, _NORMAL_BLOCK, __FILE__, __LINE__) : ::_realloc_dbg(memory, size, _NORMAL_BLOCK, __FILE__, __LINE__))
#define free(memory)			(g_pMpHeap ? g_pMpHeap->Free(memory) : ::_free_dbg(memory, _NORMAL_BLOCK))

// To use the debug new operator, put the following at the top of each *.cpp file,
// after any #includes:
//		#ifdef _DEBUG
//		#define new DEBUG_NEW
//		#endif
#define DEBUG_NEW new(_NORMAL_BLOCK, __FILE__, __LINE__)

#else

// release build: map CRT memory functions to MP heap functions non-debug versions
#define malloc(size)			(g_pMpHeap ? g_pMpHeap->Alloc(0, size) : ::malloc(size))
#define calloc(num, size)		(g_pMpHeap ? g_pMpHeap->Calloc(num, size) : ::calloc(num, size))
#define realloc(memory, size)	(g_pMpHeap ? g_pMpHeap->ReAlloc(memory, size) : ::realloc(memory, size))
#define free(memory)			(g_pMpHeap ? g_pMpHeap->Free(memory) : ::free(memory))

#endif


#define _heapmin()				(g_pMpHeap ? g_pMpHeap->Min() : ::_heapmin())


#pragma message(MESSAGE("<-Compiled MEMMAN.H"))

#endif
