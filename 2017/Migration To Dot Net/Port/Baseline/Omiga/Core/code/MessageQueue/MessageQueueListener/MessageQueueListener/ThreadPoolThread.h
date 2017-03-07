///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolThread.h
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

#if !defined(AFX_THREADPOOLTHREAD_H__A94D0FA7_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
#define AFX_THREADPOOLTHREAD_H__A94D0FA7_4B60_11D4_8237_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

///////////////////////////////////////////////////////////////////////////////

#pragma message(MESSAGE("->Compiling THREADPOOLTHREAD.H"))

#ifdef __cplusplus
extern "C" {
#endif
// copied from <winbase.h>
WINBASEAPI
DWORD
WINAPI
SetThreadIdealProcessor(
    HANDLE hThread,
    DWORD dwIdealProcessor
    );
#ifdef __cplusplus
}
#endif


class CThreadPoolManager;
class CThreadPoolMessage;

///////////////////////////////////////////////////////////////////////////////

// OVERVIEW (see ThreadPoolManager.h)

///////////////////////////////////////////////////////////////////////////////

// conditional expression is constant - reported for _ASSERTE.
#pragma warning(disable: 4127)

class CThreadPoolThread : public CObject
{
protected:
	CThreadPoolThread(CThreadPoolManager* pThreadPoolManager); // object can only be created when a suspended thread is created
	virtual ~CThreadPoolThread(); // object can only be destroyed on CloseDownAndDelete()

public:
    // Thread Creation called by CThreadPoolManager
    static CThreadPoolThread* CreateSuspendedThread(CThreadPoolManager* pThreadPoolManager); // returns NULL if failed (object returned is valid until CloseDownAndDelete())
    DWORD SetThreadIdealProcessor(DWORD dwIdealProcessor) 
	{
		_ASSERTE(IsThreadCreated()); 
		return ::SetThreadIdealProcessor(m_hThread, dwIdealProcessor);
	}

    // Thread And Object Destruction - self destruct
    // (an extra ResumeThread which calls ensures that the thread cannot be supsended)
    bool CloseDownAndDelete() { m_bKeepThreadAlive = false; return ResumeThread(); }

    void SetNextInSequence(CThreadPoolThread* pThreadPoolThread) 
	{
		m_pThreadPoolThreadNextInSequence = pThreadPoolThread;
	}
    CThreadPoolThread* GetNextInSequence() const {return m_pThreadPoolThreadNextInSequence;} 

    // Thread Status called by CThreadPoolManager
    bool IsThreadCreated() const {return m_hThread != NULL;}
    BOOL IsThreadAvailableForReuse() const {_ASSERTE(IsThreadCreated()); return m_bThreadAvailableForReuse;}
    BOOL TryToReuseThread() 
	{
		_ASSERTE(IsThreadCreated()); 
		return reinterpret_cast<BOOL>(::InterlockedCompareExchange((PVOID*)&m_bThreadAvailableForReuse, (PVOID)FALSE, (PVOID)TRUE));
	}
    void SetThreadAvailableForReuse(BOOL bThreadAvailableForReuse = TRUE) 
	{
		m_bThreadAvailableForReuse = bThreadAvailableForReuse;
	}

    // Thread Runtime called by CThreadPoolManager
    void SetRuntimeFunction(CThreadPoolMessage* pThreadPoolMessage) 
	{
		m_pThreadPoolMessage = pThreadPoolMessage;
	}

    // Thread Activation called by CThreadPoolManager (returns an internal suspend count)
    // ... On the transition from 0 to 1 of the suspend count, the thread is suspended
    // ... On the transition from 1 to 0 of the suspend count, the thread is resumed
	bool ResumeThread();
    bool SuspendThread();
	inline DWORD GetSuspendCount() const { return m_dwSuspendCount; }

protected:
    static DWORD WINAPI ThreadEntry(LPVOID p); 
    DWORD WINAPI ThreadEntry(); 

private:
    DWORD m_dwThreadId;
    HANDLE m_hThread;
    volatile BOOL m_bThreadAvailableForReuse; // must be BOOL (32bit) rather than bool for use with InterlockedCompareExchange
    volatile bool m_bKeepThreadAlive; // made volatile so that it is always re-read
    volatile DWORD m_dwSuspendCount;
    
    CThreadPoolMessage* volatile m_pThreadPoolMessage;
    CThreadPoolThread* m_pThreadPoolThreadNextInSequence;

	CThreadPoolManager* m_pThreadPoolManager; // back-pointer to corresponding manager
};

///////////////////////////////////////////////////////////////////////////////

#pragma message(MESSAGE("<-Compiled THREADPOOLTHREAD.H"))

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_THREADPOOLTHREAD_H__A94D0FA7_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
