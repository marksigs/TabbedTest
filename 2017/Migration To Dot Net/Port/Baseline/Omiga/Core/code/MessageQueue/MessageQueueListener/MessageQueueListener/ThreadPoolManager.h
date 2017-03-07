///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManager.h
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

#ifndef THREADPOOLMANAGER_H
#define THREADPOOLMANAGER_H

#pragma message(MESSAGE("->Compiling THREADPOOLMANAGER.H"))

#include "mutex.h"
#include "ThreadPoolMessage.h"

class CThreadPoolThread;

///////////////////////////////////////////////////////////////////////////////

// OVERVIEW
// 1. Suspended threads are created by the thread pool manager (StartUp() calls CreateSuspendedThread() for each instance
//    and they are assigned to their ideal processor.  i.e for a four processor machine the sequence of threads on which
//    they run are 0, 1, 2, 3, 0, 1, 2, 3, 0 ... etc)
// 2. The thread pool manager get a request for a thread to run a function
//      2.1. If the function queue is empty, the thread pool manager finds an an available thread or queues the function if none is available
//           (It updates the function to be called whilst the thread is suspended (SetRuntimeFunction()) and then resumes the
//            suspended thread)
//      2.4. The thread does the work according to the runtime function set up (see ThreadEntry())
//      2.5. When the thread is finished it asks for more work from the thread pool manager (RemoveMessageFromQueue()).  If there is 
//           no work available (queue is empty) then the thread suspends itself
// 3. The thread pool manager get a request to close down 
//      3.1. The thread pool manger asks each thread to close down (CloseDownAndDelete())
//      3.2  The program must wait for all threads to exit (WaitForAllThreadsToCloseDown());
//          

///////////////////////////////////////////////////////////////////////////////

class CThreadPoolManager : public CObject
{
protected:
	CThreadPoolManager();
	virtual ~CThreadPoolManager();

public:
    // initialisation
	enum {eDefaultNumberOfThreads = -1};
    virtual bool StartUp(int nNumberOfThreads = eDefaultNumberOfThreads); // returns true if atleast one thread is created
    virtual void CloseDown();

	// dynamic resizing of thread pool
	int GetnRequestedNumberOfThreads();
	void SetnRequestedNumberOfThreads(int nNumberOfThreads);

	// snapshop information information
	int GetnThreadsActive(); // number of threads not suspended
	int GetnThreadsNotActive(); // number of threads suspended

public:
	CThreadPoolThread* TryGetpAvailableThreadPoolThread(); // returns NULL if thread is not available
	namespaceMutex::CSharedLockSmallOperations& GetrsharedlockResizeThreads() {return m_sharedlockResizeThreads;}

	static DWORD DetermineNumberOfProcessors();

protected:
	// overrideables - allows specialisation of queues and threads
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread) = 0; // returns NULL if there is no more work to be done 
	virtual CThreadPoolThread* CreateSuspendedThread(); // returns NULL if failed (object returned is valid until CloseDownAndDelete())

	// callbacks from threads
	friend class CThreadPoolThread;
	long OnThreadCreated();
	long OnThreadDestroyed(HANDLE hThreadClosingDown);
	int GetWorkerThreadPriority() { return m_nWorkerThreadPriority; }
	
	// helper functions for resizing thread pool
	void IncreaseThreads(int nNumberOfThreads);
	void DecreaseThreads(int nNumberOfThreads);

	void SetWorkerThreadPriority(int nPriority) { m_nWorkerThreadPriority = nPriority; }
	void WaitForAllThreadsToCloseDown();

    DWORD GetNextIdealProcessor();
    static DWORD s_dwNumberOfProcessors; // cached value

protected:
	// performance monitoring
	void IncHitCtr(BOOL bHit);

    CThreadPoolThread* volatile m_pThreadPoolThreadHead;

private:
    long m_lThreadsCreated; // number of instances of CThreadPoolThread which exist
    HANDLE m_hLastThreadDestroyedEvent; // event raised when all threads have closed down
    HANDLE m_hThreadLastToExit; // handle of last thread to close down
	int m_nWorkerThreadPriority;

	bool m_bThreadPoolManagerStarted; // set to true between StartUp() and CloseDown()

	int m_nRequestedNumberOfThreads; // number of threads requested (this will be out of sync with the number of threads when resizing)
	int m_nActualNumberOfThreads; // actual number of threads (this will be out of sync with the number of threads when resizing)

    namespaceMutex::CSharedLockSmallOperations m_sharedlockResizeThreads;

	// performance monitoring
	LONG	m_lHitCtr;
	LONG	m_lHitCtrTotal;
	LONG	m_lHitCtrPercent;
};

#pragma message(MESSAGE("<-Compiled THREADPOOLMANAGER.H"))

#endif // #ifndef THREADPOOLMANAGER_H
