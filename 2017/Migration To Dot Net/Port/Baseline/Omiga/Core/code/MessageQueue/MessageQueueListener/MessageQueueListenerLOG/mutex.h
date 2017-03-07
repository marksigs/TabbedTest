///////////////////////////////////////////////////////////////////////////////
//	FILE:			mutex.h
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

#ifndef MUTEX_H
#define MUTEX_H

#include <exception>
// <exception> enables a lot of warnings - so disable them.
#pragma warning(disable: 4018 4097 4100 4127 4146 4201 4238 4244 4511 4512 4663)
#include <string>
#include <crtdbg.h>

WINBASEAPI
BOOL
WINAPI
InitializeCriticalSectionAndSpinCount(
    IN OUT LPCRITICAL_SECTION lpCriticalSection,
    IN DWORD dwSpinCount
    );


namespace namespaceMutex
{
using namespace std;

#pragma warning(disable: 4127)	// conditional expression is constant

class CSyncObject;
class CSemaphore;
class CMutex;
class CEvent;
class CCriticalSection;

class CSingleLock;
class CMultiLock;

/////////////////////////////////////////////////////////////////////////////
// Basic synchronization object

class CSyncObject
{
public:
	enum ESecurityAttribute 
	{
		saNull,
		saEveryone
	};
	
protected:
	SECURITY_ATTRIBUTES		m_sa;
	SECURITY_DESCRIPTOR		m_sd;
	
// Constructor
public:
	CSyncObject(LPCTSTR pstrName);

// Attributes
public:
	inline operator HANDLE() const { return m_hObject; }
	HANDLE  m_hObject;

// Operations
	virtual BOOL Lock(DWORD dwTimeout = INFINITE);
	virtual BOOL Unlock() = 0;
	virtual BOOL Unlock(LONG /* lCount */, LPLONG /* lpPrevCount=NULL */)
		{ return TRUE; }

// Implementation
public:
	virtual ~CSyncObject();
#ifdef _DEBUG
	#ifdef _UNICODE
		wstring m_strName;
	#else
		string m_strName;
	#endif
	//virtual void AssertValid() const;
	//virtual void Dump(CDumpContext& dc) const;
#endif
	friend class CSingleLock;
	friend class CMultiLock;

protected:
	BOOL CreateAccessibleSecurityDescriptor();
};

/////////////////////////////////////////////////////////////////////////////
// CSemaphore

class CSemaphore : public CSyncObject
{
// Constructor
public:
	CSemaphore(LONG lInitialCount = 1, LONG lMaxCount = 1,
		LPCTSTR pstrName=NULL, LPSECURITY_ATTRIBUTES lpsaAttributes = NULL);

// Implementation
public:
	virtual ~CSemaphore();
	inline virtual BOOL Unlock() { return Unlock(1, NULL); }
	virtual BOOL Unlock(LONG lCount, LPLONG lprevCount = NULL);
};

/////////////////////////////////////////////////////////////////////////////
// CMutex

class CMutex : public CSyncObject
{
// Constructor
public:
	CMutex(
		BOOL bInitiallyOwn = FALSE, 
		LPCTSTR lpszName = NULL,
		LPSECURITY_ATTRIBUTES lpsaAttribute = NULL,
		ESecurityAttribute SecurityAttribute = saNull);

// Implementation
public:
	virtual ~CMutex();
	BOOL Unlock();
};

/////////////////////////////////////////////////////////////////////////////
// CEvent

class CEvent : public CSyncObject
{
// Constructor
public:
	CEvent(BOOL bInitiallyOwn = FALSE, BOOL bManualReset = FALSE,
		LPCTSTR lpszNAme = NULL, LPSECURITY_ATTRIBUTES lpsaAttribute = NULL);

// Operations
public:
	inline BOOL SetEvent() { _ASSERTE(m_hObject != NULL); return ::SetEvent(m_hObject); }
	inline BOOL PulseEvent() { _ASSERTE(m_hObject != NULL); return ::PulseEvent(m_hObject); }
	inline BOOL ResetEvent() { _ASSERTE(m_hObject != NULL); return ::ResetEvent(m_hObject); }

	BOOL Unlock();

// Implementation
public:
	virtual ~CEvent();
};

/////////////////////////////////////////////////////////////////////////////
// CCriticalSection

class CCriticalSection : public CSyncObject
{
private:
//	static FARPROC volatile	m_fpInitializeCriticalSectionAndSpinCount;
//	static FARPROC volatile m_fpSetCriticalSectionSpinCountProc;

// Constructor
public:
	inline CCriticalSection(DWORD /* dwSpinCount */ = 0) : CSyncObject(NULL)
	{ 
		::InitializeCriticalSection(&m_sect);
	}
/*
	inline CCriticalSection(DWORD dwSpinCount = 0) : CSyncObject(NULL)
	{ 
		((BOOL (*)(LPCRITICAL_SECTION, DWORD))
			(*GetInitializeCriticalSectionAndSpinCountProc()))(&m_sect, dwSpinCount); 
	}
	inline DWORD SetSpinCount(DWORD dwSpinCount) 
	{ 
		return 
			((DWORD (*)(LPCRITICAL_SECTION, DWORD))
				(*GetSetCriticalSectionSpinCountProc()))(&m_sect, dwSpinCount); 
	}
*/

// Attributes
public:
	inline operator CRITICAL_SECTION*() { return (CRITICAL_SECTION*) &m_sect; }
	CRITICAL_SECTION m_sect;

// Operations
public:
	inline BOOL Unlock() { ::LeaveCriticalSection(&m_sect); return TRUE; }
	inline BOOL Lock() { ::EnterCriticalSection(&m_sect); return TRUE; }
	inline BOOL Lock(DWORD /* dwTimeout */) { return Lock(); }

// Implementation
public:
	inline virtual ~CCriticalSection() { ::DeleteCriticalSection(&m_sect); }

private:
//	static FARPROC GetInitializeCriticalSectionAndSpinCountProc();
//	static FARPROC GetSetCriticalSectionSpinCountProc();
};

/////////////////////////////////////////////////////////////////////////////
// CSingleLock

class CSingleLock
{
// Constructors
public:
	CSingleLock(CSyncObject* pObject, BOOL bInitialLock = FALSE);

// Operations
public:
	BOOL Lock(DWORD dwTimeOut = INFINITE);
	BOOL Unlock();
	BOOL Unlock(LONG lCount, LPLONG lPrevCount = NULL);
	inline BOOL IsLocked() { return m_bAcquired; }

// Implementation
public:
	inline ~CSingleLock() { Unlock(); }

protected:
	CSyncObject* m_pObject;
	HANDLE  m_hObject;
	BOOL    m_bAcquired;
};

/////////////////////////////////////////////////////////////////////////////
// CMultiLock

class CMultiLock
{
// Constructor
public:
	CMultiLock(CSyncObject* ppObjects[], DWORD dwCount, BOOL bInitialLock = FALSE);

// Operations
public:
	DWORD Lock(DWORD dwTimeOut = INFINITE, BOOL bWaitForAll = TRUE,
		DWORD dwWakeMask = 0);
	BOOL Unlock();
	BOOL Unlock(LONG lCount, LPLONG lPrevCount = NULL);
	inline BOOL IsLocked(DWORD dwObject) 
	{ 
		_ASSERTE(dwObject >= 0 && dwObject < m_dwCount);
		return m_bLockedArray[dwObject]; 
	}

// Implementation
public:
	~CMultiLock();

protected:
	HANDLE  m_hPreallocated[8];
	BOOL    m_bPreallocated[8];

	CSyncObject* const * m_ppObjectArray;
	HANDLE* m_pHandleArray;
	BOOL*   m_bLockedArray;
	DWORD   m_dwCount;
};

///////////////////////////////////////////////////////////////////////////////

class CSharedLock : public CObject
{
public:
	CSharedLock() {}
	virtual ~CSharedLock() {}

	virtual bool TryAcquireWriteLock() = 0;
	virtual void AcquireWriteLock(bool bAllowNewReadsToContinue = false) = 0;
	virtual void ReleaseWriteLock() = 0;
	virtual void AcquireReadLock() = 0;
	virtual void ReleaseReadLock() = 0;
};

///////////////////////////////////////////////////////////////////////////////

class CSharedLockSmallOperations : public CSharedLock
{
public:
	CSharedLockSmallOperations() :
		m_lLockType(LOCK_NONE),
		m_lReadLockCount(0),
		m_bAllowNewReadsToContinue(true)
	{
	}

	virtual ~CSharedLockSmallOperations()
	{
	}

	virtual bool TryAcquireWriteLock()
	{
		bool bLockAcquired = false;
#if _WIN32_WINNT >= 0x0500
		if (::InterlockedCompareExchange(&m_lLockType, LOCK_WRITE, LOCK_NONE) == LOCK_NONE)
#else
		if (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_WRITE, (PVOID)LOCK_NONE)) == LOCK_NONE)
#endif
		{
			bLockAcquired = true;
		}
		return bLockAcquired;
	};

	virtual void AcquireWriteLock(bool bAllowNewReadsToContinue = false)
	{
		m_bAllowNewReadsToContinue = bAllowNewReadsToContinue;
#if _WIN32_WINNT >= 0x0500
		while (::InterlockedCompareExchange(&m_lLockType, LOCK_WRITE, LOCK_NONE) != LOCK_NONE)
#else
		while (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_WRITE, (PVOID)LOCK_NONE)) != LOCK_NONE)
#endif
		{
			::Sleep(0);
		}
	}

	virtual void ReleaseWriteLock()
	{
		// release write lock
		VERIFY(::InterlockedExchange(&m_lLockType, LOCK_NONE) == LOCK_WRITE);
		// allow new reads to continue
		m_bAllowNewReadsToContinue = true;
	}

	virtual void AcquireReadLock()
	{
		// wait for any pending write locks if necessary
		while (m_bAllowNewReadsToContinue == false)
		{
			::Sleep(0);
		}
		
		// acquire the read lock
		::InterlockedIncrement(&m_lReadLockCount);
#if _WIN32_WINNT >= 0x0500
		while (::InterlockedCompareExchange(&m_lLockType, LOCK_READ, LOCK_NONE) == LOCK_WRITE &&
			   ::InterlockedCompareExchange(&m_lLockType, LOCK_READ, LOCK_READ) == LOCK_WRITE)
#else
		while (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_READ, (PVOID)LOCK_NONE)) == LOCK_WRITE &&
			   reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_READ, (PVOID)LOCK_READ)) == LOCK_WRITE)
#endif
		{
			::Sleep(0);
		}
	}

	virtual void ReleaseReadLock()
	{
		// decrement the number of read-locks
		if (::InterlockedDecrement(&m_lReadLockCount) == 0)
		{
			// last read-lock so release the lock.
			// ... Context switch could have occurred here (another read acquired by another thread),
			
			// ... Set m_lLockType to m_lReadLockCount which will either be 0/LOCK_NONE for no context switch
			// ... or  greater than 0 / LOCK_UNDEFINED (for a context switch).
			// ... The LOCK_UNDEFINED values will prevent all further read and write locks being acquired.
			do
			{
				::InterlockedExchange((long*)&m_lLockType, m_lReadLockCount);
				if (m_lLockType >= LOCK_UNDEFINED)
				{
					// reset the lock type and try again later (otherwise in AcquireReadLock, m_lReadLockCount may have been incremented, but acquiring the 
					// actual lock will block)
					m_lLockType = LOCK_READ;
					::Sleep(0);
				}
			} while (m_lLockType == LOCK_READ);
		}
	}

private:
	enum TeLockType
	{
		LOCK_WRITE = -2, // only one thread can hold a write lock, other will block
		LOCK_READ = -1,  // many threads can simulateously hold a read lock - an attempt to acquire for writing will block 
		LOCK_NONE = 0,   // no lock - can be acquired for either simulaneous reads, or, a single write

		LOCK_UNDEFINED   // m_lLockType may temporarily take on undefined values (i.e. greater than zero) when m_lReadLockCount is assigned to m_lLockType in ReleaseLock
	};

	long m_lLockType;
	long m_lReadLockCount; // number of read-locks

	bool m_bAllowNewReadsToContinue; // use for high priority writes (prevents new reads, and allows existing reads to finish)
};

///////////////////////////////////////////////////////////////////////////////

class CSharedLockLargeOperations : public CSharedLock
{
public:
	CSharedLockLargeOperations() :
		m_lLockType(LOCK_NONE),
		m_lReadLockCount(0)
	{
		m_hAllowNewReadsToContinueEvent = ::CreateEvent(NULL, TRUE, TRUE, NULL); // initially signalled, manually reset event
		m_hWriteNotInProgressEvent = ::CreateEvent(NULL, TRUE, TRUE, NULL); // initially signalled, manually reset event
		m_hReadsNotInProgressEvent = ::CreateEvent(NULL, TRUE, TRUE, NULL); // initially signalled, manually reset event
	}

	virtual ~CSharedLockLargeOperations()
	{
		VERIFY(::CloseHandle(m_hAllowNewReadsToContinueEvent));
		VERIFY(::CloseHandle(m_hWriteNotInProgressEvent));
		VERIFY(::CloseHandle(m_hReadsNotInProgressEvent));
	}

	virtual bool TryAcquireWriteLock()
	{
		bool bLockAcquired = false;
#if _WIN32_WINNT >= 0x0500
		if (::InterlockedCompareExchange(&m_lLockType, LOCK_WRITE, LOCK_NONE) == LOCK_NONE)
#else
		if (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_WRITE, (PVOID)LOCK_NONE)) == LOCK_NONE)
#endif
		{
			bLockAcquired = true;
			// signal that a write lock has been acquired
			VERIFY(::ResetEvent(m_hWriteNotInProgressEvent));
		}
		return bLockAcquired;
	};

	virtual void AcquireWriteLock(bool bAllowNewReadsToContinue = false)
	{
		if (bAllowNewReadsToContinue == false)
		{
			// signal to other threads that a pending write is in progress
			VERIFY(::ResetEvent(m_hAllowNewReadsToContinueEvent));
		}
#if _WIN32_WINNT >= 0x0500
		while (::InterlockedCompareExchange(&m_lLockType, LOCK_WRITE, LOCK_NONE) != LOCK_NONE)
#else
		while (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_WRITE, (PVOID)LOCK_NONE)) != LOCK_NONE)
#endif
		{
			::Sleep(0); // guard against tight loop due to context switch (locks acquired, but events not yet reset or overwritten as reset)

			// wait for reads and any other write not to be in progress
			HANDLE pHandles[] = {m_hWriteNotInProgressEvent, m_hReadsNotInProgressEvent};
			VERIFY(::WaitForMultipleObjects(2, pHandles, TRUE, INFINITE) == WAIT_OBJECT_0);
		}
		// signal that a write lock has been acquired
		VERIFY(::ResetEvent(m_hWriteNotInProgressEvent));
	}

	virtual void ReleaseWriteLock()
	{
		// release write lock
		VERIFY(::InterlockedExchange(&m_lLockType, LOCK_NONE) == LOCK_WRITE);
		// allow waiting threads to continue
		VERIFY(::SetEvent(m_hWriteNotInProgressEvent));
		// allow new reads to continue
		VERIFY(::SetEvent(m_hAllowNewReadsToContinueEvent));
	}

	virtual void AcquireReadLock()
	{
		// wait for any pending write locks if necessary
		VERIFY(::WaitForSingleObject(m_hAllowNewReadsToContinueEvent, INFINITE) == WAIT_OBJECT_0);
		
		// acquire the read lock
		::InterlockedIncrement(&m_lReadLockCount);
#if _WIN32_WINNT >= 0x0500
		while (::InterlockedCompareExchange(&m_lLockType, LOCK_READ, LOCK_NONE) == LOCK_WRITE &&
			   ::InterlockedCompareExchange(&m_lLockType, LOCK_READ, LOCK_READ) == LOCK_WRITE)
#else
		while (reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_READ, (PVOID)LOCK_NONE)) == LOCK_WRITE &&
			   reinterpret_cast<long>(::InterlockedCompareExchange((PVOID*)&m_lLockType, (PVOID)LOCK_READ, (PVOID)LOCK_READ)) == LOCK_WRITE)
#endif
		{
			::Sleep(0); // guard against tight loop due to context switch (write lock acquired, but event not yet reset or overwritten as reset)

			VERIFY(::WaitForSingleObject(m_hWriteNotInProgressEvent, INFINITE) == WAIT_OBJECT_0);
		}
		// signal that a read lock has been acquired
		VERIFY(::ResetEvent(m_hReadsNotInProgressEvent));
	}

	virtual void ReleaseReadLock()
	{
		// decrement the number of read-locks
		if (::InterlockedDecrement(&m_lReadLockCount) == 0)
		{
			// last read-lock so release the lock.
			// ... Context switch could have occurred here (another read acquired by another thread),
			
			// ... Set m_lLockType to m_lReadLockCount which will either be 0/LOCK_NONE for no context switch
			// ... or  greater than 0 / LOCK_UNDEFINED (for a context switch).
			// ... The LOCK_UNDEFINED values will prevent all further read and write locks being acquired.
			do
			{
				::InterlockedExchange((long*)&m_lLockType, m_lReadLockCount);
				if (m_lLockType >= LOCK_UNDEFINED)
				{
					// reset the lock type and try again later (otherwise in AcquireReadLock, m_lReadLockCount may have been incremented, but acquiring the 
					// actual lock will block)
					m_lLockType = LOCK_READ;
					::Sleep(0);
				}
			} while (m_lLockType == LOCK_READ);
			
			// allow waiting threads to continue
			VERIFY(::SetEvent(m_hReadsNotInProgressEvent));
		}
	}

private:
	enum TeLockType
	{
		LOCK_WRITE = -2, // only one thread can hold a write lock, other will block
		LOCK_READ = -1,  // many threads can simulateously hold a read lock - an attempt to acquire for writing will block 
		LOCK_NONE = 0,   // no lock - can be acquired for either simulaneous reads, or, a single write

		LOCK_UNDEFINED   // m_lLockType may temporarily take on undefined values (i.e. greater than zero) when m_lReadLockCount is assigned to m_lLockType in ReleaseLock
	};

	long m_lLockType;
	long m_lReadLockCount; // number of read-locks

	HANDLE m_hAllowNewReadsToContinueEvent; // use for high priority writes (prevents new reads, and allows existing reads to finish)
	HANDLE m_hWriteNotInProgressEvent;
	HANDLE m_hReadsNotInProgressEvent;
};

///////////////////////////////////////////////////////////////////////////////

class CReadLock : public CObject
{
public:
	CReadLock(CSharedLock& rSharedLock) :
		m_bLocked(false),
		m_rSharedLock(rSharedLock)
	{
		Lock();
	}
	~CReadLock()
	{
		UnLock();
	}

	bool Lock()
	{
		if (!m_bLocked)
		{
			m_rSharedLock.AcquireReadLock();
			m_bLocked = true;
		}
		return true;
	}
	bool UnLock()
	{
		if (m_bLocked)
		{
			m_rSharedLock.ReleaseReadLock();
			m_bLocked = false;
		}
		return true;
	}


private:
	bool m_bLocked;
	CSharedLock& m_rSharedLock;
};

///////////////////////////////////////////////////////////////////////////////

class CWriteLock : public CObject
{
public:
	CWriteLock(CSharedLock& rSharedLock, bool bLock = true, bool bAllowNewReadsToContinue = false) :
		m_bLocked(false),
		m_rSharedLock(rSharedLock)
	{
		if (bLock)
		{
			Lock(bAllowNewReadsToContinue);
		}
	}
	~CWriteLock()
	{
		UnLock();
	}

	bool TryLock()
	{
		if (!m_bLocked)
		{
			m_bLocked = m_rSharedLock.TryAcquireWriteLock();
		}
		return m_bLocked;
	}
	bool Lock(bool bAllowNewReadsToContinue = false)
	{
		if (!m_bLocked)
		{
			m_rSharedLock.AcquireWriteLock(bAllowNewReadsToContinue);
			m_bLocked = true;
		}
		return true;
	}
	bool UnLock()
	{
		if (m_bLocked)
		{
			m_rSharedLock.ReleaseWriteLock();
			m_bLocked = false;
		}
		return true;
	}

private:
	bool m_bLocked;
	CSharedLock& m_rSharedLock;
};

///////////////////////////////////////////////////////////////////////////////

}

#endif	// MUTEX_H
