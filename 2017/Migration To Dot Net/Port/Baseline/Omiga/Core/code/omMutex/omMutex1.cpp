/* <VERSION  CORELABEL="" LABEL="R15" DATE="10/10/2003 15:52:58" VERSION="254" PATH="$/CodeCoreCust/4UATCust/Code/Synchronisation/omMutex/omMutex1.cpp"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			omMutex1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      13/03/03    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OmMutex.h"
#include "omMutex1.h"

/////////////////////////////////////////////////////////////////////////////
// ComMutex1

HRESULT ComMutex1::FinalConstruct()
{
	return CComObjectRootEx<CComMultiThreadModel>::FinalConstruct();
}

void ComMutex1::FinalRelease()
{
	ReleaseMutex();
	CComObjectRootEx<CComMultiThreadModel>::FinalRelease();
}

STDMETHODIMP ComMutex1::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IomMutex1
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP ComMutex1::AcquireMutexWithTimeout(BSTR bstrMutexName, VARIANT_BOOL bHighPrioirty, LONG dwMilliseconds)
{
	HRESULT hr = E_FAIL;
	try
	{
		ReleaseMutex();

		m_pMutex = new namespaceMutex::CMutex(
			FALSE,
			bstrMutexName,
			NULL,
			namespaceMutex::CSyncObject::saEveryone);

		m_pSingleLock = new namespaceMutex::CSingleLock(m_pMutex, FALSE);

		HANDLE hThread = GetCurrentThread();
		DWORD dwThreadPriorityOriginal = GetThreadPriority(hThread);
		if (bHighPrioirty == VARIANT_TRUE)
		{
			SetThreadPriority(hThread, THREAD_PRIORITY_ABOVE_NORMAL);
		}

		if (m_pSingleLock->Lock(dwMilliseconds))
		{
			hr = S_OK;
		}
		else
		{
			Error(L"Failed to acquire mutex");
		}

		if (bHighPrioirty == VARIANT_TRUE)
		{
			// revert back to orinal priority
			SetThreadPriority(hThread, dwThreadPriorityOriginal);
		}
	
	}
	catch(...)
	{
		Error(L"Unexpected exception caught");
	}

	return hr;
}

STDMETHODIMP ComMutex1::AcquireMutex(BSTR bstrMutexName, VARIANT_BOOL bHighPrioirty)
{
	return AcquireMutexWithTimeout(bstrMutexName, bHighPrioirty, INFINITE);
}

STDMETHODIMP ComMutex1::ReleaseMutex()
{
	HRESULT hr = E_FAIL;
	try
	{
		if (m_pSingleLock)
		{
			m_pSingleLock->Unlock();
			delete m_pSingleLock;
			m_pSingleLock = NULL;
		}
		if (m_pMutex)
		{
			delete m_pMutex;
			m_pMutex = NULL;
		}
		hr = S_OK;
	}
	catch(...)
	{
		Error(L"Unexpected exception caught");
	}

	return hr;
}

