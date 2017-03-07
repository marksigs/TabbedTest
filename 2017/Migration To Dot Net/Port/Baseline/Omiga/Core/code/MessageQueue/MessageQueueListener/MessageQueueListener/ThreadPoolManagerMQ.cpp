///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerMQ.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      06/09/00    Initial version
//	LD		06/04/01	SYS2248 - Add performance counters
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "mutex.h"
using namespace namespaceMutex;
#include "ThreadPoolManagerMQ.h"


CThreadPoolManagerMQ::CThreadPoolManagerMQ()
{
}

CThreadPoolManagerMQ::~CThreadPoolManagerMQ()
{

}

STDMETHODIMP CThreadPoolManagerMQ::get_ComponentsStalled(VARIANT *pVal)
{
	HRESULT hr = S_OK;

	CReadLock lckStalledComponents(m_sharedlockStalledComponents);

	VariantInit(pVal);
	V_VT(pVal) = VT_ARRAY | VT_BSTR;
	SAFEARRAYBOUND  rgsabound[1];
	rgsabound[0].cElements = m_listStalledComponents.size();
	rgsabound[0].lLbound = 0;
	V_ARRAY(pVal) = SafeArrayCreate(VT_BSTR, 1, rgsabound);
	if(V_ARRAY(pVal) != NULL)
	{	
	    BSTR* pbstr = NULL;
	    SafeArrayAccessData(V_ARRAY(pVal), reinterpret_cast<void**>(&pbstr));
		for (TIteratorStalledComponents iterator = m_listStalledComponents.begin(); iterator != m_listStalledComponents.end(); iterator++)
		{
			*pbstr++ = (*iterator).copy();
		}
	    SafeArrayUnaccessData(V_ARRAY(pVal));
	}
	else
	{
		hr = E_FAIL;
	}

	return hr;
}


STDMETHODIMP CThreadPoolManagerMQ::put_AddStalledComponents(VARIANT newVal)
{
	HRESULT hr = S_OK;

//	CWriteLock lckStalledComponents(m_sharedlockStalledComponents);

	LONG lLBound = 0;
	LONG lUBound = 0;
	BSTR* pbstr = NULL;

	if (V_VT(&newVal) != (VT_ARRAY | VT_BSTR) || // array of BSTRs
		SafeArrayGetDim(V_ARRAY(&newVal)) != 1 || // one dimensional 
		FAILED(hr = SafeArrayGetLBound(V_ARRAY(&newVal), 1, &lLBound)) ||
		FAILED(SafeArrayGetUBound(V_ARRAY(&newVal), 1, &lUBound)) ||
		FAILED(SafeArrayAccessData(V_ARRAY(&newVal), reinterpret_cast<void**>(&pbstr))))
	{
		hr = E_FAIL;
	}
	else
	{

		LONG cElements = lUBound-lLBound+1;
		for (int i = 0; i < cElements; i++)
		{
			if (pbstr[i] != NULL)
			{
				AddStallComponent(pbstr[i]);
			}
		}
	}
	
	SafeArrayUnaccessData(V_ARRAY(&newVal));
	
	return hr;
}

STDMETHODIMP CThreadPoolManagerMQ::put_RestartComponents(VARIANT newVal)
{
	HRESULT hr = S_OK;

//	CWriteLock lckStalledComponents(m_sharedlockStalledComponents);

	
	LONG lLBound = 0;
	LONG lUBound = 0;
	BSTR* pbstr = NULL;

	if (V_VT(&newVal) != (VT_ARRAY | VT_BSTR) || // array of BSTRs
		SafeArrayGetDim(V_ARRAY(&newVal)) != 1 || // one dimensional 
		FAILED(hr = SafeArrayGetLBound(V_ARRAY(&newVal), 1, &lLBound)) ||
		FAILED(SafeArrayGetUBound(V_ARRAY(&newVal), 1, &lUBound)) ||
		FAILED(SafeArrayAccessData(V_ARRAY(&newVal), reinterpret_cast<void**>(&pbstr))))
	{
		hr = E_FAIL;
	}
	else
	{

		LONG cElements = lUBound-lLBound+1;
		for (int i = 0; i < cElements; i++)
		{
			if (pbstr[i] != NULL)
			{
				RemoveStallComponent(pbstr[i]);
			}
		}
	}
	
	SafeArrayUnaccessData(V_ARRAY(&newVal));
		
	
	return hr;
}



void CThreadPoolManagerMQ::AddStallComponent(_bstr_t bstrProgID, BOOL* pbAlreadyStalled)
{
	CWriteLock lckStalledComponents(m_sharedlockStalledComponents);
	BOOL bAlreadyStalled = FALSE;
	for (TIteratorStalledComponents iterator = m_listStalledComponents.begin(); iterator != m_listStalledComponents.end(); iterator++)
	{
		if (*iterator == bstrProgID)
		{
			bAlreadyStalled = TRUE;
			break;
		}
	}
	if (!bAlreadyStalled)
	{
		m_listStalledComponents.push_back(bstrProgID);
	}
	if (pbAlreadyStalled != NULL)
	{
		*pbAlreadyStalled = bAlreadyStalled;
	}
}

void CThreadPoolManagerMQ::RemoveStallComponent(_bstr_t bstrProgID)
{
	CWriteLock lckStalledComponents(m_sharedlockStalledComponents);
	for (TIteratorStalledComponents iterator = m_listStalledComponents.begin(); iterator != m_listStalledComponents.end(); iterator++)
	{
		if (*iterator == bstrProgID)
		{
			m_listStalledComponents.erase(iterator);
			break;
		}
	}
}

bool CThreadPoolManagerMQ::IsComponentStalled(_bstr_t bstrProgID)
{
	CReadLock lckStalledComponents(m_sharedlockStalledComponents);
	bool bStalled = false;
	for (TIteratorStalledComponents iterator = m_listStalledComponents.begin(); iterator != m_listStalledComponents.end(); iterator++)
	{
		if (*iterator == bstrProgID)
		{
			bStalled = true;
			break;
		}
	}
	return bStalled;
}

bool CThreadPoolManagerMQ::IsAnyComponentStalled()
{
	CReadLock lckStalledComponents(m_sharedlockStalledComponents);
	int nStalledComponents = m_listStalledComponents.size();
	return nStalledComponents > 0;
}

int CThreadPoolManagerMQ::GetnDifferentComponentsStalled()
{
	CReadLock lckStalledComponents(m_sharedlockStalledComponents);
	int nStalledComponents = m_listStalledComponents.size();
	return nStalledComponents;
}
