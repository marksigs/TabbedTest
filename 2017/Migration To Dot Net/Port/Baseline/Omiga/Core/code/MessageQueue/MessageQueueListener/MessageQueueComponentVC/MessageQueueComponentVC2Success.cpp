///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC2Success.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      10/04/01    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MessageQueueComponentVC.h"
#include "MessageQueueComponentVC2Success.h"
#import "..\MessageQueueListenerLOG\MessageQueueListenerLOG.tlb"

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA CLogIn LogIn(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA_INOUT CLogInOut LogInOut(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC2Success

HRESULT CMessageQueueComponentVC2Success::Activate()
{
	HRESULT hr = GetObjectContext(&m_spObjectContext);
	if (SUCCEEDED(hr))
		return S_OK;
	return hr;
} 

BOOL CMessageQueueComponentVC2Success::CanBePooled()
{
	return FALSE;
} 

void CMessageQueueComponentVC2Success::Deactivate()
{
	m_spObjectContext.Release();
} 

STDMETHODIMP CMessageQueueComponentVC2Success::OnMessage(BSTR in_xmlConfig, BSTR in_xmlData, long* plMESSQ_RESP)
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueComponentVC2Success::OnMessage\n"));
	if (m_spObjectContext)
	{
		m_spObjectContext->SetComplete();
	}
	*plMESSQ_RESP = MESSQ_RESP_SUCCESS;

	return S_OK;
}

