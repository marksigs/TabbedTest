///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC2.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      15/11/00    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MessageQueueComponentVC.h"
#include "MessageQueueComponentVC2.h"
#include "VC2ResponseDialog1.h"

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA CLogIn LogIn(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA_INOUT CLogInOut LogInOut(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC2

HRESULT CMessageQueueComponentVC2::Activate()
{
	HRESULT hr = GetObjectContext(&m_spObjectContext);
	if (SUCCEEDED(hr))
		return S_OK;
	return hr;
} 

BOOL CMessageQueueComponentVC2::CanBePooled()
{
	return FALSE;
} 

void CMessageQueueComponentVC2::Deactivate()
{
	m_spObjectContext.Release();
} 

STDMETHODIMP CMessageQueueComponentVC2::OnMessage(BSTR in_xmlConfig, BSTR in_xmlData, long* plMESSQ_RESP)
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueComponentVC2::OnMessage\n"));
	CVC2ResponseDialog1 VC2ResponseDialog1(in_xmlConfig, in_xmlData);
	VC2ResponseDialog1.DoModal();
	*plMESSQ_RESP = VC2ResponseDialog1.m_lMESSQ_RESP;

	if (m_spObjectContext)
	{
		switch (*plMESSQ_RESP)
		{
			case MESSQ_RESP_SUCCESS:
				m_spObjectContext->SetComplete();
				break;
			case MESSQ_RESP_RETRY_NOW:
				m_spObjectContext->SetAbort();
				break;
			case MESSQ_RESP_RETRY_LATER:
				m_spObjectContext->SetAbort();
				break;
			case MESSQ_RESP_STALL_COMPONENT:
				m_spObjectContext->SetAbort();
				break;
			case MESSQ_RESP_STALL_QUEUE:
				m_spObjectContext->SetAbort();
				break;
			default :
				break;
		}
	}

	return S_OK;
}

