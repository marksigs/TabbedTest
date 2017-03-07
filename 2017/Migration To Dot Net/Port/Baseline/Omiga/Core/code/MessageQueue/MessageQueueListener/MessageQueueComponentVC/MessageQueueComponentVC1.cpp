///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MessageQueueComponentVC.h"
#include "MessageQueueComponentVC1.h"
#include "VC1ResponseDialog1.h"

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA CLogIn LogIn(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THREADDATA_INOUT CLogInOut LogInOut(pThreadData->m_rThreadPoolManagerMSMQ1, MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT); LogInOut.AnyThreadInitialize

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC1


HRESULT CMessageQueueComponentVC1::Activate()
{
	HRESULT hr = GetObjectContext(&m_spObjectContext);
	if (SUCCEEDED(hr))
		return S_OK;
	return hr;
} 

BOOL CMessageQueueComponentVC1::CanBePooled()
{
	return FALSE;
} 

void CMessageQueueComponentVC1::Deactivate()
{
	m_spObjectContext.Release();
} 

STDMETHODIMP CMessageQueueComponentVC1::OnMessage(BSTR in_xml, long* plMESSQ_RESP)
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueComponentVC1::OnMessage\n"));
	CVC1ResponseDialog1 VC1ResponseDialog1(in_xml);
	VC1ResponseDialog1.DoModal();
	*plMESSQ_RESP = VC1ResponseDialog1.m_lMESSQ_RESP;

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

