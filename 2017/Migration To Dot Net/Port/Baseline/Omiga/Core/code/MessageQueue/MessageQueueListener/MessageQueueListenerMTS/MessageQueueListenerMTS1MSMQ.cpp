///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerMTS1MSMQ.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//	LD		06/04/01	SYS2248 - Add performance counters
//	LD		09/04/01	SYS2249 - Corrections for two listeners on the same queue
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//						Use m_spObjectContext->CreateInstance
//  LD		20/06/02	SYS4933 - Add originating thread id
//  LD		19/03/03	Updates to support queuenames specified as a format name
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <stdio.h>
#include "MessageQueueListenerMTS.h"
#include "MessageQueueListenerMTS1.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerMTS1

#import "..\MessageQueueComponentVC\MessageQueueComponentVC.tlb" no_namespace

const long s_lMSMQPeekTimeOut = 1000L; // in ms
const long s_lMSMQReceiveTimeOut = INFINITE; // in ms

#define LOG CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1MSMQ, dwOriginatingThreadId); LogIn.Initialize
#define LOG_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1MSMQ, dwOriginatingThreadId); LogInOut.Initialize

///////////////////////////////////////////////////////////////////////////////

STDMETHODIMP CMessageQueueListenerMTS1::MSMQReceiveExecute(DWORD dwOriginatingThreadId, BSTR bstrConfig, BSTR bstrQueueName, BSTR bstrFormatName, BSTR bstrMoveQueueName, BSTR bstrMoveFormatName, VARIANT vMsgId, long lMessagesToScan, BSTR* pbstrErrorMessage, long *plMESSQ_RESP)
{
	LOG_INOUT(_T("%s MSMQReceiveExecute\n"), (LPWSTR)bstrQueueName);
	const LPCTSTR pszFunctionName = _T("MSMQReceiveExecute");
	
	// assume an error
	long lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
	if (m_spObjectContext)
	{
		m_spObjectContext->SetAbort();
	}
	ResetErrorMessage();

	enum
	{
		eNULL,
		eOpenQueue,
		eReceive,
		eLabel,
		eBody,
		eCloseQueue,
		eGetComponentInterface,
		eCallComponent,
	} eAction = eNULL;
	
	_bstr_t bstrLabel;
	_variant_t vBody;
	bool bReceivedMessage = false;
	bool bInTransaction = false;

	try
	{
		bInTransaction = m_spObjectContext && m_spObjectContext->IsInTransaction();

		// open the queue
		eAction = eOpenQueue;
		IMSMQQueueInfoPtr MSMQQueueInfoPtr(__uuidof(MSMQQueueInfo));
		MSMQQueueInfoPtr->put_FormatName(bstrFormatName);
		IMSMQQueuePtr MSMQQueuePtr = MSMQQueueInfoPtr->Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE);
		
		// process the message if found (identified by the Id)
		eAction = eReceive;
		IMSMQMessagePtr MSMQMessagePtr = HelperMSMQReceiveMessage(MSMQQueuePtr, vMsgId, lMessagesToScan);
		if (MSMQMessagePtr)
		{
			bReceivedMessage = true;

			// extract the label
			eAction = eLabel;
			bstrLabel = MSMQMessagePtr->GetLabel();
			
			// get the body
			eAction = eBody;
			vBody = MSMQMessagePtr->GetBody();

			// close the queue
			eAction = eCloseQueue;
			MSMQQueuePtr->Close();

			// call the component
			eAction = eGetComponentInterface;
			CLSID clsid;
			if (SUCCEEDED(CLSIDFromProgID(bstrLabel, &clsid)))
			{
				LOG_INOUT(_T("%s MSMQReceiveExecute - CreateComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrLabel);
				IUnknownPtr UnknownPtr(clsid);
				if (UnknownPtr)
				{
					lMESSQ_RESP = MESSQ_RESP_RETRY_NOW;
					// attempt to create MessageQueueComponentVC2Ptr
					IMessageQueueComponentVC2Ptr MessageQueueComponentVC2Ptr;
					if (m_spObjectContext)
					{
						m_spObjectContext->CreateInstance(clsid, __uuidof(IMessageQueueComponentVC2), reinterpret_cast<void**>(&MessageQueueComponentVC2Ptr));
					}
					else
					{
						MessageQueueComponentVC2Ptr.CreateInstance(clsid, NULL);
					}
					if (MessageQueueComponentVC2Ptr)
					{
						//  MessageQueueComponentVC2Ptr created
						// if within a transaction this component needs to be re-created
						// if not transactional then do the retries here (as the message read/deleted will not be rolled back)
						int nMaxTries = bInTransaction ? 1 : 3;
						int nTry = 0;
						while (nTry < nMaxTries && lMESSQ_RESP == MESSQ_RESP_RETRY_NOW)
						{
							eAction = eCallComponent;
							LOG_INOUT(_T("%s MSMQReceiveExecute - CallComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrLabel);
							lMESSQ_RESP = MessageQueueComponentVC2Ptr->OnMessage(bstrConfig, _bstr_t(vBody));
							nTry++;
						}
						if (m_spObjectContext)
						{
							if (lMESSQ_RESP == MESSQ_RESP_SUCCESS)
							{
								m_spObjectContext->SetComplete();
							}
						}
					}
					else
					{
						// attempt to create MessageQueueComponentVC1Ptr
						IMessageQueueComponentVC1Ptr MessageQueueComponentVC1Ptr;
						if (m_spObjectContext)
						{
							m_spObjectContext->CreateInstance(clsid, __uuidof(IMessageQueueComponentVC1), reinterpret_cast<void**>(&MessageQueueComponentVC1Ptr));
						}
						else
						{
							MessageQueueComponentVC1Ptr.CreateInstance(clsid, NULL);
						}
						if (MessageQueueComponentVC1Ptr)
						{
							//  MessageQueueComponentVC1Ptr created
							// if within a transaction this component needs to be re-created
							// if not transactional then do the retries here (as the message read/deleted will not be rolled back)
							int nMaxTries = bInTransaction ? 1 : 3;
							int nTry = 0;
							while (nTry < nMaxTries && lMESSQ_RESP == MESSQ_RESP_RETRY_NOW)
							{
								eAction = eCallComponent;
								LOG_INOUT(_T("%s MSMQReceiveExecute - CallComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrLabel);
								lMESSQ_RESP = MessageQueueComponentVC1Ptr->OnMessage(_bstr_t(vBody));
								nTry++;
							}
							if (m_spObjectContext)
							{
								if (lMESSQ_RESP == MESSQ_RESP_SUCCESS)
								{
									m_spObjectContext->SetComplete();
								}
							}
						}
						else
						{
							AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Error querying for known interface. Progid (%ls).  Stalling Component"), (LPWSTR)bstrLabel);
							lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
						}
					}
				}
				else
				{
					AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Error querying IUnknown for Progid (%ls).  Stalling Component"), (LPWSTR)bstrLabel);
					lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
				}
			}
			else
			{
				AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Error converting Progid (%ls) to CLSID.  Stalling Component"), (LPWSTR)bstrLabel);
				lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
			}
		}
		else
		{
			// Message must have been stolen by another process
			// close the queue
			eAction = eCloseQueue;
			MSMQQueuePtr->Close();
			lMESSQ_RESP = MESSQ_RESP_HANDLED_STOLEN;
		}
	}
	catch(_com_error comerr)
	{
		lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
		switch (eAction)
		{
			case eOpenQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error opening queue"));
				break;
			case eReceive:
				if (comerr.Error() == MQ_ERROR_MESSAGE_ALREADY_RECEIVED)
				{
					lMESSQ_RESP = MESSQ_RESP_HANDLED_STOLEN;
				}
				else
				{
					AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error receiving messsage"));
				}
				break;
			case eLabel:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error getting label from message"));
				break;
			case eBody:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error getting body from message"));
				break;
			case eCloseQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error closing queue"));
				break;
			case eGetComponentInterface:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error creating component or querying interface"));
				break;
			case eCallComponent:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error calling component"));
				break;
			default :
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Unknown error"));
				break;
		}
	}
	catch(...)
	{
		lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
		AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Unknown error"));
	}

	// corrections for transactional behaviour result in this component being called again in another transaction

	// correction for non-transactional behaviour must be done here, otherwise the message is lost
	if (!bInTransaction)
	{
		try
		{
			switch (lMESSQ_RESP)
			{
				case MESSQ_RESP_SUCCESS:
					break;
				case MESSQ_RESP_RETRY_NOW: // (immediate retries already done)
					// place the message back on the queue if it has been received (i.e. removed)
					if (bReceivedMessage)
					{
						HelperMSMQSendMessage(bstrQueueName, bstrFormatName, bstrLabel, vBody);
						lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_NOW; // modify response
					}
					break;
				case MESSQ_RESP_RETRY_LATER:
					// place the message back on the queue if it has been received (i.e. removed)
					if (bReceivedMessage)
					{
						HelperMSMQSendMessage(bstrQueueName, bstrFormatName, bstrLabel, vBody);
						lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_LATER; // modify response
					}
					break;
				case MESSQ_RESP_RETRY_MOVE_MESSAGE:
				{
					_ASSERTE(bReceivedMessage);
					HelperMSMQSendMessage(bstrMoveQueueName, bstrMoveFormatName, bstrLabel, vBody);
					lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_MOVE_MESSAGE; // modify response
					break;
				}
				case MESSQ_RESP_STALL_COMPONENT:
					_ASSERTE(bReceivedMessage);
					HelperMSMQSendMessage(bstrQueueName, bstrFormatName, bstrLabel, vBody);
					break;
				case MESSQ_RESP_STALL_QUEUE:
					_ASSERTE(bReceivedMessage);
					HelperMSMQSendMessage(bstrQueueName, bstrFormatName, bstrLabel, vBody);
					break;
				case MESSQ_RESP_HANDLED_RETRY_NOW:
					_ASSERTE(0); // should not reach here
					break;
				case MESSQ_RESP_HANDLED_RETRY_LATER:
					_ASSERTE(0); // should not reach here
					break;
				case MESSQ_RESP_HANDLED_RETRY_MOVE_MESSAGE:
					_ASSERTE(0); // should not reach here
					break;
				case MESSQ_RESP_HANDLED_STOLEN:
					break;
				case MESSQ_RESP_HANDLED_ERROR:
					break;
				default :
					_ASSERTE(0); // should not reach here 
					break;
			}
		}
		catch(_com_error comerr)
		{
			AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Unknown error"));
		}	
		catch(...)
		{
			AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Unknown error"));
		}
	}


	*pbstrErrorMessage = AllocbstrErrorMessage();
	*plMESSQ_RESP = lMESSQ_RESP;
	return S_OK;
}


STDMETHODIMP CMessageQueueListenerMTS1::MSMQMoveMessage(DWORD dwOriginatingThreadId, BSTR bstrQueueName, BSTR bstrFormatName, BSTR bstrMoveQueueName, BSTR bstrMoveFormatName, VARIANT vMsgId, long lMessagesToScan, BSTR *pbstrErrorMessage)
{
	LOG_INOUT(_T("%s MSMQMoveMessage\n"), (LPWSTR)bstrQueueName);

	const LPCTSTR pszFunctionName = _T("MSMQMoveMessage");
	
	// move the message from bstrQueueName to (bstrQueueName + 'x')
	
	// assume a failure
	if (m_spObjectContext)
	{
		m_spObjectContext->SetAbort();
	}
	ResetErrorMessage();

	enum
	{
		eNULL,
		eSrcOpenQueue,
		eSrcReceive,
		eMoveMessage,
		eSrcCloseQueue,
	} eAction = eNULL;
	
	try
	{
		// open the source queue
		eAction = eSrcOpenQueue;
		IMSMQQueueInfoPtr MSMQQueueInfoPtrSrc(__uuidof(MSMQQueueInfo));
		MSMQQueueInfoPtrSrc->put_FormatName(bstrFormatName);
		IMSMQQueuePtr MSMQQueuePtrSrc = MSMQQueueInfoPtrSrc->Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE);
		
		// open the source message if found (identified by the Id)
		eAction = eSrcReceive;
		IMSMQMessagePtr MSMQMessagePtrSrc = HelperMSMQReceiveMessage(MSMQQueuePtrSrc, vMsgId, lMessagesToScan);
		if (MSMQMessagePtrSrc)
		{
			// move the message
			eAction = eMoveMessage;
			HelperMSMQSendMessage(bstrMoveQueueName, bstrMoveFormatName, MSMQMessagePtrSrc->GetLabel(), MSMQMessagePtrSrc->GetBody());
			
			// close the source queue
			eAction = eSrcCloseQueue;
			MSMQQueuePtrSrc->Close();

			if (m_spObjectContext)
			{
				m_spObjectContext->SetComplete();
			}
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eSrcOpenQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error opening source queue"));
				break;
			case eSrcReceive:
				if (comerr.Error() != MQ_ERROR_MESSAGE_ALREADY_RECEIVED)
				{
					AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error receiving messsage"));
				}
				break;
			case eMoveMessage:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error moving messsage"));
				break;
			case eSrcCloseQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Error closing source queue"));
				break;
			default :
				AddToErrorMessage(pszFunctionName, comerr, _T("MSMQMTS1 - Unknown error"));
				break;
		}
	}
	catch(...)
	{
		AddToErrorMessage(pszFunctionName, _T("MSMQMTS1 - Unknown error"));
	}

	*pbstrErrorMessage = AllocbstrErrorMessage();
	return S_OK;
}

IMSMQMessagePtr CMessageQueueListenerMTS1::HelperMSMQReceiveMessage(IMSMQQueuePtr MSMQQueuePtr, VARIANT vMsgId, long lMessagesToScan)
{
	IMSMQMessagePtr MSMQMessagePtr;
	bool bFound = HelperMSMQFindMessage(MSMQQueuePtr, vMsgId, lMessagesToScan);
	if (bFound)
	{
		// receive the message
		long lTransactionType = MQ_NO_TRANSACTION;
		if (m_spObjectContext && m_spObjectContext->IsInTransaction())
		{
			lTransactionType = MQ_MTS_TRANSACTION;
		}
		MSMQMessagePtr = MSMQQueuePtr->ReceiveCurrent(&_variant_t(lTransactionType), &_variant_t(false), &_variant_t(true), &_variant_t(s_lMSMQReceiveTimeOut));
	}
	else
	{
		// Message must have been stolen by another process
	}
	return MSMQMessagePtr;
}

bool CMessageQueueListenerMTS1::HelperMSMQFindMessage(IMSMQQueuePtr MSMQQueuePtr, VARIANT vMsgId, long lMessagesToScan)
{
	bool bFound = false;
	long lMessagesScanned = 0;
	HRESULT hr = E_FAIL;

	// extract the 20 bytes identifier from the message passes as a parameter
	const int nBytesMax = 20;
	unsigned char MsgId20bytes[20];
	bool bExtractedBytes = false;
	if (vMsgId.vt == (VT_ARRAY | VT_UI1))
	{
		SAFEARRAY* psa = ((VARIANT)vMsgId).parray;

		unsigned char* p20Bytes = NULL;
		HRESULT hr = ::SafeArrayAccessData(psa, (void HUGEP**)&p20Bytes);
		if (SUCCEEDED(hr))
		{
			if (psa->rgsabound->cElements == nBytesMax)
			{
				memcpy(MsgId20bytes, p20Bytes, nBytesMax);
				bExtractedBytes =  true;
			}
		}
		::SafeArrayUnaccessData(psa);
	}

	if (bExtractedBytes)
	{
		IMSMQMessagePtr MSMQMessagePtr;

		_variant_t vIdMessage;
		for (MSMQMessagePtr = MSMQQueuePtr->PeekCurrent(&_variant_t(false), &_variant_t(false), &_variant_t(s_lMSMQPeekTimeOut));
			 MSMQMessagePtr != 0 && lMessagesScanned < lMessagesToScan;
			 MSMQMessagePtr = MSMQQueuePtr->PeekNext(&_variant_t(false), &_variant_t(false), &_variant_t(s_lMSMQPeekTimeOut)))
		{
			vIdMessage = MSMQMessagePtr->GetId();
			// compare message ids
			if (vIdMessage.vt == (VT_ARRAY | VT_UI1))
			{
				SAFEARRAY* psa = ((VARIANT)vIdMessage).parray;

				unsigned char* p20Bytes = NULL;
				HRESULT hr = ::SafeArrayAccessData(psa, (void HUGEP**)&p20Bytes);
				if (SUCCEEDED(hr))
				{
					if (psa->rgsabound->cElements == nBytesMax)
					{
						if (memcmp(MsgId20bytes, p20Bytes, nBytesMax) == 0)
						{
							bFound = true;
						}
					}
					::SafeArrayUnaccessData(psa);
				}
				if (bFound)
				{
					break;
				}
			}
			lMessagesScanned++;
		}
	}
	return bFound;
}

void CMessageQueueListenerMTS1::HelperMSMQSendMessage(BSTR bstrQueueName, BSTR bstrFormatName, BSTR bstrLabel, VARIANT vBody)
{
	IMSMQQueueInfoPtr MSMQQueueInfoPtr(__uuidof(MSMQQueueInfo));
	MSMQQueueInfoPtr->put_FormatName(bstrFormatName);
	IMSMQQueuePtr MSMQQueuePtr = MSMQQueueInfoPtr->Open(MQ_SEND_ACCESS, MQ_DENY_NONE);
	HelperMSMQSendMessage(MSMQQueuePtr, bstrLabel, vBody);
	MSMQQueuePtr->Close();
}

void CMessageQueueListenerMTS1::HelperMSMQSendMessage(IMSMQQueuePtr MSMQQueuePtr, BSTR bstrLabel, VARIANT vBody)
{
	IMSMQMessagePtr MSMQMessagePtr(__uuidof(MSMQMessage));
	MSMQMessagePtr->PutLabel(bstrLabel);
	MSMQMessagePtr->PutBody(vBody);
	HelperMSMQSendMessage(MSMQQueuePtr, MSMQMessagePtr);
}

void CMessageQueueListenerMTS1::HelperMSMQSendMessage(IMSMQQueuePtr MSMQQueuePtr, IMSMQMessagePtr MSMQMessagePtr)
{
	long lTransactionType = MQ_SINGLE_MESSAGE;
	if (m_spObjectContext && m_spObjectContext->IsInTransaction())
	{
		lTransactionType = MQ_MTS_TRANSACTION;
	}
	MSMQMessagePtr->Send(MSMQQueuePtr, &_variant_t(lTransactionType));
}

