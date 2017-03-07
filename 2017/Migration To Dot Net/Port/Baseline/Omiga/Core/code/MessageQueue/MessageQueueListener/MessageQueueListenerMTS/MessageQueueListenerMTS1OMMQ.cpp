///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerMTS1OMMQ.cpp
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
//	LD		10/09/01	SYS2705 - Support for SQL Server added
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		09/02/02	SYS4414	- Add Oracle's OLEDB provider
//								- Use m_spObjectContext->CreateInstance
//								- If XML in table MQLOMMQ is empty then 
//								  look in MQLOMQLARGE
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
//  LD		06/02/03	OPC0006008 - Explicity close connections when an excpetion
//						occurs
//	LD		13/02/03	Clarify error messages
//  RF		24/02/04	BMIDS727 - Different queues can use different tables
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <stdio.h>
#include "MessageQueueListenerMTS.h"
#include "MessageQueueListenerMTS1.h"
#include "DbHelperOMMQ.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerMTS1

#import "..\MessageQueueComponentVC\MessageQueueComponentVC.tlb" no_namespace

#define LOG CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1OMMQ, dwOriginatingThreadId); LogIn.Initialize
#define LOG_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1OMMQ, dwOriginatingThreadId); LogInOut.Initialize

/////////////////////////	//////////////////////////////////////////////////////

STDMETHODIMP CMessageQueueListenerMTS1::OMMQReceiveExecute(
	DWORD dwOriginatingThreadId, 
	BSTR bstrConfig, 
	long lProvider, 
	BSTR bstrConnectionString, 
	BSTR bstrQueueName, 
	VARIANT vMsgId,  
    BSTR bstrTableSuffix,
	BSTR* pbstrErrorMessage, 
	long *plMESSQ_RESP)
{
	LOG_INOUT(_T("%s OMMQReceiveExecute\n"), (LPWSTR)bstrQueueName);
	const LPCTSTR pszFunctionName = _T("OMMQReceiveExecute");

	// assume an error
	long lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
	if (m_spObjectContext)
	{
		m_spObjectContext->SetAbort();
	}
	ResetErrorMessage();

	_ConnectionPtr ConnectionPtr;
	_RecordsetPtr RecordsetPtr;
	enum
	{
		eNULL,
		eOpenQueue,
		eReceive,
		eProgID,
		eXML,
		eCloseQueue,
		eGetComponentInterface,
		eCallComponent,
	} eAction = eNULL;
	
	_bstr_t bstrProgID;
	_bstr_t bstrXML;
	bool bReceivedMessage = false;
	bool bInTransaction = false;

	try
	{
		bInTransaction = m_spObjectContext && m_spObjectContext->IsInTransaction();

		// get the message (ProgID and XML) defined by its message id
		// ... open the connection
		eAction = eOpenQueue;
		ConnectionPtr.CreateInstance(__uuidof(Connection));
		ConnectionPtr->ConnectionString = bstrConnectionString;
		ConnectionPtr->CursorLocation = adUseClient;
		ConnectionPtr->Open("", "", "", 0);

		// ... execute command
		eAction = eReceive;
		RecordsetPtr = HelperOMMQReceiveMessage(
			lProvider, 
			ConnectionPtr, 
			bstrTableSuffix,
			vMsgId);
		if (RecordsetPtr)
		{
			bReceivedMessage = true;

			// ... retrive the ProgID and XML
			eAction = eProgID;
			bstrProgID = RecordsetPtr->GetFields()->GetItem(s_vZero)->GetValue();
			eAction = eXML;
			bstrXML = RecordsetPtr->GetFields()->GetItem(s_vOne)->GetValue();
			CloseRecordset(RecordsetPtr);

			// ... closing connection	
			eAction = eCloseQueue;
			CloseConnection(ConnectionPtr);

			// call the component 
			eAction = eGetComponentInterface;
			CLSID clsid;
			if (SUCCEEDED(CLSIDFromProgID(bstrProgID, &clsid)))
			{
				LOG_INOUT(_T("%s OMMQReceiveExecute - CreateComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrProgID);
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
						LOG_INOUT(_T("%s OMMQReceiveExecute - CallComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrProgID);
						lMESSQ_RESP = MessageQueueComponentVC2Ptr->OnMessage(bstrConfig, bstrXML);
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
							LOG_INOUT(_T("%s OMMQReceiveExecute - CallComponent %s\n"), (LPWSTR)bstrQueueName, (LPWSTR)bstrProgID);
							lMESSQ_RESP = MessageQueueComponentVC1Ptr->OnMessage(bstrXML);
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
						AddToErrorMessage(pszFunctionName, _T("OMMQMTS1 - Error creating component or querying for known interface. Progid (%ls).  Stalling Component.  (The component may be attempting to run as the interactive user, but there no one may be logged on)"), (LPWSTR)bstrProgID);
						lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
					}
				}
			}
			else
			{
				AddToErrorMessage(pszFunctionName, _T("OMMQMTS1 - Error converting Progid (%ls) to CLSID.  Stalling Component"), (LPWSTR)bstrProgID);
				lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
			}
		}
		CloseRecordset(RecordsetPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(_com_error comerr)
	{
		lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
		LogEventErrorProvider(ConnectionPtr);
		switch (eAction)
		{
			case eOpenQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error opening queue"));
				break;
			case eReceive:
				if (comerr.Error() == MQ_ERROR_MESSAGE_ALREADY_RECEIVED)
				{
					lMESSQ_RESP = MESSQ_RESP_HANDLED_STOLEN;
				}
				else
				{
					AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error receiving messsage"));
				}
				break;
			case eProgID:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error getting progid from message"));
				break;
			case eXML:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error getting xml from message"));
				break;
			case eCloseQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error closing queue"));
				break;
			case eGetComponentInterface:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error creating component or querying interface for %s"), (LPWSTR)bstrProgID);
				break;
			case eCallComponent:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error calling component %s"), (LPWSTR)bstrProgID);
				break;
			default :
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Unknown error"));
				break;
		}
		CloseRecordset(RecordsetPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(...)
	{
		lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
		LogEventErrorProvider(ConnectionPtr);
		AddToErrorMessage(pszFunctionName, _T("OMMQMTS1 - Unknown error"));
		CloseRecordset(RecordsetPtr);
		CloseConnection(ConnectionPtr);
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
						HelperOMMQSendMessage(
							lProvider, 
							bstrConnectionString, 
							vMsgId, 
							bstrQueueName, 
							bstrTableSuffix,
							bstrProgID, 
							bstrXML);
						lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_NOW; // modify response
					}
					break;
				case MESSQ_RESP_RETRY_LATER:
					// place the message back on the queue if it has been received (i.e. removed)
					if (bReceivedMessage)
					{
						HelperOMMQSendMessage(
							lProvider, 
							bstrConnectionString, 
							vMsgId, 
							bstrQueueName, 
							bstrTableSuffix,
							bstrProgID, 
							bstrXML);
						lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_LATER; // modify response
					}
					break;
				case MESSQ_RESP_RETRY_MOVE_MESSAGE:
				{
					_ASSERTE(bReceivedMessage);
					_bstr_t bstrQueueNameDest = HelperOMMQGetMoveQueueName(bstrQueueName);
					HelperOMMQSendMessage(
						lProvider, 
						bstrConnectionString, 
						vMsgId, 
						(BSTR)bstrQueueNameDest, 
						bstrTableSuffix,
						bstrProgID, 
						bstrXML);
					lMESSQ_RESP = MESSQ_RESP_HANDLED_RETRY_MOVE_MESSAGE; // modify response
					break;
				}
				case MESSQ_RESP_STALL_COMPONENT:
					_ASSERTE(bReceivedMessage);
					HelperOMMQSendMessage(
						lProvider, 
						bstrConnectionString, 
						vMsgId, 
						bstrQueueName, 
						bstrTableSuffix,
						bstrProgID, 
						bstrXML);
					break;
				case MESSQ_RESP_STALL_QUEUE:
					_ASSERTE(bReceivedMessage);
					HelperOMMQSendMessage(
						lProvider, 
						bstrConnectionString, 
						vMsgId, 
						bstrQueueName, 
						bstrTableSuffix,
						bstrProgID, 
						bstrXML);
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
			lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
			LogEventErrorProvider(ConnectionPtr);
			AddToErrorMessage(pszFunctionName, comerr, 
				_T("OMMQMTS1 - Unknown error"));

		}	
		catch(...)
		{
			lMESSQ_RESP = MESSQ_RESP_HANDLED_ERROR;
			LogEventErrorProvider(ConnectionPtr);
			AddToErrorMessage(pszFunctionName, _T("OMMQMTS1 - Unknown error"));
		}
	}
	
	*pbstrErrorMessage = AllocbstrErrorMessage();
	*plMESSQ_RESP = lMESSQ_RESP;
	return S_OK;
}

STDMETHODIMP CMessageQueueListenerMTS1::OMMQMoveMessage(
	DWORD dwOriginatingThreadId, 
	long lProvider, 
	BSTR bstrConnectionString, 
	BSTR bstrQueueName, 
	BSTR bstrTableSuffix,
	VARIANT vMsgId, 
	BSTR *pbstrErrorMessage)
{
	LOG_INOUT(_T("%s OMMQMoveMessage\n"), (LPWSTR)bstrQueueName);
	const LPCTSTR pszFunctionName = _T("OMMQMoveMessage");

	// move the message from bstrPathName to (bstrPathName + 'x')
	
	// assume a failure
	if (m_spObjectContext)
	{
		m_spObjectContext->SetAbort();
	}
	ResetErrorMessage();

	enum
	{
		eNULL,
		eOpenQueue,
		eMoveMessage,
		eCloseQueue,
	} eAction = eNULL;
	
	_ConnectionPtr ConnectionPtr;
	_CommandPtr CommandPtr;

	try
	{
		// This function should only be called when in a transaction.
		// When not in a transaction (see the receive function for details)
		_ASSERTE(m_spObjectContext && m_spObjectContext->IsInTransaction());

		// open the source queue
		eAction = eOpenQueue;
		ConnectionPtr.CreateInstance(__uuidof(Connection));
		ConnectionPtr->ConnectionString = bstrConnectionString;
		ConnectionPtr->CursorLocation = adUseClient;
		ConnectionPtr->Open("", "", "", 0);
		
		// ... set up the ADO command
		eAction = eMoveMessage;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = ConnectionPtr;
		
		// ... set up the parameters for the commmand
		_ParameterPtr ParameterPtr;
		switch (lProvider)
		{
			case PROVIDER_MSDAORA:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.MoveMessage(?)}";
				ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMsgId);
				break;
			case PROVIDER_SQLOLEDB:
				CommandPtr->CommandType = adCmdStoredProc;
				// RF BMIDS727 Start
				//CommandPtr->CommandText = "USP_MQLOMMQMOVEMESSAGE";
				CDbHelperOMMQ* ptrHelper;
				ptrHelper = new CDbHelperOMMQ(); 
				if (ptrHelper == NULL)
				{
					_ASSERTE(0); // should not reach here
					throw L"Unable to set command text";
				}
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					bstrTableSuffix,
					CDbHelperOMMQ::eMoveMessage);
				ptrHelper = NULL;
				// RF BMIDS727 End
				ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBinary, adParamInput, 16, vMsgId);
				break;
			case PROVIDER_ORAOLEDB:
				CommandPtr->CommandType = adCmdText;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.MoveMessage(?)}";
				ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMsgId);
				break;
			default :
				_ASSERTE(0); // should not reach here 
				throw L"Unrecognised provider";
				break;
		}
		CommandPtr->GetParameters()->Append(ParameterPtr);

		// ... execute command
		CommandPtr->Execute(NULL, NULL, adExecuteNoRecords);

		// ... closing connection	
		eAction = eCloseQueue;
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
		
		if (m_spObjectContext)
		{
			m_spObjectContext->SetComplete();
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eOpenQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error opening source queue"));
				break;
			case eMoveMessage:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error moving messsage"));
				break;
			case eCloseQueue:
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Error closing source queue"));
				break;
			default :
				AddToErrorMessage(pszFunctionName, comerr, _T("OMMQMTS1 - Unknown error"));
				break;
		}
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}
	catch(...)
	{
		AddToErrorMessage(pszFunctionName, _T("pszFunctionName, OMMQMTS1 - Unknown error"));
		CloseCommand(CommandPtr);
		CloseConnection(ConnectionPtr);
	}


	*pbstrErrorMessage = AllocbstrErrorMessage();
	return S_OK;
}

_RecordsetPtr CMessageQueueListenerMTS1::HelperOMMQReceiveMessage(
	long lProvider, 
	_ConnectionPtr ConnectionPtr, 
	BSTR bstrTableSuffix,
	VARIANT vMsgId)
{
	// ... set up the ADO command
	_CommandPtr CommandPtr(__uuidof(Command));
	CommandPtr->ActiveConnection = ConnectionPtr;

	// ... set up the parameters for the commmand
	_ParameterPtr ParameterPtr;
	switch (lProvider)
	{
		case PROVIDER_MSDAORA:
			// LONG's only supported by this provider
			CommandPtr->CommandType = adCmdText;
			CommandPtr->CommandText = "{call sp_MQLOMMQ.GetMessageOrLONG(?)}";
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMsgId);
			break;
		case PROVIDER_SQLOLEDB:
			CommandPtr->CommandType = adCmdStoredProc;
			// RF BMIDS727 Start
			//CommandPtr->CommandText = "USP_MQLOMMQGETMESSAGEORNTEXT";
			CDbHelperOMMQ* ptrHelper;
			ptrHelper = new CDbHelperOMMQ(); 
			if (ptrHelper == NULL)
			{
				_ASSERTE(0); // should not reach here
				throw L"Unable to set command text";
			}
			CommandPtr->CommandText = 
				ptrHelper->GetSQLOLEDBCommandText(
				bstrTableSuffix,
				CDbHelperOMMQ::eGetMessage);
			ptrHelper = NULL;
			// RF BMIDS727 End
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBinary, adParamInput, 16, vMsgId);
			break;
		case PROVIDER_ORAOLEDB:
			// CLOBs are more performant than LONGs
			CommandPtr->CommandType = adCmdText;
			CommandPtr->CommandText = "{call sp_MQLOMMQ.GetMessageOrCLOB(?)}";
			CommandPtr->Properties->Item["SPPrmsLOB"]->Value = VARIANT_TRUE;
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adGUID, adParamInput, 0, vMsgId);
			break;
		default :
			_ASSERTE(0); // should not reach here 
			throw L"Unrecognised provider";
			break;
	}
	CommandPtr->GetParameters()->Append(ParameterPtr);

	// ... execute command
	_RecordsetPtr RecordsetPtr;
	RecordsetPtr = CommandPtr->Execute(NULL, NULL, 0);
	return RecordsetPtr;
}


void CMessageQueueListenerMTS1::HelperOMMQSendMessage(
	long lProvider, 
	BSTR bstrConnectionString, 
	VARIANT vMsgId, 
	BSTR bstrQueueName, 
	BSTR bstrTableSuffix,
	BSTR bstrProgID, 
	const _bstr_t& bstrXML)
{
	_ConnectionPtr ConnectionPtr(__uuidof(Connection));
	ConnectionPtr->ConnectionString = bstrConnectionString;
	ConnectionPtr->CursorLocation = adUseClient;
	ConnectionPtr->Open("", "", "", 0);
	HelperOMMQSendMessage(
		lProvider, 
		ConnectionPtr, 
		vMsgId, 
		bstrQueueName, 
		bstrTableSuffix,
		bstrProgID, 
		bstrXML);
	CloseConnection(ConnectionPtr);
}

void CMessageQueueListenerMTS1::HelperOMMQSendMessage(
	long lProvider, 
	_ConnectionPtr ConnectionPtr, 
	VARIANT vMsgId, 
	BSTR bstrQueueName, 
	BSTR bstrTableSuffix,
	BSTR bstrProgID, 
	const _bstr_t& bstrXML)
{
	long lMessageSize = bstrXML.length();

	// ... set up the ADO command
	_CommandPtr CommandPtr(__uuidof(Command));
	CommandPtr->ActiveConnection = ConnectionPtr;

	// ... set up the parameters for the commmand
	_ParameterPtr ParameterPtr;
	bool bMSDAORAChunking = false;
	switch (lProvider)
	{
		case PROVIDER_MSDAORA:
			CommandPtr->CommandType = adCmdText;
			if (lMessageSize > 32000) // corresponds to LONG variable
			{
				CommandPtr->CommandText = "{call sp_MQLOMMQ.SendMessage(?,?,?,?,?,?)}";
				bMSDAORAChunking = true;
			}
			else if (lMessageSize > 2000) // corresponds to size of XML field on database
			{
				CommandPtr->CommandText = "{call sp_MQLOMMQ.SendMessageLONG(?,?,?,?,?,?)}";
			}
			else
			{
				CommandPtr->CommandText = "{call sp_MQLOMMQ.SendMessage(?,?,?,?,?,?)}";
			}
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, 0, vMsgId);
			break;
		case PROVIDER_SQLOLEDB:
			CommandPtr->CommandType = adCmdStoredProc;
			// RF BMIDS727 Start
			CDbHelperOMMQ* ptrHelper;
			ptrHelper = new CDbHelperOMMQ(); 
			if (ptrHelper == NULL)
			{
				_ASSERTE(0); // should not reach here
				throw L"Unable to set command text";
			}
			//if (lMessageSize > 4000) // corresponds to size of XML field on database
			if (lMessageSize > 3800) // corresponds to size of XML field on database
			{
				//CommandPtr->CommandText = "USP_MQLOMMQSENDMESSAGENTEXT";
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					bstrTableSuffix,
					CDbHelperOMMQ::eSendMessageOrNtext);
			}
			else
			{
				//CommandPtr->CommandText = "USP_MQLOMMQSENDMESSAGE";
				CommandPtr->CommandText = 
					ptrHelper->GetSQLOLEDBCommandText(
					bstrTableSuffix,
					CDbHelperOMMQ::eSendMessage);
			}
			ptrHelper = NULL;
			// RF BMIDS727 End
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBinary, adParamInput, 16, vMsgId);
			break;
		case PROVIDER_ORAOLEDB:
			CommandPtr->CommandType = adCmdText;
			if (lMessageSize > 2000) // corresponds to size of XML field on database
			{
				CommandPtr->Properties->Item["SPPrmsLOB"]->Value = VARIANT_TRUE;
				CommandPtr->CommandText = "{call sp_MQLOMMQ.SendMessageCLOB(?,?,?,?,?,?)}";
			}
			else
			{
				CommandPtr->CommandText = "{call sp_MQLOMMQ.SendMessage(?,?,?,?,?,?)}";
			}
			ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, 0, vMsgId);
			break;
		default :
			_ASSERTE(0); // should not reach here 
			throw L"Unrecognised provider";
			break;
	}
	CommandPtr->GetParameters()->Append(ParameterPtr);
	ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, SysStringLen(bstrQueueName), bstrQueueName);
	CommandPtr->GetParameters()->Append(ParameterPtr);
	ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, SysStringLen(bstrProgID), bstrProgID);
	CommandPtr->GetParameters()->Append(ParameterPtr);

	if (bMSDAORAChunking)
	{
		ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, 0, s_vNull);
	    ParameterPtr->Attributes = adParamNullable;
	}
	else
	{
		ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adBSTR, adParamInput, SysStringLen(bstrXML), bstrXML);
	}
	CommandPtr->GetParameters()->Append(ParameterPtr);

	// not preserving priority or execute after date when not running under MTS
	_variant_t vPriority((long)3); // priority is not preserded
	ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adInteger, adParamInput, 0, vPriority);
	CommandPtr->GetParameters()->Append(ParameterPtr);

	ParameterPtr = CommandPtr->CreateParameter(s_bstrNULL, adDate, adParamInput, 0, s_vNull);
	ParameterPtr->Attributes = adParamNullable;
	CommandPtr->GetParameters()->Append(ParameterPtr);

	if (bMSDAORAChunking)
	{
		_RecordsetPtr RecordsetPtr(__uuidof(Recordset));
        RecordsetPtr->CursorType = adOpenKeyset;
        RecordsetPtr->LockType = adLockOptimistic;
        RecordsetPtr->Open(_T("SELECT MessageId, Xml FROM MQLOMMQLONG WHERE MESSAGEID IS NULL"), (IUnknown*)ConnectionPtr, adOpenUnspecified, adLockUnspecified, adCmdText);
        
        RecordsetPtr->AddNew();
        RecordsetPtr->GetFields()->GetItem(s_vZero)->Value = vMsgId;
        
        LPCTSTR pszXml = bstrXML;
		long lChunkSize = 32000;
        _bstr_t bstrChunk;
        long lOffset = 0;
        while (lOffset < lMessageSize)
		{
			lChunkSize = min(lChunkSize, lMessageSize - lOffset);
			bstrChunk = _bstr_t(SysAllocStringLen(pszXml + lOffset, lChunkSize), false);
			RecordsetPtr->GetFields()->GetItem(s_vOne)->AppendChunk(bstrChunk);
			lOffset = lOffset + lChunkSize;
        }
        RecordsetPtr->Update();
        CloseRecordset(RecordsetPtr);
    }

	// ... execute command
	CommandPtr->Execute(NULL, NULL, adExecuteNoRecords);
	CloseCommand(CommandPtr);
}

void CMessageQueueListenerMTS1::LogEventErrorProvider(_ConnectionPtr& ConnectionPtr)
{
	const LPCTSTR pszFunctionName = _T("LogEventErrorProvider");

    try
	{
		// Print Provider Errors from Connection object.
		// pErr is a record object in the Connection's Error collection.
		ErrorPtr    pErr  = NULL;
		long      nCount  = 0;    
		long      i     = 0;

		if( (ConnectionPtr->Errors->Count) > 0)
		{
			nCount = ConnectionPtr->Errors->Count;
			// Collection ranges from 0 to nCount -1.
			for(i = 0; i < nCount; i++)
			{
				pErr = ConnectionPtr->Errors->GetItem(i);
				TCHAR buff[MAXLOGMESSAGESIZE];
				_bstr_t bstrDescription = pErr->Description;
				VERIFY(_sntprintf(buff, MAXLOGMESSAGESIZE - 1, _T("Error number: %x.  Error Description: %s."), 
					pErr->Number,(LPCTSTR) bstrDescription) > -1);
				AddToErrorMessage(pszFunctionName, buff);
			}
		}
	}
	catch(...)
	{
	}
}

void CMessageQueueListenerMTS1::CloseConnection(_ConnectionPtr& ConnectionPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseConnection");

	try
	{
		if (ConnectionPtr != NULL && ConnectionPtr->State == adStateOpen)
		{
			ConnectionPtr->Close();
		}
		ConnectionPtr = NULL;
	}
	catch(...)
	{
		try
		{
			ConnectionPtr = NULL;
		}
		catch(...)
		{
		}
	}
}

void CMessageQueueListenerMTS1::CloseRecordset(_RecordsetPtr& RecordsetPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseRecordset");

	try
	{
		if (RecordsetPtr != NULL && RecordsetPtr->State == adStateOpen)
		{
			RecordsetPtr->Close();
		}
		RecordsetPtr = NULL;
	}
	catch(...)
	{
		try
		{
			RecordsetPtr = NULL;
		}
		catch(...)
		{
		}
	}
}

void CMessageQueueListenerMTS1::CloseCommand(_CommandPtr& CommandPtr)
{
	const LPCTSTR pszFunctionName = _T("CloseCommand");

	try
	{
		CommandPtr = NULL;
	}
	catch(...)
	{
	}
}
