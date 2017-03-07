///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestTransact.cpp
//	DESCRIPTION:	
//		Implements sending request to Transact and receiving a response. The
//		request is first converted from XML into Transact byte stream format,
//		before being sent to Transact. The response from Transact is converted
//		back from byte stream format to XML.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "MetaData.h"
#include "MetaDataEnvTransact.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "RequestTransact.h"
#include "Profiler.h"
#include "WSASocket.h"
#include <time.h>


namespaceMutex::CCriticalSection CRequestTransact::s_csWriteBufferLastRequest;
namespaceMutex::CCriticalSection CRequestTransact::s_csWriteBufferLastResponse;
namespaceMutex::CCriticalSection CRequestTransact::s_csWriteXML;


static const LPCWSTR g_pszMethod					= L"//REQUEST/METHOD";
static const LPCWSTR g_pszTestResponse				= L"//REQUEST/TESTRESPONSE";
static const LPCWSTR g_pszResponse					= L"//RESPONSE";
static const LPCWSTR g_pszData						= L"DATA";
static const LPCWSTR g_pszDefault					= L"DEFAULT";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestTransact::CRequestTransact(LPCWSTR szType) :
	CRequest(szType)
{
}

CRequestTransact::~CRequestTransact()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestTransact::Execute
//	
//	Description:
//		Implementation of virtual method.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//	
//	Return:
//		IXMLDOMNodePtr: 
//			The response in "Transact XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestTransact::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		CMetaDataEnvTransact* pMetaDataEnv = NULL;

		IF_PROFILE(CProfiler::StartProfiling());

		IF_PROFILE(CProfiler::StartTimer(_T("CRequestTransactExecute")));

		IF_PROFILE(CProfiler::StartTimer(_T("GetMetaDataEnv")));
		// Get the ODI meta data environment; it must be of type "TRANSACT".
		pMetaDataEnv = reinterpret_cast<CMetaDataEnvTransact*>(CMetaData::GetMetaDataEnv(ptrRequestNode, L"TRANSACT"));
		IF_PROFILE(CProfiler::StopTimer(_T("GetMetaDataEnv")));

		// Put a read lock on the meta data to prevent it changing for the duration of this request.
		namespaceMutex::CReadLock lckReadLock(pMetaDataEnv->GetSharedLock());

		// Random sleep of 0 - 10 seconds useful for testing locking.
		//Sleep(((2 * rand() * 10000 + RAND_MAX) / RAND_MAX - 1) / 2);

		// Requires pMetaDataEnv to be defined.
		IF_LOG(_Module.LogDebug(_T("->CRequestTransact::Execute()\n")));

		ptrResponseNode = GetResponseNode(ptrRequestNode, pMetaDataEnv);

		IF_PROFILE(CProfiler::StopTimer(_T("CRequestTransactExecute")));

		IF_PROFILE(CProfiler::StopProfiling());

		IF_LOG(_Module.LogDebug(_T("<-CRequestTransact::Execute()\n\n")));
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return ptrResponseNode;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestTransact::GetResponseNode
//	
//	Description:
//		Gets the response from Transact to request.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Transact XML" format.
//		CMetaDataEnvTransact* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Transact object map meta data, that maps Transact object and 
//			property ids to the names used in the XML.
//	
//	Return:
//		IXMLDOMNodePtr: 	
//			Response in "Transact XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestTransact::GetResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvTransact* pMetaDataEnv) const
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		ptrResponseNode = GetTestResponseNode(ptrRequestNode, pMetaDataEnv);

		if (ptrResponseNode == NULL)
		{
			// There is no dummy test response for this request, so get real response from Transact.
			ptrResponseNode = GetTransactResponseNode(ptrRequestNode, pMetaDataEnv);
		}
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return ptrResponseNode;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestTransact::GetTransactResponseNode
//	
//	Description:
//		Gets real response from Transact.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Transact XML" format.
//		CMetaDataEnvTransact* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Transact object map meta data, that maps Transact object and 
//			property ids to the names used in the XML.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Response in "Transact XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestTransact::GetTransactResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvTransact* pMetaDataEnv) const
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		CWSASocket Socket;
		bool bSuccess = false;
		const int nMaxSizeBuffer = 0xFFFF;
		BYTE Buffer[nMaxSizeBuffer];
		int nRequestSize = 0;
		int nResponseSize = 0;

		if (wcslen(ptrRequestNode->xml) < nMaxSizeBuffer)
		{
			IF_PROFILE(CProfiler::StartTimer(_T("ConvertRequest")));
			// TODO Convert request XML to buffer to send to Transact.
			IF_PROFILE(CProfiler::StopTimer(_T("ConvertRequest")));
		}
		else
		{
			throw CException(E_BUFFEROVERRUN, __FILE__, __LINE__, _T("Request buffer too small"));
		}

		LPCWSTR szTransactHost	= pMetaDataEnv->GetTransactHost();
		int nTransactPort		= GetTransactPort(ptrRequestNode, pMetaDataEnv);
		int nSktTimeOutMs		= pMetaDataEnv->GetSktTimeOutMs();
		int nSktMaxSend			= pMetaDataEnv->GetSktMaxSend();
		int nSktMaxRecv			= pMetaDataEnv->GetSktMaxRecv();

		Socket.SetSendBlockSize(nSktMaxSend);
		Socket.SetRecvBlockSize(nSktMaxRecv);

		IF_LOG(_Module.LogDebug(_T("TRANSACTHOST = %s, TRANSACTPORT = %d, SKTTIMEOUTMS = %d, SKTMAXSEND = %d, SKTMAXRECV = %d\n"), szTransactHost, nTransactPort, nSktTimeOutMs, nSktMaxSend, nSktMaxRecv));

		IF_PROFILE(CProfiler::StartTimer(_T("Connect")));
		bSuccess = 
			Socket.Open() &&
			Socket.SetSockOpt(SOL_SOCKET, SO_SNDTIMEO, &nSktTimeOutMs, sizeof(nSktTimeOutMs)) &&
			Socket.SetSockOpt(SOL_SOCKET, SO_RCVTIMEO, &nSktTimeOutMs, sizeof(nSktTimeOutMs)) &&
			Socket.Connect(szTransactHost, nTransactPort) == TRUE;
		IF_PROFILE(CProfiler::StopTimer(_T("Connect")));

#ifdef _DEBUG
		namespaceMutex::CSingleLock lckWriteBufferLastRequest(&s_csWriteBufferLastRequest, TRUE);
		WriteBuffer(_T("last_request"), Buffer, nRequestSize);
		lckWriteBufferLastRequest.Unlock();
#endif

		IF_PROFILE(CProfiler::StartTimer(_T("Transact")));

		IF_PROFILE(CProfiler::StartTimer(_T("Send")));
		bSuccess = Socket.SendBuffer(reinterpret_cast<LPCSTR>(Buffer), nRequestSize) != SOCKET_ERROR;
		IF_PROFILE(CProfiler::StopTimer(_T("Send")));

		if (bSuccess)
		{
			IF_PROFILE(CProfiler::StartTimer(_T("Recv")));
			nResponseSize = Socket.RecvBuffer(reinterpret_cast<LPSTR>(Buffer), nMaxSizeBuffer);
			IF_PROFILE(CProfiler::StopTimer(_T("Recv")));
			bSuccess = nResponseSize != SOCKET_ERROR && nResponseSize > 0;

			if (!bSuccess)
			{
				throw CException(E_INVALIDRESPONSE, __FILE__, __LINE__, _T("Invalid response"));
			}
		}

		IF_PROFILE(CProfiler::StopTimer(_T("Transact")));

#ifdef _DEBUG
		namespaceMutex::CSingleLock lckWriteBufferLastResponse(&s_csWriteBufferLastResponse, TRUE);
		WriteBuffer(_T("last_response"), Buffer, nResponseSize);
		lckWriteBufferLastResponse.Unlock();
#endif

		if (bSuccess)
		{
			IF_PROFILE(CProfiler::StartTimer(_T("ConvertResponse")));
			// TODO Convert Transact response buffer to XML.
			IF_PROFILE(CProfiler::StopTimer(_T("ConvertResponse")));
		}

		IF_PROFILE(CProfiler::StartTimer(_T("Disconnect")));
		Socket.Close();
		IF_PROFILE(CProfiler::StopTimer(_T("Disconnect")));
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return ptrResponseNode;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestTransact::GetTransactPort
//	
//	Description:
//		Gets the Transact port for this request.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Transact XML" format.
//		CMetaDataEnvTransact* pMetaDataEnv:
//			The ODIConverter meta data environment. 
//	
//	Return:
//		int: 
//			The Transact port.
///////////////////////////////////////////////////////////////////////////////
int CRequestTransact::GetTransactPort(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvTransact* pMetaDataEnv) const
{
	int nTransactPort = _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"TRANSACTPORT"));
	if (nTransactPort == 0)
	{
		// Port not specified so use the default Transact sign on port.
		nTransactPort = pMetaDataEnv->GetTransactPort();
	}

	return nTransactPort;
}


///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestTransact::GetTestResponseNode
//	
//	Description:
//		Gets test response.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request.
//		CMetaDataEnvTransact* pMetaDataEnv:
//			The ODIConverter meta data environment.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Test response. Will be NULL if there is no test response.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestTransact::GetTestResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvTransact* pMetaDataEnv) const
{
	MSXML::IXMLDOMNodePtr ptrTestResponseNode = NULL;

	try
	{
		MSXML::IXMLDOMDocumentPtr ptrTestResponseDoc(__uuidof(MSXML::DOMDocument));;
		MSXML::IXMLDOMNodePtr ptrRequestTestResponseNode = NULL;

		ptrRequestTestResponseNode = ptrRequestNode->selectSingleNode(g_pszTestResponse);
		if (ptrRequestTestResponseNode != NULL)
		{
			// Test response is included as a child node in the incoming request; use this to create response.
			MSXML::IXMLDOMNodePtr ptrXMLDOMParentNode = ptrRequestTestResponseNode->GetparentNode();
			
			if (ptrXMLDOMParentNode != NULL)
			{
				// Remove test response node from parent.
				ptrRequestTestResponseNode = ptrXMLDOMParentNode->removeChild(ptrRequestTestResponseNode);
				if (ptrRequestTestResponseNode != NULL)
				{
					// Test response contains the dummy response XML as an attribute.
					LPCWSTR szXML = static_cast<IXMLAssistDOMNodePtr>(ptrRequestTestResponseNode)->getAttributeText(L"XML");
					ptrTestResponseDoc->loadXML(szXML);
					ptrTestResponseNode = ptrTestResponseDoc->documentElement;
				}
			}

			if (ptrTestResponseNode == NULL)
			{
				throw CException(E_GENERICERROR, __FILE__, __LINE__, L"Invalid TESTRESPONSE node");
			}
		}
		else if (pMetaDataEnv->GetTestResponse())
		{
			// This environment sends back test responses, and does not call Transact.

			IF_LOG(_Module.LogDebug(_T("TESTRESPONSE = \"Y\"\n")));

			// Read the response from a file in the ODIConverter runtime directory.
			MSXML::IXMLDOMNodePtr ptrMethodNode(ptrRequestNode->selectSingleNode(g_pszMethod));
			if (ptrMethodNode != NULL)
			{
				LPCWSTR szMethod = static_cast<IXMLAssistDOMNodePtr>(ptrMethodNode)->getAttributeText(L"DATA");

				if (wcslen(szMethod) > 0)
				{
					_TCHAR szTestResponsePath[_MAX_PATH] = _T("");
					_stprintf(szTestResponsePath, _T("%s%s%s_response.xml"), _Module.m_szDrive, _Module.m_szDir, szMethod);
					ptrTestResponseDoc->async = false;
					_variant_t varbLoaded = ptrTestResponseDoc->load(szTestResponsePath);
					if (varbLoaded.boolVal != false)
					{
						ptrTestResponseNode = ptrTestResponseDoc->selectSingleNode(g_pszResponse);
					}
					else
					{
						throw CException(E_GENERICERROR, __FILE__, __LINE__, L"Invalid TESTRESPONSE file: %s", szTestResponsePath);
					}
				}
			}

			if (ptrTestResponseNode == NULL)
			{
				throw CException(E_GENERICERROR, __FILE__, __LINE__, L"Invalid TESTRESPONSE node");
			}

			::Sleep(pMetaDataEnv->GetTestResponseDelayMs());
		}
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return ptrTestResponseNode;
}

void CRequestTransact::WriteXML(MSXML::IXMLDOMNodePtr ptrNode, LPCTSTR pszType, int nMaxWriteXML, int nIncrement) const
{
	if (nMaxWriteXML > 0 && ptrNode != NULL)
	{
		static int nWriteXML = 0;
		if (nWriteXML > nMaxWriteXML)
		{
			nWriteXML = 0;
		}
		nWriteXML += nIncrement;
		_TCHAR szFName[_MAX_PATH] = _T("");
		_stprintf(szFName, _T("ODIConverter%s%03d"), pszType, nWriteXML);
		namespaceMutex::CSingleLock lckWriteXML(&s_csWriteXML, TRUE);
		CRequest::WriteXML(szFName, ptrNode);
	}
}

