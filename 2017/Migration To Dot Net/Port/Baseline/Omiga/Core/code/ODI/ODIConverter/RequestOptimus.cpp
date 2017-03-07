///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestOptimus.cpp
//	DESCRIPTION:	
//		Implements sending request to Optimus and receiving a response. The
//		request is first converted from XML into Optimus byte stream format,
//		before being sent to Optimus. The response from Optimus is converted
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
#include "MetaDataEnvOptimus.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "RequestOptimus.h"
#include "OptimusResponse2XML.h"
#include "OptimusXML2Request.h"
#include "Profiler.h"
#include "WSASocket.h"
#include <time.h>


namespaceMutex::CCriticalSection CRequestOptimus::s_csWriteBufferLastRequest;
namespaceMutex::CCriticalSection CRequestOptimus::s_csWriteBufferLastResponse;
namespaceMutex::CCriticalSection CRequestOptimus::s_csWriteXML;


static const LPCWSTR g_pszMethod					= L"//REQUEST/METHOD";
static const LPCWSTR g_pszTestResponse				= L"//REQUEST/TESTRESPONSE";
static const LPCWSTR g_pszResponse					= L"//RESPONSE";
static const LPCWSTR g_pszData						= L"DATA";
static const LPCWSTR g_pszDefault					= L"DEFAULT";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestOptimus::CRequestOptimus(LPCWSTR szType) :
	CRequest(szType)
{
}

CRequestOptimus::~CRequestOptimus()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestOptimus::Execute
//	
//	Description:
//		Implementation of virtual method.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			XML containing the request in "Optimus XML" format. The root node
//			should be of the form "<REQUEST ODIENVIRONMENT="ENV"... />, where
//			"ENV" is the name of the ODI meta data environment to use.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			The response in "Optimus XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestOptimus::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		CMetaDataEnvOptimus* pMetaDataEnv = NULL;

		IF_PROFILE(CProfiler::StartProfiling());

		IF_PROFILE(CProfiler::StartTimer(_T("CRequestOptimusExecute")));

		IF_PROFILE(CProfiler::StartTimer(_T("GetMetaDataEnv")));
		// Get the ODI meta data environment; it must be of type "OPTIMUS".
		pMetaDataEnv = reinterpret_cast<CMetaDataEnvOptimus*>(CMetaData::GetMetaDataEnv(ptrRequestNode, L"OPTIMUS"));
		IF_PROFILE(CProfiler::StopTimer(_T("GetMetaDataEnv")));

		// Put a read lock on the meta data to prevent it changing for the duration of this request.
		namespaceMutex::CReadLock lckReadLock(pMetaDataEnv->GetSharedLock());

		// Random sleep of 0 - 10 seconds useful for testing locking.
		//Sleep(((2 * rand() * 10000 + RAND_MAX) / RAND_MAX - 1) / 2);

		// Requires pMetaDataEnv to be defined.
		IF_LOG(_Module.LogDebug(_T("->CRequestOptimus::Execute()\n")));

		IF_LOG(_Module.LogDebug(_T("OPTIMUSOBJECTMAP = %s\n"), (LPCTSTR)pMetaDataEnv->GetOptimusObjectMapPath()));
		IF_LOG(_Module.LogDebug(_T("CODEPAGEMAP = %s\n"), (LPCTSTR)pMetaDataEnv->GetCodePageMapPath()));

		// Write last 100 requests/reponses.
		WriteXML(ptrRequestNode, _T("Request"), pMetaDataEnv->GetLogXML(), 1);

		ptrResponseNode = GetResponseNode(ptrRequestNode, pMetaDataEnv, pMetaDataEnv->GetCodePage(ptrRequestNode));

		WriteXML(ptrResponseNode, _T("Response"), pMetaDataEnv->GetLogXML(), 0);

		IF_PROFILE(CProfiler::StopTimer(_T("CRequestOptimusExecute")));

		if (ptrResponseNode != NULL)
		{
			IF_PROFILE
			(
				if (pMetaDataEnv && pMetaDataEnv->GetProfile())
				{
					CProfiler::GetTimesXML(ptrResponseNode);
					IF_LOG(LogProfile(ptrResponseNode));
				}
			)
		}

		IF_PROFILE(CProfiler::StopProfiling());

		IF_LOG(_Module.LogDebug(_T("<-CRequestOptimus::Execute()\n\n")));
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
//	Function: CRequestOptimus::GetResponseNode
//	
//	Description:
//		Gets the response from Optimus to request.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Optimus XML" format.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property ids to the names used in the XML.
//		CCodePage* pCodePage:
//			The code page to be used in converting the payload data from 
//			Unicode to Ebcdic and vice versa.
//	
//	Return:
//		IXMLDOMNodePtr: 	
//			Response in "Optimus XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestOptimus::GetResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvOptimus* pMetaDataEnv, 
	CCodePage* pCodePage) const
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		ptrResponseNode = GetTestResponseNode(ptrRequestNode, pMetaDataEnv, pCodePage);

		if (ptrResponseNode == NULL)
		{
			// There is no dummy test response for this request, so get real response from Optimus.
			ptrResponseNode = GetOptimusResponseNode(ptrRequestNode, pMetaDataEnv, pCodePage);
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
//	Function: CRequestOptimus::GetOptimusResponseNode
//	
//	Description:
//		Gets real response from Optimus.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Optimus XML" format.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property ids to the names used in the XML.
//		CCodePage* pCodePage:
//			The code page to be used in converting the payload data from 
//			Unicode to Ebcdic and vice versa.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Response in "Optimus XML" format.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestOptimus::GetOptimusResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvOptimus* pMetaDataEnv,
	CCodePage* pCodePage) const
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

		bool bUpdatedStartNewSession = UpdateStartNewSessionRequest(ptrRequestNode, pMetaDataEnv);
		if (bUpdatedStartNewSession)
		{
			// We've changed the request, so write it out to file.
			WriteXML(ptrRequestNode, _T("Request"), pMetaDataEnv->GetLogXML(), 0);
		}

		if (wcslen(ptrRequestNode->xml) < nMaxSizeBuffer)
		{
			IF_PROFILE(CProfiler::StartTimer(_T("ConvertRequest")));
			nRequestSize = COptimusXML2Request::Convert(ptrRequestNode, Buffer, sizeof(Buffer), pMetaDataEnv, pCodePage);
			IF_PROFILE(CProfiler::StopTimer(_T("ConvertRequest")));
		}
		else
		{
			throw CException(E_BUFFEROVERRUN, __FILE__, __LINE__, _T("Request buffer too small"));
		}

		LPCWSTR szOptimusHost	= pMetaDataEnv->GetOptimusHost();
		int nOptimusPort		= GetOptimusPort(ptrRequestNode, pMetaDataEnv);
		int nSktTimeOutMs		= pMetaDataEnv->GetSktTimeOutMs();
		int nSktMaxSend			= pMetaDataEnv->GetSktMaxSend();
		int nSktMaxRecv			= pMetaDataEnv->GetSktMaxRecv();

		Socket.SetSendBlockSize(nSktMaxSend);
		Socket.SetRecvBlockSize(nSktMaxRecv);

		IF_LOG(_Module.LogDebug(_T("OPTIMUSHOST = %s, OPTIMUSPORT = %d, SKTTIMEOUTMS = %d, SKTMAXSEND = %d, SKTMAXRECV = %d\n"), szOptimusHost, nOptimusPort, nSktTimeOutMs, nSktMaxSend, nSktMaxRecv));

		IF_PROFILE(CProfiler::StartTimer(_T("Connect")));
		bSuccess = 
			Socket.Open() &&
			Socket.SetSockOpt(SOL_SOCKET, SO_SNDTIMEO, &nSktTimeOutMs, sizeof(nSktTimeOutMs)) &&
			Socket.SetSockOpt(SOL_SOCKET, SO_RCVTIMEO, &nSktTimeOutMs, sizeof(nSktTimeOutMs)) &&
			Socket.Connect(szOptimusHost, nOptimusPort) == TRUE;
		IF_PROFILE(CProfiler::StopTimer(_T("Connect")));

#ifdef _DEBUG
		namespaceMutex::CSingleLock lckWriteBufferLastRequest(&s_csWriteBufferLastRequest, TRUE);
		WriteBuffer(_T("last_request"), Buffer, nRequestSize);
		lckWriteBufferLastRequest.Unlock();
#endif

		IF_PROFILE(CProfiler::StartTimer(_T("Optimus")));

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

		IF_PROFILE(CProfiler::StopTimer(_T("Optimus")));

#ifdef _DEBUG
		namespaceMutex::CSingleLock lckWriteBufferLastResponse(&s_csWriteBufferLastResponse, TRUE);
		WriteBuffer(_T("last_response"), Buffer, nResponseSize);
		lckWriteBufferLastResponse.Unlock();
#endif

		if (bSuccess)
		{
			IF_PROFILE(CProfiler::StartTimer(_T("ConvertResponse")));
			ptrResponseNode = COptimusResponse2XML::Convert(Buffer, nResponseSize, pMetaDataEnv, pCodePage);
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
//	Function: CRequestOptimus::GetOptimusPort
//	
//	Description:
//		Gets the Optimus port for this request.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Optimus XML" format.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. 
//	
//	Return:
//		int: 
//			The Optimus port.
///////////////////////////////////////////////////////////////////////////////
int CRequestOptimus::GetOptimusPort(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvOptimus* pMetaDataEnv) const
{
	int nOptimusPort = _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"OPTIMUSPORT"));
	if (nOptimusPort == 0)
	{
		// Port not specified so use the Optimus sign on port (assumes this is a sign on request).
		// For other types of request the port number should be specified, having been previously
		// returned from Optimus in a sign on response.
		nOptimusPort = pMetaDataEnv->GetOptimusPort();
	}

	return nOptimusPort;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestOptimus::UpdateStartNewSessionRequest
//	
//	Description:
//		If the request is a "startNewSession", this function updates it if 
//		necessary.
//
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Optimus XML" format.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. 
//	
//	Return:
//		bool: 
//			Returns true if the request has been updated.
///////////////////////////////////////////////////////////////////////////////
bool CRequestOptimus::UpdateStartNewSessionRequest(
	const MSXML::IXMLDOMNodePtr ptrRequestNode,
	CMetaDataEnvOptimus* pMetaDataEnv) const
{
	bool bUpdatedHost	= false;
	bool bUpdatedEnv	= false;

	// Check whether this is a startNewSession request.
	IXMLAssistDOMNodePtr ptrXMLDOMNodeMethod = ptrRequestNode->selectSingleNode(L"//REQUEST/METHOD");
	if (ptrXMLDOMNodeMethod != NULL && wcsicmp(ptrXMLDOMNodeMethod->getAttributeText(L"DATA"), L"startNewSession") == 0)
	{
		// startNewSession request; if the Optimus HOST and ENVIRONMENT are not defined in
		// the request then set them from the ODI environment meta data.
		IXMLAssistDOMNodePtr ptrSignOnNode = ptrRequestNode->selectSingleNode(L".//SIGNONPROFILE");
		bUpdatedHost	= UpdateAttribute(ptrSignOnNode, L"HOST", L"DATA", pMetaDataEnv->GetOptimusHostName());
		bUpdatedEnv		= UpdateAttribute(ptrSignOnNode, L"ENVIRONMENT", L"DATA", pMetaDataEnv->GetOptimusEnvironment());
	}

	return bUpdatedHost || bUpdatedEnv;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestOptimus::UpdateAttribute
//	
//	Description:
//		Updates an attribute with a default value if the attribute is not alrea
//		defined. Will create the attribute's node if necessary.dy 
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrParentNode:
//			The parent node for the node containing the attribute.
//		LPCWSTR szNodeKey:
//			The XPATH relative to the parent node for the node containing the 
//			attribute.
//		LPCWSTR szAttributeKey:
//			The name of the attribute.
//		LPCWSTR szAttributeValue:
//			The default value for the attribute.
//	
//	Return:
//		bool: 	
//			Returns true if the attribute is updated.
///////////////////////////////////////////////////////////////////////////////
bool CRequestOptimus::UpdateAttribute(
	const MSXML::IXMLDOMNodePtr ptrParentNode,
	LPCWSTR szNodeKey,
	LPCWSTR szAttributeKey,
	LPCWSTR szAttributeValue) const
{
	bool bUpdated = false;

	if (ptrParentNode != NULL)
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrParentNode->selectSingleNode(szNodeKey);
		_bstr_t bstrAttributeValue = L"";
		if (ptrXMLDOMElement != NULL)
		{
			// Element exists, so get attribute.
			_variant_t varAttributeValue = ptrXMLDOMElement->getAttribute(szAttributeKey);
			if (varAttributeValue.vt != VT_NULL && wcslen(varAttributeValue.bstrVal) > 0)
			{
				bstrAttributeValue = varAttributeValue.bstrVal;
			}
		}
		else
		{
			// Element does not exists so create it.
			MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument = ptrParentNode->ownerDocument;
			if (ptrXMLDOMDocument != NULL)
			{
				ptrXMLDOMElement = ptrXMLDOMDocument->createElement(szNodeKey);
				ptrParentNode->appendChild(ptrXMLDOMElement);
			}
			else
			{
				_ASSERTE(FALSE);
			}
		}
		if (wcslen(bstrAttributeValue) == 0)
		{
			// Attribute not defined, so set it to default value.
			ptrXMLDOMElement->setAttribute(szAttributeKey, szAttributeValue);
			bUpdated = true;
		}
	}
	else
	{
		_ASSERTE(FALSE);
	}

	return bUpdated;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestOptimus::GetTestResponseNode
//	
//	Description:
//		Gets test response.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			Request in "Optimus XML" format.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property ids to the names used in the XML.
//		CCodePage* pCodePage:
//			The code page to be used in converting the payload data from 
//			Unicode to Ebcdic and vice versa.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Test response in "Optimus XML" format. Will be NULL if there is no
//			test response.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestOptimus::GetTestResponseNode(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	CMetaDataEnvOptimus* pMetaDataEnv, 
	CCodePage* pCodePage) const
{
	MSXML::IXMLDOMNodePtr ptrTestResponseNode = NULL;

	try
	{
		MSXML::IXMLDOMDocumentPtr ptrTestResponseDoc(__uuidof(MSXML::DOMDocument));
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
					LPCWSTR szBinary	= static_cast<IXMLAssistDOMNodePtr>(ptrRequestTestResponseNode)->getAttributeText(L"BINARY");
					LPCWSTR szXML		= static_cast<IXMLAssistDOMNodePtr>(ptrRequestTestResponseNode)->getAttributeText(L"XML");
					if (wcslen(szBinary) > 0)
					{
						// Test response contains the name of the binary response file as an attribute.
						const int nMaxSizeBuffer = 0xFFFF;
						BYTE Buffer[nMaxSizeBuffer];
						FILE* fp = _wfopen(szBinary, L"rb");
						int nResponseSize = fread(Buffer, sizeof(BYTE), sizeof(Buffer), fp);
						fclose(fp);
						ptrTestResponseNode = COptimusResponse2XML::Convert(Buffer, nResponseSize, pMetaDataEnv, pCodePage);
					}
					else if (wcslen(szXML) > 0)
					{
						// Test response contains the dummy response XML as an attribute.
						ptrTestResponseDoc->loadXML(szXML);
						ptrTestResponseNode = ptrTestResponseDoc->documentElement;
					}
				}
			}

			if (ptrTestResponseNode == NULL)
			{
				throw CException(E_GENERICERROR, __FILE__, __LINE__, L"Invalid TESTRESPONSE node");
			}
		}
		else if (pMetaDataEnv->GetTestResponse())
		{
			// This environment sends back test responses, and does not call Optimus.

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

void CRequestOptimus::WriteXML(MSXML::IXMLDOMNodePtr ptrNode, LPCTSTR pszType, int nMaxWriteXML, int nIncrement) const
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

