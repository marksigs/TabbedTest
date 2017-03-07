///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaDataEnvTransact.cpp
//	DESCRIPTION:	
//		ODI Transact environment specific meta data.
//
//		Each CMetaDataEnvTransact defines a specific Transact configuration,
//		e.g., Transact host and port, Transact object map, code pages etc. A  
//		single instance of ODIConverter can simultaneously support multiple  
//		Transact meta data environments. For each call to the Request method on the
//		ODIConverter.ODIConverter interface a different meta data environment
//		can be specified. Thus a single instance of ODIConverter can simultaneously
//		talk to multiple Transact hosts (as well as other hosts).
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		17/10/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "MetaData.h"
#include "MetaDataEnvTransact.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "Profiler.h"


static LPCWSTR g_pszTransactHostKey					= L"TRANSACTHOST";
static LPCWSTR g_pszTransactPortKey					= L"TRANSACTPORT";
static LPCWSTR g_pszSktTimeOutMsKey					= L"SKTTIMEOUTMS";
static LPCWSTR g_pszSktMaxSendKey					= L"SKTMAXSEND";
static LPCWSTR g_pszSktMaxRecvKey					= L"SKTMAXRECV";
static LPCWSTR g_pszProfileKey						= L"PROFILE";
static LPCWSTR g_pszTestResponseKey					= L"TESTRESPONSE";
static LPCWSTR g_pszTestResponseDelayMsKey			= L"TESTRESPONSEDELAYMS";
static LPCWSTR g_pszLogKey							= L"LOG";
static LPCWSTR g_pszLogLogDebugKey					= L"LOG/LOGDEBUG";
static LPCWSTR g_pszLogLogXMLKey					= L"LOG/LOGXML";
static LPCWSTR g_pszLogDebugKey						= L"LOGDEBUG";
static LPCWSTR g_pszLogXMLKey						= L"LOGXML";
static LPCWSTR g_pszDefault							= L"DEFAULT";
static LPCWSTR g_pszData							= L"DATA";

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnvTransact::CMetaDataEnvTransact 
//	 
//	Description:  
//		Constructor.
//	 
//	Parameters:  
//		LPCWSTR szType
//			The type name for this meta data environment; should be L"TRANSACT".
//		LPCWSTR szName:    
//			The unique name for this meta data environment.
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnvTransact::CMetaDataEnvTransact(LPCWSTR szType, LPCWSTR szName) : 
	CMetaDataEnv(szType, szName)
{
	m_bstrTransactHost			= L"localhost";
	m_nTransactPort				= 0;
	m_nSktTimeOutMs				= 0;	// By default, no time out.
	m_nSktMaxSend				= 1024;	// Default maximum size of send buffer.
	m_nSktMaxRecv				= 1024; // Default maximum size of recv buffer.
	m_bProfile					= false;
	m_bLogDebug					= false;
	m_nLogXML					= 0;
	m_bTestResponse				= false;
	m_nTestResponseDelayMs		= 0;
}

CMetaDataEnvTransact::~CMetaDataEnvTransact()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnvTransact::Init 
//	 
//	Description:  
//		Initialises a meta data environment from XML.   
//	 
//	Parameters:  
//		const IXMLDOMElementPtr ptrXMLDOMNodeParent:    
//			XML that contains the meta data. See ODIConverter.dtd for details.   
//	 
//	Return:  
//		bool:    
//			Returns true if successfully initialised. 
///////////////////////////////////////////////////////////////////////////////
bool CMetaDataEnvTransact::Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent)
{
	try
	{
		// Put a write lock on the meta data as we're going to be changing it.
		namespaceMutex::CWriteLock lckWriteLock(GetSharedLock());

		// Random sleep of 0 - 10 seconds useful for testing locking.
		//Sleep(((2 * rand() * 10000 + RAND_MAX) / RAND_MAX - 1) / 2);

		MSXML::IXMLDOMElementPtr ptrXMLDOMElement;
		_variant_t varAttribute;

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszTransactHostKey);
		if (ptrXMLDOMElement != NULL)
		{
			// m_bstrTransactHost is the IP Address/DNS name of the Transact server which
			// ODIConverter will connect to for sending requests.
			m_bstrTransactHost = ptrXMLDOMElement->Gettext();
		}
		if (wcslen(m_bstrTransactHost) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszTransactHostKey);
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszTransactPortKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nTransactPort = _wtoi(ptrXMLDOMElement->Gettext());
		}
		if (m_nTransactPort == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszTransactPortKey);
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszSktTimeOutMsKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nSktTimeOutMs = _wtoi(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszSktMaxSendKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nSktMaxSend = _wtoi(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszSktMaxRecvKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nSktMaxRecv = _wtoi(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszProfileKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bProfile = IsTrueString(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszLogLogDebugKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bLogDebug = IsTrueString(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszLogLogXMLKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nLogXML = _wtoi(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszTestResponseKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bTestResponse = IsTrueString(ptrXMLDOMElement->Gettext());
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszTestResponseDelayMsKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nTestResponseDelayMs = _wtoi(ptrXMLDOMElement->Gettext());
		}

		SetInitialised(true);
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

	return GetInitialised();
}

void CMetaDataEnvTransact::Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent)
{
	try
	{
		const IXMLAssistDOMNodePtr ptrNodeParent = ptrXMLDOMNodeParent;

		ptrNodeParent->setChildNodeText(g_pszTransactHostKey, m_bstrTransactHost, true);

		wchar_t szTransactPort[16] = L"";
		swprintf(szTransactPort, L"%d", m_nTransactPort);
		ptrNodeParent->setChildNodeText(g_pszTransactPortKey, szTransactPort, true);

		wchar_t szSktTimeOutMs[16] = L"";
		swprintf(szSktTimeOutMs, L"%d", m_nSktTimeOutMs);
		ptrNodeParent->setChildNodeText(g_pszSktTimeOutMsKey, szSktTimeOutMs, true);

		wchar_t szSktMaxSend[16] = L"";
		swprintf(szSktMaxSend, L"%d", m_nSktMaxSend);
		ptrNodeParent->setChildNodeText(g_pszSktMaxSendKey, szSktMaxSend, true);

		wchar_t szSktMaxRecv[16] = L"";
		swprintf(szSktMaxRecv, L"%d", m_nSktMaxRecv);
		ptrNodeParent->setChildNodeText(g_pszSktMaxRecvKey, szSktMaxRecv, true);

		ptrNodeParent->setChildNodeText(g_pszProfileKey, m_bProfile ? L"Y" : L"N", true);

		const IXMLAssistDOMNodePtr ptrNodeLog = ptrNodeParent->setChildNodeText(g_pszLogKey, NULL, true);
		ptrNodeLog->setChildNodeText(g_pszLogDebugKey, m_bLogDebug ? L"Y" : L"N", true);
		wchar_t szLogXML[16] = L"";
		swprintf(szLogXML, L"%d", m_nLogXML);
		ptrNodeLog->setChildNodeText(g_pszLogXMLKey, szLogXML, true);

		ptrNodeParent->setChildNodeText(g_pszTestResponseKey, m_bTestResponse ? L"Y" : L"N", true);

		wchar_t szTestResponseDelayMs[16] = L"";
		swprintf(szTestResponseDelayMs, L"%d", m_nTestResponseDelayMs);
		ptrNodeParent->setChildNodeText(g_pszTestResponseDelayMsKey, szTestResponseDelayMs, true);
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
}

