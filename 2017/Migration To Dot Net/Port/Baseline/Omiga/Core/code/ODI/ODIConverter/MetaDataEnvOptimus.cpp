///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaDataEnvOptimus.cpp
//	DESCRIPTION:	
//		ODI Optimus environment specific meta data.
//
//		Each CMetaDataEnvOptimus defines a specific Optimus configuration,
//		e.g., Optimus host and port, Optimus object map, code pages etc. A  
//		single instance of ODIConverter can simultaneously support multiple  
//		Optimus meta data environments. For each call to the Request method on the
//		ODIConverter.ODIConverter interface a different meta data environment
//		can be specified. Thus a single instance of ODIConverter can simultaneously
//		talk to multiple Optimus hosts (as well as other hosts).
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
#include "CodePage.h"
#include "Exception.h"
#include "MetaData.h"
#include "MetaDataEnvOptimus.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "OptimusResponse2XML.h"
#include "Profiler.h"


static LPCWSTR g_pszOptimusHostKey					= L"OPTIMUSHOST";
static LPCWSTR g_pszOptimusPortKey					= L"OPTIMUSPORT";
static LPCWSTR g_pszOptimusHostNameKey				= L"OPTIMUSHOSTNAME";
static LPCWSTR g_pszOptimusEnvironmentKey			= L"OPTIMUSENVIRONMENT";
static LPCWSTR g_pszListenPortKey					= L"LISTENPORT";
static LPCWSTR g_pszSktTimeOutMsKey					= L"SKTTIMEOUTMS";
static LPCWSTR g_pszSktMaxSendKey					= L"SKTMAXSEND";
static LPCWSTR g_pszSktMaxRecvKey					= L"SKTMAXRECV";
static LPCWSTR g_pszOptimusObjectMapKey				= L"OPTIMUSOBJECTMAP";
static LPCWSTR g_pszProfileKey						= L"PROFILE";
static LPCWSTR g_pszLogKey							= L"LOG";
static LPCWSTR g_pszLogLogDebugKey					= L"LOG/LOGDEBUG";
static LPCWSTR g_pszLogLogHexKey					= L"LOG/LOGHEX";
static LPCWSTR g_pszLogLogXMLKey					= L"LOG/LOGXML";
static LPCWSTR g_pszLogDebugKey						= L"LOGDEBUG";
static LPCWSTR g_pszLogHexKey						= L"LOGHEX";
static LPCWSTR g_pszLogXMLKey						= L"LOGXML";
static LPCWSTR g_pszTestResponseKey					= L"TESTRESPONSE";
static LPCWSTR g_pszTestResponseDelayMsKey			= L"TESTRESPONSEDELAYMS";
static LPCWSTR g_pszCodePageMapKey					= L"CODEPAGEMAP";
static LPCWSTR g_pszLookUpTableMapKey				= L"LOOKUPTABLEMAP";
static LPCWSTR g_pszDefault							= L"DEFAULT";
static LPCWSTR g_pszData							= L"DATA";
static LPCWSTR g_pszRequestCountryKey				= L"//REQUEST/SESSION/SESSIONIMPL/PREFERREDLOCALE/LOCALE/COUNTRY";
static LPCWSTR g_pszRequestLanguageKey				= L"//REQUEST/SESSION/SESSIONIMPL/PREFERREDLOCALE/LOCALE/LANGUAGE";

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnvOptimus::CMetaDataEnvOptimus 
//	 
//	Description:  
//		Constructor.
//	 
//	Parameters:  
//		LPCWSTR szType
//			The type name for this meta data environment; should be L"OPTIMUS".
//		LPCWSTR szName:    
//			The unique name for this meta data environment.
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnvOptimus::CMetaDataEnvOptimus(LPCWSTR szType, LPCWSTR szName) : 
	CMetaDataEnv(szType, szName)
{
	m_bstrOptimusHost			= L"localhost";
	m_bstrOptimusObjectMapPath	= L"";
	m_bstrOptimusHostName		= m_bstrOptimusHost;
	m_bstrOptimusEnvironment	= L"";
	m_bstrCodePageMapPath		= L"";
	m_bstrLookUpTablesPath		= L"";
	m_bstrNewCodePageMapPath	= L"";
	m_bstrNewLookUpTablesPath	= L"";
	m_nOptimusPort				= 0;
	m_nListenPort				= 0;
	m_nSktTimeOutMs				= 0;	// By default, no time out.
	m_nSktMaxSend				= 1024;	// Default maximum size of send buffer.
	m_nSktMaxRecv				= 1024; // Default maximum size of recv buffer.
	m_bProfile					= false;
	m_bLogDebug					= false;
	m_bLogHex					= false;
	m_nLogXML					= 0;
	m_bTestResponse				= false;
	m_nTestResponseDelayMs		= 0;
}

CMetaDataEnvOptimus::~CMetaDataEnvOptimus()
{
	// Causes listening socket thread to terminate.
	if (m_sktListener.IsValid())
	{
		m_sktListener.Close();
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnvOptimus::Init 
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
bool CMetaDataEnvOptimus::Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent)
{
	try
	{
		// Put a write lock on the meta data as we're going to be changing it.
		namespaceMutex::CWriteLock lckWriteLock(GetSharedLock());

		// Random sleep of 0 - 10 seconds useful for testing locking.
		//Sleep(((2 * rand() * 10000 + RAND_MAX) / RAND_MAX - 1) / 2);

		MSXML::IXMLDOMElementPtr ptrXMLDOMElement;
		_variant_t varAttribute;

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszOptimusHostKey);
		if (ptrXMLDOMElement != NULL)
		{
			// m_bstrOptimusHost is the IP Address/DNS name of the Optimus server which
			// ODIConverter will connect to for sending requests.
			// m_bstrOptimusHostName is the name of the Optimus server that will be used
			// in startNewSession requests that are sent to Optimus; it defaults to
			// the same as m_bstrOptimusHost.
			m_bstrOptimusHost = m_bstrOptimusHostName = ptrXMLDOMElement->Gettext();
		}
		if (wcslen(m_bstrOptimusHost) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszOptimusHostKey);
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszOptimusPortKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nOptimusPort = _wtoi(ptrXMLDOMElement->Gettext());
		}
		if (m_nOptimusPort == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszOptimusPortKey);
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszOptimusHostNameKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bstrOptimusHostName = ptrXMLDOMElement->Gettext();
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszOptimusEnvironmentKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bstrOptimusEnvironment = ptrXMLDOMElement->Gettext();
		}
		if (wcslen(m_bstrOptimusEnvironment) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszOptimusEnvironmentKey);
		}

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszCodePageMapKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bstrNewCodePageMapPath = _Module.MakeModulePath(ptrXMLDOMElement->Gettext());
			if (wcslen(m_bstrCodePageMapPath) == 0)
			{
				// Path is not already defined, so set to path in XML.
				m_bstrCodePageMapPath = m_bstrNewCodePageMapPath;
			}
			else if (wcslen(m_bstrNewCodePageMapPath) == 0)
			{
				// New path is not defined in XML, so set to existing path.
				m_bstrNewCodePageMapPath = m_bstrCodePageMapPath;
			}
			m_CodePages.Init(m_bstrCodePageMapPath);
		}
		if (wcslen(m_bstrCodePageMapPath) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszCodePageMapKey);
		}
		m_CodePages.InitNode(ptrXMLDOMNodeParent);

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszLookUpTableMapKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bstrNewLookUpTablesPath = _Module.MakeModulePath(ptrXMLDOMElement->Gettext());
			if (wcslen(m_bstrLookUpTablesPath) == 0)
			{
				// Path is not already defined, so set to path in XML.
				m_bstrLookUpTablesPath = m_bstrNewLookUpTablesPath;
			}
			else if (wcslen(m_bstrNewLookUpTablesPath) == 0)
			{
				// New path is not defined in XML, so set to existing path.
				m_bstrNewLookUpTablesPath = m_bstrLookUpTablesPath;
			}
			m_LookUpTables.Init(m_bstrLookUpTablesPath);
		}
		if (wcslen(m_bstrLookUpTablesPath) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszLookUpTableMapKey);
		}
		m_LookUpTables.InitNode(ptrXMLDOMNodeParent);

		InitOptimusObjectMapPath();
		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszOptimusObjectMapKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bstrOptimusObjectMapPath = _Module.MakeModulePath(ptrXMLDOMElement->Gettext());
			if (_taccess(m_bstrOptimusObjectMapPath, 00) != 0)
			{
				throw CException(E_FILEDOESNOTEXIST, __FILE__, __LINE__, _T("File does not exist: %s"), static_cast<LPWSTR>(m_bstrOptimusObjectMapPath));
			}
			else
			{	
				m_OptimusMetaData.Init(this);
			}
		}
		if (wcslen(m_bstrOptimusObjectMapPath) == 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszOptimusObjectMapKey);
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

		ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszLogLogHexKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bLogHex = IsTrueString(ptrXMLDOMElement->Gettext());
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

		InitListeningSocket(ptrXMLDOMNodeParent);

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

void CMetaDataEnvOptimus::Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent)
{
	try
	{
		const IXMLAssistDOMNodePtr ptrNodeParent = ptrXMLDOMNodeParent;

		ptrNodeParent->setChildNodeText(g_pszOptimusHostKey, m_bstrOptimusHost, true);

		wchar_t szOptimusPort[16] = L"";
		swprintf(szOptimusPort, L"%d", m_nOptimusPort);
		ptrNodeParent->setChildNodeText(g_pszOptimusPortKey, szOptimusPort, true);

		ptrNodeParent->setChildNodeText(g_pszOptimusHostNameKey, m_bstrOptimusHostName, true);
		ptrNodeParent->setChildNodeText(g_pszOptimusEnvironmentKey, m_bstrOptimusEnvironment, true);

		m_bstrCodePageMapPath = m_bstrNewCodePageMapPath;
		ptrNodeParent->setChildNodeText(g_pszCodePageMapKey, _Module.RemoveModulePath(m_bstrCodePageMapPath), true);
		m_CodePages.Save(m_bstrCodePageMapPath);

		m_bstrLookUpTablesPath = m_bstrNewLookUpTablesPath;
		ptrNodeParent->setChildNodeText(g_pszLookUpTableMapKey, _Module.RemoveModulePath(m_bstrLookUpTablesPath), true);
		m_LookUpTables.Save(m_bstrLookUpTablesPath);

		ptrNodeParent->setChildNodeText(g_pszOptimusObjectMapKey, _Module.RemoveModulePath(m_bstrOptimusObjectMapPath), true);

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
		ptrNodeLog->setChildNodeText(g_pszLogHexKey, m_bLogHex ? L"Y" : L"N", true);
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

void CMetaDataEnvOptimus::InitListeningSocket(const MSXML::IXMLDOMElementPtr ptrXMLDOMNodeParent)
{
	try
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMNodeParent->selectSingleNode(g_pszListenPortKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_nListenPort = _wtoi(ptrXMLDOMElement->Gettext());

			if (m_nListenPort != 0)
			{
				bool bSuccess = 
					m_sktListener.Open() && 
					m_sktListener.Bind(m_nListenPort) &&
					m_sktListener.Listen();

				if (bSuccess)
				{
					_Module.LogEventInfo(_T("Listening on port %d"), m_nListenPort);

					DWORD dwThreadId = 0;
					if (::CreateThread(NULL, 0, ListeningSocketThreadProc, this, 0, &dwThreadId) == NULL)
					{
						throw CException(E_GENERICERROR, __FILE__, __LINE__, _T("Unable to create listening socket thread"));
					}
				}
				else
				{
					throw CException(E_GENERICERROR, __FILE__, __LINE__, _T("Unable to create listening socket"));
				}
			}
			else
			{
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszListenPortKey);
			}
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
}

DWORD WINAPI CMetaDataEnvOptimus::ListeningSocketThreadProc(LPVOID lpParameter)
{
	try
	{
		CMetaDataEnvOptimus* pMetaDataEnvOptimus = reinterpret_cast<CMetaDataEnvOptimus*>(lpParameter);

		while (pMetaDataEnvOptimus != NULL && pMetaDataEnvOptimus->m_sktListener.IsValid())
		{
			CWSASocket* psktAccept = new CWSASocket;

			if (psktAccept != NULL && pMetaDataEnvOptimus->m_sktListener.Accept(*psktAccept))
			{
				// Incoming connection from Optimus. Spawn another thread to handle it, while
				// this thread continues to listen for incoming connections.

				PairAcceptDataType* pAcceptData = new PairAcceptDataType(pMetaDataEnvOptimus, psktAccept);
				DWORD dwThreadId = 0;

				if (::CreateThread(NULL, 0, AcceptSocketThreadProc, pAcceptData, 0, &dwThreadId) == NULL)
				{
					CException Exception(E_GENERICERROR, __FILE__, __LINE__, _T("Unable to create accept socket thread"));
					CODIConverter1::LogError(Exception);
				}
			}
		}
	}
	catch(...)
	{
	}

	return 0;
}

DWORD WINAPI CMetaDataEnvOptimus::AcceptSocketThreadProc(LPVOID lpParameter)
{
	bool bSuccess = false;

	PairAcceptDataType* pAcceptData = reinterpret_cast<PairAcceptDataType*>(lpParameter);

	_ASSERTE(pAcceptData != NULL);

	if (pAcceptData != NULL)
	{
		CMetaDataEnvOptimus* pMetaDataEnvOptimus	= pAcceptData->first;
		CWSASocket* psktAccept						= pAcceptData->second;

		_ASSERTE(pMetaDataEnvOptimus != NULL);
		_ASSERTE(psktAccept != NULL);

		if (pMetaDataEnvOptimus != NULL && psktAccept != NULL)
		{
			const int nMaxSizeBuffer = 0xFFFF;
			BYTE Buffer[nMaxSizeBuffer];
			int nResponseSize = 0;
			MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

			nResponseSize = psktAccept->Recv(reinterpret_cast<LPSTR>(Buffer), sizeof(Buffer));
			bSuccess = nResponseSize != SOCKET_ERROR && nResponseSize > 0;

			if (bSuccess)
			{
				ptrResponseNode = COptimusResponse2XML::Convert(Buffer, nResponseSize, pMetaDataEnvOptimus, pMetaDataEnvOptimus->m_CodePages.GetCodePage());
			}

			delete psktAccept;
		}

		delete pAcceptData;
	}

	return 0;
}


void CMetaDataEnvOptimus::InitOptimusObjectMapPath()
{
	try
	{
		m_bstrOptimusObjectMapPath = CMetaData::GetDefaultOptimusObjectMapPath();
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


CCodePage* CMetaDataEnvOptimus::GetCodePage(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode)
{
	CCodePage* pCodePage = NULL;

	try
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL;
		_variant_t varAttribute;
		_bstr_t bstrCountry	= L"";
		_bstr_t bstrLanguage = L"";

		ptrXMLDOMElement = ptrXMLDOMNode->selectSingleNode(g_pszRequestCountryKey);
		if (ptrXMLDOMElement != NULL && (varAttribute = ptrXMLDOMElement->getAttribute(g_pszData), varAttribute.vt != VT_NULL))
		{
			bstrCountry = varAttribute.bstrVal;
		}

		ptrXMLDOMElement = ptrXMLDOMNode->selectSingleNode(g_pszRequestLanguageKey);
		if (ptrXMLDOMElement != NULL && (varAttribute = ptrXMLDOMElement->getAttribute(g_pszData), varAttribute.vt != VT_NULL))
		{
			bstrLanguage = varAttribute.bstrVal;
		}

		if (wcslen(bstrCountry) == 0 || wcslen(bstrLanguage) == 0)
		{
			bstrCountry = g_pszDefault;
		}

		pCodePage = m_CodePages.GetCodePage(bstrCountry, bstrLanguage);
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

	return pCodePage;
}
