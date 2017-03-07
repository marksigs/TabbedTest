///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListener1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AD      30/08/00    Initial version
//  AD		Oct 2000	Change to include xml soft configuration support and
//						dynamic queue's
//  LD		05/10/01    SYS2295 - Replace Error message with a warning for absense of 
//						configuration file.   Enumerate error codes.
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//  ML		17/03/03	Use password repository (OmigaPlus)
//  RF		24/02/04	BMIDS727 - Different queues can use different tables
//						Make XML parsing more robust
//  PSC		11/07/06	CORE284 - Return Queue Type and Table Suffix from 
//						RequestAllQueueInfo and RequestQueueInfo
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MessageQueueListener.h"
#include "MessageQueueListener1.h"

using namespace namespaceMutex;

#define ANYTHREADLOG_THIS CLogIn LogIn(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogIn.AnyThreadInitialize
#define ANYTHREADLOG_THIS_INOUT CLogInOut LogInOut(*this, MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL); LogInOut.AnyThreadInitialize

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListener1


CMessageQueueListener1::CMessageQueueListener1()
{
}

CMessageQueueListener1::~CMessageQueueListener1()
{
}

HRESULT CMessageQueueListener1::FinalConstruct()
{
	return CComObjectRootEx<CComMultiThreadModel>::FinalConstruct();
}

void CMessageQueueListener1::FinalRelease()
{
	CComObjectRootEx<CComMultiThreadModel>::FinalRelease();
}

///////////////////////////////////////////////////////////////////////////////

STDMETHODIMP CMessageQueueListener1::Start()
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueListener1::Start\n"));

	try
	{
		// start up the scheduler
		m_ScheduleManagerDaily.StartUp();
	}
	catch(...)
	{
	}
	return S_OK;
}

STDMETHODIMP CMessageQueueListener1::Stop()
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueListener1::Stop\n"));

	try
	{
		m_ScheduleManagerDaily.CloseDown();

		// closedown each queue
		while (!m_MapQueueNameToCommon1.empty())
		{
			MapQueueNameToCommon1::iterator theCommonIterator = m_MapQueueNameToCommon1.begin();
			(*theCommonIterator).second->put_Started(false);
			(*theCommonIterator).second = NULL;
			m_MapQueueNameToCommon1.erase(theCommonIterator);
		}
	}
	catch(...)
	{
	}
	return S_OK;
}


STDMETHODIMP CMessageQueueListener1::Configure(BSTR in_xml, BSTR *out_xml)
{

	XmlNS::IXMLDOMDocumentPtr	ptrDomDocument(__uuidof(XmlNS::DOMDocument));	
	XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));	
	bstr_t bstrXml_Out;

	enum
	{
		eLoadXml,
		eGetRootElement,
		eParseDom,
		eRequest,
		eNULL,
	} eAction = eNULL;

    // PSC 06/08/03 MSMS214
    CSingleLock lckConfigure(&m_csConfigure);
	// CWriteLock lckConfigure(m_sharedlockConfigure);

	try
	{
	
		// put in some critical section code to serialise access here
		
		
		eAction = eLoadXml;
		// load in the XML
		_variant_t varBRes;
		
		varBRes = ptrDomDocument->loadXML(in_xml);

		if(varBRes.boolVal == false)
		{
			throw (_com_error(0,NULL));
		}

		eAction = eGetRootElement;
		
		// get the root element
		XmlNS::IXMLDOMElementPtr ptrElement;
		ptrElement = ptrDomDocument->GetdocumentElement();

		if(ptrElement == NULL)
		{
			throw (_com_error(0,NULL));
		}

		eAction = eParseDom;

		// get the ACTION info 
		bstr_t bstrElement1("ACTION");
		_variant_t varValue;
		varValue = ptrElement->getAttribute(bstrElement1);
	
		// create / update / get / delete

		if (varValue.vt == VT_NULL)
		{
			throw (_com_error(0,NULL));
		}

		eAction = eRequest;
		
		if (wcsicmp (varValue.bstrVal, L"CREATE") == 0)
		{
			bstrXml_Out = Create(ptrDomDocument);
		}
		else if (wcsicmp (varValue.bstrVal, L"UPDATE") == 0)
		{
			bstrXml_Out = Update(ptrDomDocument);
		}
		else if (wcsicmp (varValue.bstrVal, L"GET") == 0)
		{
			// get used for the stalled component info etc
			bstrXml_Out = Get(ptrDomDocument);
		}
		else if (wcsicmp (varValue.bstrVal, L"DELETE") == 0)
		{
			bstrXml_Out = Delete(ptrDomDocument);
		}
		else
		{
				// have undefined info coming in
			throw (_com_error(0,NULL));
		}
		
		// for test purposes - destroy m_ScheduleManagerDaily;
		//	m_ScheduleManagerDaily.CloseDown();
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eLoadXml:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOCREATEDOMDOCUMENT, 
					L"CMessageQueueListener1::Configure", 
					L"Unable to create Dom Document",comerr);
				break;
			case eGetRootElement:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOGETROOTELEMENT, 
					L"CMessageQueueListener1::Configure", 
					L"Unable to retrieve Dom root element",comerr);
				break;
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::Configure", 
					L"Unable to process tag",comerr);
				break;
			case eRequest:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::Configure", 
					L"Unable to complete request",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNDEFINEDERROR, 
					L"CMessageQueueListener1::Configure", 
					L"Undefined error",comerr);
				break;
		}

		XmlNS::IXMLDOMElementPtr ptrReturnElement;

		// get the root element
		ptrReturnElement = ptrDomOut->GetdocumentElement();

		bstrXml_Out = ptrReturnElement->xml;
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
			L"CMessageQueueListener1::Configure", 
			L"Undefined error",(_com_error)NULL);

		XmlNS::IXMLDOMElementPtr	ptrReturnElement;

		// get the root element
		ptrReturnElement = ptrDomOut->GetdocumentElement();	

		bstrXml_Out = ptrReturnElement->xml;	
	}

    // PSC 06/08/03 MSMS214
	lckConfigure.Unlock();
	*out_xml = bstrXml_Out;
	
	return S_OK;
}


STDMETHODIMP CMessageQueueListener1::LoadConfigurationDetails()
{
	ANYTHREADLOG_THIS_INOUT(_T("CMessageQueueListener1::LoadConfigurationDetails\n"));

	// look for file "MessageQueueListener.xml", read it and then call configure

	XmlNS::IXMLDOMDocumentPtr	ptrDomDocument;
	XmlNS::IXMLDOMDocumentPtr	ptrDomOut;

	HRESULT hr = S_OK;
	
	enum
	{
		eCreateDOM,
		eLoadXml,
		eGetRootElement,
		eParseDom,
		eNULL,
	} eAction = eNULL;

	
	try
	{
		eAction = eCreateDOM;

		ptrDomDocument.CreateInstance(__uuidof(XmlNS::DOMDocument));	
		ptrDomOut.CreateInstance(__uuidof(XmlNS::DOMDocument));	
		
		// load in the XML

		eAction = eLoadXml;

		TCHAR szModuleFileName[_MAX_PATH]; 
		GetModuleFileName(NULL, szModuleFileName, _MAX_PATH); 

#ifdef _UNICODE	
		USES_CONVERSION;
		string sModuleFileName = T2A(szModuleFileName);
#else
		string sModuleFileName = szModuleFileName;
#endif

		string sModulePath, sXmlFileName;
		string::size_type nPos = 0;
		if ((nPos = sModuleFileName.rfind('\\')) < sModuleFileName.npos)
		{
			sModulePath = sModuleFileName.substr(0, nPos);
			sXmlFileName = sModulePath + '\\' + "MessageQueueListener.xml";
		}

#ifdef _UNICODE	
		sprintf(T2A(szModuleFileName), "%s", sXmlFileName.c_str() );
#else
		sprintf(szModuleFileName, "%s", sXmlFileName.c_str() );
#endif

		FILE *fTest;

		fTest = fopen(sXmlFileName.c_str(), "r");

		if (fTest != NULL)
		{
			// if the start up file is present...
			fclose(fTest);
			
			_variant_t varBRes;
				
			varBRes = ptrDomDocument->load(sXmlFileName.c_str());
	
			if(varBRes.boolVal == false)
			{
				throw (_com_error(0,NULL));
			}
	
			eAction = eGetRootElement;
			XmlNS::IXMLDOMElementPtr ptrElement;
			ptrElement = ptrDomDocument->GetdocumentElement();

			if(ptrElement == NULL)
			{
				throw (_com_error(0,NULL));
			}
	
			eAction = eParseDom;
		
			// root element is CONFIGURATION
			bstr_t bstrElement1("REQUEST");
			XmlNS::IXMLDOMNodeListPtr ptrElementList;
			ptrElementList = ptrDomDocument->getElementsByTagName(bstrElement1);
			long lRequestList = ptrElementList->Getlength();
	
			// if we don't have any requests then that's not technically a problem.
			
			for(int iRequestLoop = 0; iRequestLoop<lRequestList; iRequestLoop++)
			{
				XmlNS::IXMLDOMNodePtr ptrNode;
				ptrNode=ptrElementList->Getitem(iRequestLoop);

				ptrDomOut->PutRefdocumentElement((XmlNS::IXMLDOMElementPtr)ptrNode);
			
				ptrElement = ptrDomOut->GetdocumentElement();

				bstr_t bstrElement2("ACTION");
			    _variant_t varValue;
				varValue = ptrElement->getAttribute(bstrElement2);
	
				if (varValue.vt == VT_NULL)
				{
					throw (_com_error(0,NULL));
				}
				bstr_t Xml_Out;
				if (varValue.vt == VT_BSTR)
			    {
					if (wcsicmp (varValue.bstrVal, L"CREATE") == 0)
					{
						Xml_Out = Create(ptrDomOut);
					}
					else if (wcsicmp (varValue.bstrVal, L"UPDATE") == 0)
					{
						Xml_Out = Update(ptrDomOut);
					}
					else if (wcsicmp (varValue.bstrVal, L"GET") == 0)
					{
						Xml_Out = Get(ptrDomOut);
					}
					else if (wcsicmp (varValue.bstrVal, L"DELETE") == 0)
					{
					}
					else	
					{
						// have undefined info coming in
						throw (_com_error(0,NULL));
					}
				}
			}
		}
		else
		{
			_Module.LogEventInfo(_T("CMessageQueueListener1::LoadConfigurationDetails - Configuration file not found. "));
		}

		// check to see if the output DOM is populated - if it isn't then
		// nothing's gone wrong and we can shove success back

		bstr_t bstrElement3("ERROR");
		XmlNS::IXMLDOMNodeListPtr	ptrElementList;
		ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
		long lOutList = ptrElementList->Getlength();

		if (lOutList != 0)
		{
			XmlNS::IXMLDOMElementPtr	ptrReturnElement;

			// get the root element
			ptrReturnElement = ptrDomOut->GetdocumentElement();
		
			bstr_t XmlOut(ptrReturnElement->xml);
			_Module.LogEventError(_T("CMessageQueueListener1::LoadConfigurationDetails - Unable to create Dom Document : %s"), (LPWSTR)ptrReturnElement->xml);
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eCreateDOM:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOCREATEDOMDOCUMENT, 
					L"CMessageQueueListener1::LoadConfigurationDetails ", 
					L"Unable to create Dom Document ",comerr);
				break;
			case eLoadXml:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOCREATEDOMDOCUMENT, 
					L"CMessageQueueListener1::LoadConfigurationDetails ", 
					L"Unable to load Dom Document ",comerr);
				break;
			case eGetRootElement:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOGETROOTELEMENT, 
					L"CMessageQueueListener1::LoadConfigurationDetails ", 
					L"Unable to retrieve Dom root element ",comerr);
				break;
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::LoadConfigurationDetails ", 
					L"Unable to process tag ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::LoadConfigurationDetails ", 
					L"Undefined error ",comerr);
				break;
		}

		hr = E_FAIL;
	
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", 
			ERROR_UNDEFINEDERROR, 
			L"CMessageQueueListener1::LoadConfigurationDetails ", 
			L"Undefined error ",(_com_error)NULL);
		_Module.LogEventError(_T("CMessageQueueListener1::LoadConfigurationDetails - Undefined error "));
		hr = E_FAIL;
	}
	
	return hr;
}


bstr_t CMessageQueueListener1::Create(XmlNS::IXMLDOMDocument *ptrDomIn)
{
	// unpack the DOM to find which thread to create etc
	// <Queuelist> element contains <Queue> items

	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	XmlNS::IXMLDOMNodePtr		ptrNode;

	XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));

	enum
	{
		eParseDom,
		eNULL,
	} eAction = eNULL;

	try
	{
		eAction = eParseDom;
		bstr_t bstrElement1("QUEUELIST");
		ptrElementList = ptrDomIn->getElementsByTagName(bstrElement1);
		long lQueueListList = ptrElementList->Getlength();

		// hopefully should have at least 1

		if (lQueueListList==0)
		{
			throw (_com_error(0,NULL));
		}

		ptrNode=ptrElementList->Getitem(0);

		// if node is ok then grab <QUEUE> 
		
		XmlNS::IXMLDOMNodeListPtr	ptrQueueElementList;
		ptrNode->get_childNodes(&ptrQueueElementList);
	
		long lQueueList = ptrQueueElementList->Getlength();

		if (lQueueList==0)
		{
			throw (_com_error(0,NULL));
		}
		
		// if we have some loop around and get info out.

		for(int iQueueLoop = 0; iQueueLoop<lQueueList; iQueueLoop++)
		{
			// unpack child
			XmlNS::IXMLDOMNodePtr ptrQueueNode;
			ptrQueueElementList->get_item(iQueueLoop, &ptrQueueNode);
			CreateQueue(ptrQueueNode, ptrDomOut); 
		}
	}

	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNABLETOPROCESSTAG, L"CMessageQueueListener1::Create ", L"Unable to process tag ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", 
					ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Create ", L"Undefined error ",comerr);
				break;
		}
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", 
			ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Create ", L"Undefined error ",(_com_error)NULL);
	}
	
	// check to see if the output DOM is populated - if it isn't then
	// nothing's gone wrong and we can shove success back
	
	bstr_t bstrElement3("ERROR");
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();

	if (lOutList == 0)
	{
		CreateErrorInfo(ptrDomOut, L"SUCCESS", NULL, NULL, NULL,(_com_error)NULL);
	}
	
	XmlNS::IXMLDOMElementPtr	ptrReturnElement;

	// get the root element
    ptrReturnElement = ptrDomOut->GetdocumentElement();
	
	bstr_t XmlOut(ptrReturnElement->xml);
	return XmlOut;
}

void CMessageQueueListener1::CreateQueue(
	const XmlNS::IXMLDOMNodePtr ptrQueueNode,
	XmlNS::IXMLDOMDocumentPtr ptrDomOut 
)
{
	/*
	
	RF 24/02/04 Added function to make code more robust and to include optional table suffix
	as configuration item

	ptrQueueNode contains XML in the following format:
	<QUEUE>
		<NAME></NAME>
		<TYPE></TYPE>							
		<CONNECTIONSTRING></CONNECTIONSTRING>
		<POLLINTERVAL></POLLINTERVAL>
		<TABLESUFFIX></TABLESUFFIX>		(Optional element)
		<THREADSLIST>
			<THREADS>
				<NUMBER></NUMBER>
			</THREADS>
		</THREADSLIST>
		</QUEUE>

	ptrDomOut is used to hold error info (if any)

	*/

	QueueTypeEnum eQueueType;

	enum
	{
		eGetQueueName,
		eGetQueueType,
		eGetConnStr,
		eGetPollInterval,
		eGetTableSuffix,
		eGetThreadsList,
		eNULL,
	} eAction = eNULL;

	try
	{
		XmlNS::IXMLDOMNodePtr ptrNodeTemp;
		bstr_t bstrSearchString;
			
		// get queue name (a mandatory configuration item)

		eAction = eGetQueueName;

		//bstrSearchString = ".//QUEUE/NAME";
		bstrSearchString = ".//NAME";
		ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
			
		if (ptrNodeTemp == NULL)
		{
			throw (_com_error(0,NULL));
		}

		_bstr_t bstrName = ptrNodeTemp->Gettext(); 
		if (bstrName.length() == 0)
		{
			throw (_com_error(0,NULL));
		}
				
		// get queue type (a mandatory configuration item)				

		eAction = eGetQueueType;

		bstrSearchString = ".//TYPE";
		ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
			
		if (ptrNodeTemp == NULL)
		{
			throw (_com_error(0,NULL));
		}

		_bstr_t bstrType = ptrNodeTemp->Gettext(); 
		if (bstrType.length() == 0)
		{
			throw (_com_error(0,NULL));
		}
		
		eQueueType = GeteQueueType(bstrType);

		// get extra info if this is an OM queue - get db conn string (a BSTR) 
		// and polling interval (a DWORD) and TABLESUFFIX (a BSTR)

		_bstr_t bstrConnString;
		DWORD dwPollInterval;
		_bstr_t bstrTableSuffix;

		if (eQueueType == eQueueTypeOMMQ1)
		{
			// get connection string 

			eAction = eGetConnStr;

			bstrSearchString = ".//CONNECTIONSTRING";
			ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
				
			if (ptrNodeTemp == NULL)
			{
				throw (_com_error(0,NULL));
			}

			_bstr_t bstrTemp = ptrNodeTemp->Gettext(); 
		
			// OPC6596 Optionally get the connection details from the registry and/or repository
			bstrConnString = LookupConnectionString(bstrTemp);
			
			// get polling interval (a mandatory configuration item)				

			eAction = eGetPollInterval;

			bstrSearchString = ".//POLLINTERVAL";
			ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
				
			if (ptrNodeTemp == NULL)
			{
				throw (_com_error(0,NULL));
			}

			bstrTemp = ptrNodeTemp->Gettext(); 
			if (bstrTemp.length() == 0)
			{
				throw (_com_error(0,NULL));
			}

			dwPollInterval = _wtol(bstrTemp);

			// get table suffix (an optional configuration item)				

			eAction = eGetTableSuffix;

			bstrSearchString = ".//TABLESUFFIX";
			ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
				
			if (ptrNodeTemp != NULL)
			{
				bstrTableSuffix = ptrNodeTemp->Gettext(); 
			}
		}
		
		// get the threads list 

		eAction = eGetThreadsList;

		bstrSearchString = ".//THREADSLIST";
		ptrNodeTemp = ptrQueueNode->selectSingleNode(bstrSearchString);   
		
		// loop thru the thread list
			
		XmlNS::IXMLDOMNodeListPtr ptrThreadElementList;
		ptrNodeTemp->get_childNodes(&ptrThreadElementList);

		long lThreadList = ptrThreadElementList->Getlength();

		if(lThreadList == 0)
		{
			throw (_com_error(0,NULL));
		}

		for(int iThreadLoop = 0; iThreadLoop<lThreadList; iThreadLoop++)
		{
			// action each thread
			CreateUnpackThreadInfo(
				ptrThreadElementList->Getitem(iThreadLoop), 
				eQueueType, 
				bstrName, 
				bstrConnString, 
				dwPollInterval,
				bstrTableSuffix,
				ptrDomOut);
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eGetQueueName:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process queue name tag ", comerr);
				break;
			case eGetQueueType:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process queue type tag ", comerr);
				break;
			case eGetConnStr:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process connection string tag ", comerr);
				break;
			case eGetPollInterval:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process polling interval tag ", comerr);
				break;
			case eGetTableSuffix:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process table suffix tag ", comerr);
				break;
			case eGetThreadsList:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Unable to process thread list tag ", comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
					L"CMessageQueueListener1::CreateQueue ", 
					L"Undefined error ", comerr);
				break;
		}
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
			L"CMessageQueueListener1::CreateQueue ", 
			L"Undefined error ",(_com_error)NULL);
	}
}

bstr_t CMessageQueueListener1::Update(XmlNS::IXMLDOMDocument *ptrDomIn)
{
	// unpack the DOM to find which thread to update etc
	// <Queuelist> element contains <Queue> items


	// for each queue there can be a paticular task
	// start queue / stop queue / restart stalled components / restart stalled queue

	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	XmlNS::IXMLDOMNodeListPtr	ptrQueueElementList;
	QueueTypeEnum eQueueType;

	XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));
	
	enum
	{
		eParseDom,
		eNULL,
	} eAction = eNULL;

	try
	{
		// find queuelist element

		eAction = eParseDom;
		bstr_t bstrElement1("QUEUELIST");
		ptrElementList = ptrDomIn->getElementsByTagName(bstrElement1);
		long lQueueListList = ptrElementList->Getlength();

		// hopefully should have at least 1

		if(lQueueListList==0) 
		{
			throw (_com_error(0,NULL));
		}
			
		XmlNS::IXMLDOMNodePtr ptrNode;
		ptrNode=ptrElementList->Getitem(0);

		// if node is ok then grab <QUEUE> 
		
		ptrNode->get_childNodes(&ptrQueueElementList);
	
		long lQueueList = ptrQueueElementList->Getlength();

		// if we have some loop around and get info out.
		if(lQueueList==0) 
		{
			throw (_com_error(0,NULL));
		}

		for(int iQueueLoop = 0; iQueueLoop<lQueueList; iQueueLoop++)
		{
			// unpack child
					
			XmlNS::IXMLDOMNodePtr ptrThreadListNode;
			ptrThreadListNode = ptrQueueElementList->Getitem(iQueueLoop);

				
			XmlNS::IXMLDOMNodeListPtr ptrQueueInfoList;
			ptrThreadListNode->get_childNodes(&ptrQueueInfoList);
					
			XmlNS::IXMLDOMNodePtr ptrThreadNode;
			ptrQueueInfoList->get_item(0, &ptrThreadNode);
			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
				
			// first child is the queue name
			bstr_t bstrName;
			bstr_t bstrType;
		
			bstrName = ptrThreadNode->Gettext();

			// then get the queue type
				
			ptrQueueInfoList->get_item(1, &ptrThreadNode);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			bstrType = ptrThreadNode->Gettext();
			eQueueType = GeteQueueType(bstrType);

			// then you have the task

			ptrQueueInfoList->get_item(2, &ptrThreadNode);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
				
			bstrType = ptrThreadNode->Gettext();

			// process the task
			if (wcsicmp (bstrType, L"START") == 0)
			{
				StartQueue(bstrName, eQueueType, ptrDomOut);
			}
			else if (wcsicmp (bstrType, L"STOP") == 0)
			{
				StopQueue(bstrName, eQueueType, ptrDomOut);
			}
			else if (wcsicmp (bstrType, L"RESTARTQUEUE") == 0)
			{
				ReStartQueue(bstrName, eQueueType, ptrDomOut);
			}
			else if (wcsicmp (bstrType, L"RESTARTCOMPONENT") == 0 || wcsicmp (bstrType, L"ADDSTALLEDCOMPONENTS") == 0)
			{
				// turn the list of components into a variant array of strings - _bstr_t
				// run thru number of components - will be items 3 - n

				// list is formatted as a variant that contains a safearray of BSTR's - the progid's	

				// list may be 1 or many components to be restarted - this is different to the original
				// way that this function is specced. What is in the array will be restarted as opposed to 
				// what is left out.

				bstr_t bstrElement2("COMPONENTLIST");
				XmlNS::IXMLDOMNodeListPtr	ptrComponentList;
				ptrComponentList = ptrDomIn->getElementsByTagName(bstrElement2);
				long lComponentListList = ptrComponentList->Getlength();

				// hopefully should have at least 1

				if(lComponentListList==0) 
				{
					throw (_com_error(0,NULL));
				}

				// build up list. If no elements then no probs

				XmlNS::IXMLDOMNodePtr ptrRestartNode;
				ptrRestartNode=ptrComponentList->Getitem(0);

				// if node is ok then grab <QUEUE> 
		
				XmlNS::IXMLDOMNodeListPtr	ptrComponentElementList;
				ptrRestartNode->get_childNodes(&ptrComponentElementList);
		
				long lComponentList = ptrComponentElementList->Getlength();

				// if we have some loop around and get info out.

				VARIANT pvarComponents;
				VariantInit(&pvarComponents);
				pvarComponents.vt = VT_ARRAY | VT_BSTR;
				SAFEARRAYBOUND  rgsabound[1];
				rgsabound[0].cElements = lComponentList;
				rgsabound[0].lLbound = 0;
				pvarComponents.parray = SafeArrayCreate(VT_BSTR, 1, rgsabound);

				if(pvarComponents.parray != NULL)
				{	
				    BSTR* pbstr = NULL;
				    SafeArrayAccessData(pvarComponents.parray, reinterpret_cast<void**>(&pbstr));
				
					for(int iComponentLoop = 0; iComponentLoop<lComponentList; iComponentLoop++)
					{
				
						ptrComponentElementList->get_item(iComponentLoop, &ptrRestartNode);
						if (ptrRestartNode == NULL)
						{
							throw (_com_error(0,NULL));
						}
						
					
						// first child is the queue name
						bstr_t bstrComponentName;
				
						bstrComponentName = ptrRestartNode->Gettext();
						*pbstr++ = bstrComponentName.copy();
					}

					SafeArrayUnaccessData(pvarComponents.parray);
				}

				if (wcsicmp (bstrType, L"RESTARTCOMPONENT") == 0)
				{
					ReStartComponents(bstrName, eQueueType, ptrDomOut, pvarComponents);
				}
				else
				{
					AddStalledComponents(bstrName, eQueueType, ptrDomOut, pvarComponents);
				}

				// clear variant
			}
			else if (wcsicmp (bstrType, L"THREADS") == 0)
			{
				// change the thread pool size on the queues
				ptrQueueInfoList->get_item(3, &ptrThreadNode);
				if (ptrThreadNode == NULL)
				{
					throw (_com_error(0,NULL));
				}

				bstr_t bstrNumber;
				bstrNumber = ptrThreadNode->Gettext();
				int nThreads = _wtoi(bstrNumber);
				ChangeQueueThreads(bstrName, eQueueType, ptrDomOut, nThreads);

			}
			else if (wcsicmp (bstrType, L"REMOVEEVENT") == 0)
			{
				// remove specific events from the schedule manager
				// key is passed in (a guid) - we find it and remove it.

				// loop per queue
				// loop around each event

				bstr_t bstrElement2("EVENT");
				XmlNS::IXMLDOMNodeListPtr	ptrEventList;
				ptrEventList = ptrDomIn->getElementsByTagName(bstrElement2);
				long lEventList = ptrEventList->Getlength();

				XmlNS::IXMLDOMNodePtr ptrEventNode;

				for(int iEventLoop = 0; iEventLoop<lEventList; iEventLoop++)
				{
					ptrEventList->get_item(iEventLoop, &ptrEventNode);

					bstr_t bstrKey;
					bstrKey = ptrEventNode->Gettext();
					
					unsigned int uiLength = bstrKey.length();
					
					if (uiLength != 38)
					{
						// guid is missing or not formed properly
						throw (_com_error(0,NULL));
					}
					
					// convert to a guid

					GUID gKey;

					GUIDFromString(bstrKey, &gKey);


					
					RemoveEvent(bstrName, eQueueType, ptrDomOut, gKey);
				}
			}
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, L"CMessageQueueListener1::Update ", L"Unable to process tag ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Update ", L"Undefined error ",comerr);
				break;
		}
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Update ", L"Undefined error ",(_com_error)NULL);
	}
	
	bstr_t bstrElement3("ERROR");
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();

	if (lOutList == 0)
	{
		CreateErrorInfo(ptrDomOut, L"SUCCESS", NULL, NULL, NULL,(_com_error)NULL);
	}
	
	
	XmlNS::IXMLDOMElementPtr	ptrReturnElement;

	// get the root element
    ptrReturnElement = ptrDomOut->GetdocumentElement();
	
	
	bstr_t XmlOut(ptrReturnElement->xml);
	return XmlOut;
}


bstr_t CMessageQueueListener1::Get(XmlNS::IXMLDOMDocument *ptrDomIn)
{

	// unpack the get request - we get queue name and that's it

	// get back thread pool size, number of items, status etc. 
	// List of stalled components if any...
	
	// Also - if there are no queue names then we get all of the 
	// information on all of the queues that are available
	
	bstr_t XmlOut;
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));	
			
	enum
	{
		eParseDom,
		eNULL,
	} eAction = eNULL;

	try
	{
		eAction = eParseDom;
		bstr_t bstrElement1("QUEUELIST");

		ptrElementList = ptrDomIn->getElementsByTagName(bstrElement1);
		long lQueueListList = ptrElementList->Getlength();
		QueueTypeEnum eQueueType;

		// hopefully should have at least 1

		if(lQueueListList==0) 
		{
			throw (_com_error(0,NULL));
		}
	
		XmlNS::IXMLDOMNodePtr ptrRestartNode;
		ptrRestartNode=ptrElementList->Getitem(0);

		// if node is ok then grab <QUEUE> 
		
		XmlNS::IXMLDOMNodeListPtr	ptrComponentElementList;
		ptrRestartNode->get_childNodes(&ptrComponentElementList);
		
		long lQueueList = ptrComponentElementList->Getlength();

		// loop around the number of queues
	
		// if we have some loop around and get info out.
		if(lQueueList==0) 
		{
			// if we get here then do all queues

			// call routine to request info -
			XmlOut = RequestAllQueueInfo(ptrDomOut);
		}
		else
		{
			for(int iQueueLoop = 0; iQueueLoop<lQueueList; iQueueLoop++)
			{
				XmlNS::IXMLDOMNodePtr ptrThreadListNode;
				ptrThreadListNode = ptrComponentElementList->Getitem(0);

		
				XmlNS::IXMLDOMNodeListPtr ptrQueueInfoList;
				ptrThreadListNode->get_childNodes(&ptrQueueInfoList);
					
				ptrQueueInfoList->get_item(0, &ptrRestartNode);
				if (ptrRestartNode == NULL)
				{	
					throw (_com_error(0,NULL));
				}
	
				bstr_t bstrQueueName;
				bstr_t bstrType;
				
				bstrQueueName = ptrRestartNode->Gettext();

		
				ptrQueueInfoList->get_item(1, &ptrRestartNode);

				if (ptrRestartNode == NULL)
				{	
					throw (_com_error(0,NULL));
				}
				bstrType = ptrRestartNode->Gettext();
				eQueueType = GeteQueueType(bstrType);

			
				// call routine to request info -
				XmlOut = RequestQueueInfo(bstrQueueName, eQueueType, ptrDomOut);

			}
		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, L"CMessageQueueListener1::Get ", L"Unable to process tag ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Get ", L"Undefined error ",comerr);
				break;
		}
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Get ", L"Undefined error ",(_com_error)NULL);
	}
	
	
	bstr_t bstrElement3("ERROR");
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();

	if (lOutList != 0)
	{
		XmlNS::IXMLDOMElementPtr	ptrReturnElement;

		// get the root element
		ptrReturnElement = ptrDomOut->GetdocumentElement();
	
	
		XmlOut=ptrReturnElement->xml;
	}

	return XmlOut;
}

#define SafeRelease(pUnk) {if (pUnk){pUnk -> Release();pUnk = NULL; }}  

bstr_t CMessageQueueListener1::Delete(XmlNS::IXMLDOMDocument *ptrDomIn)
{

	// remove a queue from the listener - reverse of create
	bstr_t XmlOut;
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	XmlNS::IXMLDOMDocumentPtr	ptrDomOut(__uuidof(XmlNS::DOMDocument));	
			
	enum
	{
		eParseDom,
		eNULL,
	} eAction = eNULL;


	try
	{
	
		eAction = eParseDom;
		bstr_t bstrElement1("QUEUELIST");

		ptrElementList = ptrDomIn->getElementsByTagName(bstrElement1);
		long lQueueListList = ptrElementList->Getlength();
		QueueTypeEnum eQueueType;

		// hopefully should have at least 1

		if(lQueueListList==0) 
		{
			throw (_com_error(0,NULL));
		}
	

		XmlNS::IXMLDOMNodePtr ptrRestartNode;
		ptrRestartNode=ptrElementList->Getitem(0);

		// if node is ok then grab <QUEUE> 
		
		XmlNS::IXMLDOMNodeListPtr	ptrComponentElementList;
		ptrRestartNode->get_childNodes(&ptrComponentElementList);
		
		long lQueueList = ptrComponentElementList->Getlength();

		// loop around the number of queues
	
		// if we have some loop around and get info out.
		if(lQueueList==0) 
		{
			throw (_com_error(0,NULL));
		}

		for(int iQueueLoop = 0; iQueueLoop<lQueueList; iQueueLoop++)
		{
	

			XmlNS::IXMLDOMNodePtr ptrThreadListNode;
			ptrThreadListNode = ptrComponentElementList->Getitem(0);

		
			XmlNS::IXMLDOMNodeListPtr ptrQueueInfoList;
			ptrThreadListNode->get_childNodes(&ptrQueueInfoList);
					
			ptrQueueInfoList->get_item(0, &ptrRestartNode);
			if (ptrRestartNode == NULL)
			{	
				throw (_com_error(0,NULL));
			}
	
			bstr_t bstrQueueName;
			bstr_t bstrType;
				
			bstrQueueName = ptrRestartNode->Gettext();
		
			ptrQueueInfoList->get_item(1, &ptrRestartNode);

			if (ptrRestartNode == NULL)
			{	
				throw (_com_error(0,NULL));
			}
			bstrType = ptrRestartNode->Gettext();
			eQueueType = GeteQueueType(bstrType);			

			// call routine to remove queue -
			MapQueueNameToCommon1::iterator theCommonIterator = m_MapQueueNameToCommon1.find(bstrQueueName);
			if (theCommonIterator != m_MapQueueNameToCommon1.end() )
			{
				
				// remove the scheduled events associated with the queue
				m_ScheduleManagerDaily.RemoveQueueEvents((*theCommonIterator).second);
				
				// assign and delete the memory

				IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
				ThreadPoolManagerCommon1Ptr = (*theCommonIterator).second;
				m_MapQueueNameToCommon1.erase(theCommonIterator);
				SafeRelease(ThreadPoolManagerCommon1Ptr);  
			}


		}
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{

			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, L"CMessageQueueListener1::Get ", L"Unable to process tag ",comerr);
				break;
			
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Get ", L"Undefined error ",comerr);
				break;
		}

	
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::Get ", L"Undefined error ",(_com_error)NULL);
	}
	
	
	bstr_t bstrElement3("ERROR");
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();


	if (lOutList == 0)
	{
		CreateErrorInfo(ptrDomOut, L"SUCCESS", NULL, NULL, NULL,(_com_error)NULL);
	}

	
	XmlNS::IXMLDOMElementPtr	ptrReturnElement;

		// get the root element
	ptrReturnElement = ptrDomOut->GetdocumentElement();
	
	
	XmlOut=ptrReturnElement->xml;
	

	return XmlOut;
}


void CMessageQueueListener1::CreateUnpackThreadInfo(
	XmlNS::IXMLDOMNode *ptrNode, 
	QueueTypeEnum eQueueType, 
	BSTR bstrQueueName, 
	BSTR bstrConnString, 
	DWORD dwPollInterval,
	BSTR bstrTableSuffix,
	XmlNS::IXMLDOMDocument *ptrDomOut) // error info, if any
{
	/* RF 24/02/04 Added bstrTableSuffix */ 

	HRESULT hRes;

	XmlNS::IXMLDOMNodePtr		ptrNameNode;
	XmlNS::IXMLDOMNodeListPtr	ptrThreadElementList;
	XmlNS::IXMLDOMNodePtr		ptrThreadNode;

	int nThreads;
	
	enum
	{
		eParseDom,
		eStartThread,
		eSearchQueue,
		eAddEvent,
		eNULL,
	} eAction = eNULL;

	try
	{
		eAction = eParseDom;

		// get the initial setting of how many threads the queue will have

		bstr_t bstrnodestring = ptrNode->Gettext();
		ptrNode->get_childNodes(&ptrThreadElementList);

		long lThreadList = ptrThreadElementList->Getlength();

		if (lThreadList == 0)
		{
			throw (_com_error(0,NULL));
		}

		if (lThreadList == 1)
		{
			// add item to the appropriate map 

			ptrThreadNode = ptrThreadElementList->Getitem(0);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			bstrnodestring = ptrThreadNode->Gettext();
			nThreads = _wtoi(bstrnodestring);

			GUID uuidofQueueToCreate;
			switch (eQueueType)
			{
				case eQueueTypeMSMQ1:
					uuidofQueueToCreate = __uuidof(ThreadPoolManagerMSMQ1);
					break;
				case eQueueTypeOMMQ1:
					uuidofQueueToCreate = __uuidof(ThreadPoolManagerOMMQ1);
					break;
				default :
					_ASSERTE(0); // should not reach here 
					break;
			}

			eAction = eStartThread;
			IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr(uuidofQueueToCreate);
			ThreadPoolManagerCommon1Ptr->put_NumberOfThreads(nThreads);
			ThreadPoolManagerCommon1Ptr->put_QueueName(_bstr_t(bstrQueueName));
		
			// if it's an OM queue then set the connection string and polling interval
			if (eQueueType == eQueueTypeOMMQ1)
			{
				ThreadPoolManagerCommon1Ptr->put_dwNextWaitIntervalms(dwPollInterval);
				ThreadPoolManagerCommon1Ptr->put_ConnectionString(_bstr_t(bstrConnString));
				ThreadPoolManagerCommon1Ptr->put_TableSuffix(_bstr_t(bstrTableSuffix));
			}
			
			hRes = ThreadPoolManagerCommon1Ptr->put_Started(true);
			if (hRes == 1)
			{
				// throw here
				throw (_com_error(0,NULL));
			}

			switch (eQueueType)
			{
				case eQueueTypeMSMQ1:
				case eQueueTypeOMMQ1:
					m_MapQueueNameToCommon1.insert(
						MapQueueNameToCommon1::value_type(
							bstrQueueName, ThreadPoolManagerCommon1Ptr));
					break;
				default :
					_ASSERTE(0); // should not reach here 
					break;
			}
		}
		else
		{
			// thread already created so add a function to it.
			int nHour, nMinute;
			DayOfWeek eDay;
			wchar_t * token;

			// unpack event info
			// number / day / time
				
			ptrThreadNode = ptrThreadElementList->Getitem(0);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			bstrnodestring = ptrThreadNode->Gettext();
			nThreads = _wtoi(bstrnodestring);

			// convert day to enum
			ptrThreadNode = ptrThreadElementList->Getitem(1);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			bstrnodestring = ptrThreadNode->Gettext();
	
			if (wcsicmp (bstrnodestring, L"SUNDAY") == 0)
				eDay = eSunday;
			else if  (wcsicmp (bstrnodestring, L"MONDAY") == 0)
				eDay = eMonday;
			else if  (wcsicmp (bstrnodestring, L"TUESDAY") == 0)
				eDay = eTuesday;
			else if  (wcsicmp (bstrnodestring, L"WEDNESDAY") == 0)
				eDay = eWednesday;
			else if  (wcsicmp (bstrnodestring, L"THURSDAY") == 0)
				eDay = eThursday;
			else if  (wcsicmp (bstrnodestring, L"FRIDAY") == 0)
				eDay = eFriday;
			else if  (wcsicmp (bstrnodestring, L"SATURDAY") == 0)
				eDay = eSaturday;			

			// work out the time
			ptrThreadNode = ptrThreadElementList->Getitem(2);

			if (ptrThreadNode == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			bstrnodestring = ptrThreadNode->Gettext();

			token = wcstok(bstrnodestring, L":");
			nHour = _wtoi(token);

			token = wcstok(NULL, L":");
			nMinute = _wtoi(token);
	
			// retrieve thread pool info so can pass on the pool info to the
			// MQEvent

		    MapQueueNameToCommon1::iterator theCommonIterator;
	        theCommonIterator = m_MapQueueNameToCommon1.find(bstrQueueName);

			eAction = eSearchQueue;
			if (theCommonIterator == m_MapQueueNameToCommon1.end() )
			{
				// haven't found the queue we're looking for
				throw (_com_error(0,NULL));
			}
		
			eAction = eAddEvent;
			if (!m_ScheduleManagerDaily.AddEventToList(
				new CMQEventChangeThreads((*theCommonIterator).second, nThreads), nHour, nMinute, eDay))
			{
				throw (_com_error(0,NULL));
			}
		}
	}

	catch(_com_error comerr)
	{
		switch (eAction)
		{
			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
					L"Unable to process tag ",comerr);
				break;

			case eStartThread:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, 
					L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
					L"Unable to start queue ",comerr);
				break;
			
			case eSearchQueue:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_SEARCHOFQUEUEFAILED, 
					L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
					L"Unable to find queue in the map ",comerr);
				break;
			
			case eAddEvent:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_ADDEVENTTOQUEUEFAILED, 
					L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
					L"Unable to add event to the schedule manager ",comerr);
				break;
				
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
					L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
					L"Undefined error ",comerr);
				break;
		}
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
			L"CMessageQueueListener1::CreateUnpackThreadInfo ", 
			L"Undefined error ",(_com_error)NULL);
	}
}


void CMessageQueueListener1::StartQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{

		// check to see if the queue's already going

		BOOL bStarted;
		ThreadPoolManagerCommon1Ptr->get_Started(&bStarted);

		if (!bStarted)
		{
		
			HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_Started(true);
			if (!SUCCEEDED(hRes))
			{
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::StartQueue ", L"Unable to start queue ",(_com_error)NULL);
			}
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::StartQueue ", L"Unable to find named queue ",(_com_error)NULL);
	}
}


void CMessageQueueListener1::StopQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{

		// check to see if the queue's going otherwise no point in stopping it
		
		BOOL bStarted;
		ThreadPoolManagerCommon1Ptr->get_Started(&bStarted);

		if (bStarted)
		{
			HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_Started(false);
			if (!SUCCEEDED(hRes))
			{
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::StopQueue ", L"Unable to stop queue ",(_com_error)NULL);
			}
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::StopQueue ", L"Unable to find named queue ",(_com_error)NULL);
	}
}


// AD 25/09/00 Given the queue name it will restart a queue that has stalled
void CMessageQueueListener1::ReStartQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{

		// check to see if the queue is actually stalled
		


		BOOL bStalled;
		ThreadPoolManagerCommon1Ptr->get_QueueStalled(&bStalled);

		if (bStalled)
		{

			HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_QueueStalled(false);
			if (!SUCCEEDED(hRes))
			{
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::ReStartQueue ", L"Unable to restart queue ",(_com_error)NULL);
			}
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::ReStartQueue ", L"Unable to find named queue ",(_com_error)NULL);
	}
	
}


// AD 25/09/00 A list of components is passed in - if a component is not on the list 
// it will restarted. A blank list means they'll all get restarted.

// list is formatted as a variant that contains a safearray of BSTR's - the progid's
void CMessageQueueListener1::ReStartComponents(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, VARIANT pVar)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{
		HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_RestartComponents(pVar);
		if (!SUCCEEDED(hRes))
		{
			CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::ReStartComponent ", L"Unable to restart queue ",(_com_error)NULL);
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::ReStartComponent ", L"Unable to find named queue ",(_com_error)NULL);
	}
}



void CMessageQueueListener1::AddStalledComponents(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut,  VARIANT pVar)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{
		HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_AddStalledComponents(pVar);
		if (!SUCCEEDED(hRes))
		{
			CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::AddStalledComponents ", L"Unable to restart queue ",(_com_error)NULL);
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::AddStalledComponents ", L"Unable to find named queue ",(_com_error)NULL);
	}
}



void CMessageQueueListener1::ChangeQueueThreads(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, int nThreads)
{

    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
	if (ThreadPoolManagerCommon1Ptr)
	{
		
		HRESULT hRes = ThreadPoolManagerCommon1Ptr->put_NumberOfThreads(nThreads);
		
		
		if (!SUCCEEDED(hRes))
		{
			CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOSTARTQUEUE, L"CMessageQueueListener1::ReStartQueue ", L"Unable to change queue's thread pool ",(_com_error)NULL);
		}
	}
	else
	{
		// haven't found the queue we're looking for
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOFINDNAMEDQUEUE, L"CMessageQueueListener1::ReStartQueue ", L"Unable to find named queue ",(_com_error)NULL);
	}
	
}


bstr_t  CMessageQueueListener1::RequestQueueInfo(
	BSTR bstrQueueName, QueueTypeEnum eQueueType, XmlNS::IXMLDOMDocument *ptrDomOut)
{
	// build up dom to contain info on the queue requested.

	// if root does not exist - create one otherwise add info to 

	bstr_t bstrElement1("QUEUELIST");
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	bstr_t XmlOut;
	
	enum
	{
		eParseDom,
		eUnpackArray,
		eNULL,
	} eAction = eNULL;

	try
	{
		eAction = eParseDom;
		ptrElementList = ptrDomOut->getElementsByTagName(bstrElement1);
		long lOutList = ptrElementList->Getlength();
		XmlNS::IXMLDOMElementPtr	ptrLevel1Element;
		XmlNS::IXMLDOMElementPtr	ptrReturnElement;
		XmlNS::IXMLDOMNodePtr		ptrNode;

		if (lOutList == 0)
		{
			// new error
			ptrReturnElement = ptrDomOut->createElement("RESPONSE");

			if (ptrReturnElement == NULL)
			{
				throw (_com_error(0,NULL));
			}
		
			ptrReturnElement->setAttribute("TYPE",L"SUCCESS");
			ptrNode = ptrDomOut->appendChild(ptrReturnElement);

			ptrLevel1Element = ptrDomOut->createElement("QUEUELIST");

			if (ptrLevel1Element == NULL)
			{
				throw (_com_error(0,NULL));
			}

			ptrNode = ptrReturnElement->appendChild(ptrLevel1Element);
		}
		else
		{
			ptrReturnElement = ptrDomOut->GetdocumentElement();

		}	

		XmlNS::IXMLDOMElementPtr	ptrLevel2Element;
		
		ptrLevel2Element = ptrDomOut->createElement("QUEUE");
		
		if (ptrLevel2Element == NULL)
		{
			throw (_com_error(0,NULL));
		}

		ptrNode = ptrLevel1Element->appendChild(ptrLevel2Element);

		XmlNS::IXMLDOMElementPtr	ptrLevel3Element;
		ptrLevel3Element = ptrDomOut->createElement("NAME");
		if (ptrLevel3Element == NULL)
		{
			throw (_com_error(0,NULL));
		}
		ptrLevel3Element->text = bstrQueueName;
		ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);

		// MMC snapin needs the queue type to be returned as well
		ptrLevel3Element = ptrDomOut->createElement("TYPE");
		if (ptrLevel3Element == NULL)
		{
			throw (_com_error(0,NULL));
		}
	
		if (eQueueType == eQueueTypeMSMQ1)
			ptrLevel3Element->text = "MSMQ1";
		else
			ptrLevel3Element->text = "OMMQ1";

		ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);

		ptrLevel3Element = ptrDomOut->createElement("STATUS");
		if (ptrLevel3Element == NULL)
		{
			throw (_com_error(0,NULL));
		}

	    IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
		ThreadPoolManagerCommon1Ptr = GetThreadPoolManagerCommon1Ptr(eQueueType, bstrQueueName);
		if (ThreadPoolManagerCommon1Ptr)
		{
			// check started / stalled etc
			BOOL bStarted;
			ThreadPoolManagerCommon1Ptr->get_Started(&bStarted);

			if (bStarted)
			{

				BOOL bStalled;
				ThreadPoolManagerCommon1Ptr->get_QueueStalled(&bStalled);
				if (bStalled)
				{
					ptrLevel3Element->text = L"STALLED";
				}
				else
				{
					ptrLevel3Element->text = L"RUNNING";
				}
		
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				// check for stalled components

				VARIANT pvarComponents;
				ThreadPoolManagerCommon1Ptr->get_ComponentsStalled(&pvarComponents);

				// unpack safearray
				LONG lLBound = 0;
				LONG lUBound = 0;
				BSTR* pbstr = NULL;
				HRESULT hr = S_OK;

				if (V_VT(&pvarComponents) != (VT_ARRAY | VT_BSTR) || // array of BSTRs
					SafeArrayGetDim(V_ARRAY(&pvarComponents)) != 1 || // one dimensional 
					FAILED(hr = SafeArrayGetLBound(V_ARRAY(&pvarComponents), 1, &lLBound)) ||
					FAILED(SafeArrayGetUBound(V_ARRAY(&pvarComponents), 1, &lUBound)) ||
					FAILED(SafeArrayAccessData(V_ARRAY(&pvarComponents), reinterpret_cast<void**>(&pbstr))))
				{
					// unable to unpack the array
					eAction = eUnpackArray;
					throw (_com_error(0,NULL));
				}
				else
				{

					LONG cElements = lUBound-lLBound+1;

					if (cElements > 0)
					{
				
						ptrLevel3Element = ptrDomOut->createElement("COMPONENTLIST");
						if (ptrLevel3Element == NULL)
						{
							throw (_com_error(0,NULL));
						}
						ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);

				
						for (int i = 0; i < cElements; i++)
						{
							if (pbstr[i] != NULL)
							{
								XmlNS::IXMLDOMElementPtr	ptrLevel4Element;
								ptrLevel4Element = ptrDomOut->createElement("COMPONENT");
								if (ptrLevel4Element == NULL)
								{
									throw (_com_error(0,NULL));
								}
								ptrLevel4Element->text = pbstr[i];
								ptrNode = ptrLevel3Element->appendChild(ptrLevel4Element);
							}
						}
					}
			
				}	
		
			}
			else
			{
				ptrLevel3Element->text = L"NOT STARTED";
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
			}

		}
		else
		{
			ptrLevel3Element->text = L"NOT CREATED";
			ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
		}


		// check the number of threads running


				
		ptrLevel3Element = ptrDomOut->createElement("THREADSLIST");
		if (ptrLevel3Element == NULL)
		{	
			throw (_com_error(0,NULL));
		}
		ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
		XmlNS::IXMLDOMElementPtr	ptrLevel4Element;
		ptrLevel4Element = ptrDomOut->createElement("THREADS");
		if (ptrLevel4Element == NULL)
		{
			throw (_com_error(0,NULL));
		}
		ptrNode = ptrLevel3Element->appendChild(ptrLevel4Element);


		XmlNS::IXMLDOMElementPtr	ptrLevel5Element;
		ptrLevel5Element = ptrDomOut->createElement("NUMBER");
		if (ptrLevel5Element == NULL)
		{
			throw (_com_error(0,NULL));
		}


		
		long nThreads;

		
		ThreadPoolManagerCommon1Ptr->get_NumberOfThreads(&nThreads);

		
		char buffer[20];
		ltoa(nThreads, buffer, 10);
   		bstr_t bstrNumber(buffer);

		
		ptrLevel5Element->text = bstrNumber;
		ptrNode = ptrLevel4Element->appendChild(ptrLevel5Element);


		// iterate thru the schedule manager getting out info on the events

		XmlNS::IXMLDOMDocumentPtr	ptrDomEvent(__uuidof(XmlNS::DOMDocument));	

		m_ScheduleManagerDaily.GetEventInfo(ThreadPoolManagerCommon1Ptr, ptrDomEvent);


		//add in subelements - not the root one.

		bstr_t bstrElement1("THREADSLIST");
		XmlNS::IXMLDOMNodeListPtr ptrElementList;
		ptrElementList = ptrDomEvent->getElementsByTagName(bstrElement1);
		long lThreadList = ptrElementList->Getlength();
	
		// only ever 1 threadlist but sometimes we have no events

		if (lThreadList != 0)
		{

			ptrNode=ptrElementList->Getitem(0);

			// if node is ok then grab <QUEUE> 
				
			XmlNS::IXMLDOMNodeListPtr	ptrQueueElementList;
			ptrNode->get_childNodes(&ptrQueueElementList);
	
			long lThreads = ptrQueueElementList->Getlength();

		
			ptrNode = ptrLevel3Element->appendChild(ptrQueueElementList->Getitem(0));

		
			// if we have some loop around and get info out.
			for(int iThreadLoop = 1; iThreadLoop<lThreads; iThreadLoop++)
			{
				// unpack child
	
				ptrNode = ptrLevel3Element->appendChild(ptrQueueElementList->nextNode());
	
			}
		}

		// PSC 11/07/2006 CORE284 - Start
		ptrLevel3Element = ptrDomOut->createElement("TABLESUFFIX");
		if (ptrLevel3Element == NULL)
		{
			throw (_com_error(0,NULL));
		}

		BSTR bstrTableSuffix;
		ThreadPoolManagerCommon1Ptr->get_TableSuffix(&bstrTableSuffix);
		ptrLevel3Element->text = bstrTableSuffix;
		ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);

		::SysFreeString(bstrTableSuffix);
		// PSC 11/07/2006 CORE284 - End

		// get the root element
		ptrReturnElement = ptrDomOut->GetdocumentElement();

#ifdef DEBUG
		_variant_t varPath = _T("c:\\xml.xml");

		ptrDomOut->save(varPath);
#endif
		XmlOut=ptrReturnElement->xml;
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{

			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, L"CMessageQueueListener1::RequestQueueInfo ", L"Unable to process tag ",comerr);
				break;
			
			case eUnpackArray:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_FAILEDTOUNPACKSAFEARRAY, L"CMessageQueueListener1::RequestQueueInfo ", L"Unable to unpack safearray ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::RequestQueueInfo ", L"Undefined error ",comerr);
				break;
		}

	
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, L"CMessageQueueListener1::RequestQueueInfo ", L"Undefined error ",(_com_error)NULL);
	}
	
	
	return XmlOut;
}



bstr_t  CMessageQueueListener1::RequestAllQueueInfo(XmlNS::IXMLDOMDocument *ptrDomOut)
{
	// build up dom to contain info on all of the queues.

	// if root does not exist - create one otherwise add info to 


	bstr_t bstrElement1("QUEUELIST");
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	bstr_t XmlOut;
	
	enum
	{
		eParseDom,
		eUnpackArray,
		eNULL,
	} eAction = eNULL;

	try
	{
	
		eAction = eParseDom;
		ptrElementList = ptrDomOut->getElementsByTagName(bstrElement1);
		long lOutList = ptrElementList->Getlength();
		XmlNS::IXMLDOMElementPtr	ptrLevel1Element;
		XmlNS::IXMLDOMElementPtr	ptrReturnElement;
		XmlNS::IXMLDOMNodePtr		ptrNode;


		if (lOutList == 0)
		{
			// new error
			ptrReturnElement = ptrDomOut->createElement("RESPONSE");

			if (ptrReturnElement == NULL)
			{
				throw (_com_error(0,NULL));
			}
			
			
			ptrReturnElement->setAttribute("TYPE",L"SUCCESS");
			ptrNode = ptrDomOut->appendChild(ptrReturnElement);

			ptrLevel1Element = ptrDomOut->createElement("QUEUELIST");

			if (ptrLevel1Element == NULL)
			{
				throw (_com_error(0,NULL));
			}

			
			ptrNode = ptrReturnElement->appendChild(ptrLevel1Element);

		}
		else
		{
			ptrReturnElement = ptrDomOut->GetdocumentElement();

		}	

		
		// loop around and get all the queues

		bool bLeaveLoop = false;
		IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;

		MapQueueNameToCommon1::iterator theCommonIterator = m_MapQueueNameToCommon1.begin();

		
		while (bLeaveLoop == false)
		{

		
			if (theCommonIterator != m_MapQueueNameToCommon1.end() )
			{
				ThreadPoolManagerCommon1Ptr = (*theCommonIterator).second;

		
				XmlNS::IXMLDOMElementPtr	ptrLevel2Element;
				
				ptrLevel2Element = ptrDomOut->createElement("QUEUE");
		
				if (ptrLevel2Element == NULL)
				{
					throw (_com_error(0,NULL));
				}

				ptrNode = ptrLevel1Element->appendChild(ptrLevel2Element);

				XmlNS::IXMLDOMElementPtr	ptrLevel3Element;
				ptrLevel3Element = ptrDomOut->createElement("NAME");
				if (ptrLevel3Element == NULL)
				{
					throw (_com_error(0,NULL));
				}
			

				BSTR bstrQueueName;

				ThreadPoolManagerCommon1Ptr->get_QueueName(&bstrQueueName);
				// get name from manager object
				ptrLevel3Element->text = bstrQueueName;
			
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);

				// PSC 11/07/2006 CORE284 - Start
				ptrLevel3Element = ptrDomOut->createElement("TYPE");
				if (ptrLevel3Element == NULL)
				{
					throw (_com_error(0,NULL));
				}

				BSTR bstrQueueType;
				ThreadPoolManagerCommon1Ptr->get_QueueType(&bstrQueueType);
				ptrLevel3Element->text = bstrQueueType;
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				// PSC 11/07/2006 CORE284 - End

				ptrLevel3Element = ptrDomOut->createElement("STATUS");
				if (ptrLevel3Element == NULL)
				{
					throw (_com_error(0,NULL));
				}

				// check started / stalled etc
				BOOL bStarted;
				ThreadPoolManagerCommon1Ptr->get_Started(&bStarted);

				if (bStarted)
				{

					BOOL bStalled;
					ThreadPoolManagerCommon1Ptr->get_QueueStalled(&bStalled);
					if (bStalled)
					{
						ptrLevel3Element->text = L"STALLED";
					}
					else
					{
						ptrLevel3Element->text = L"RUNNING";
					}	
		
					ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
					// check for stalled components
	
					VARIANT pvarComponents;
					ThreadPoolManagerCommon1Ptr->get_ComponentsStalled(&pvarComponents);

					// unpack safearray
					LONG lLBound = 0;
					LONG lUBound = 0;
					BSTR* pbstr = NULL;
					HRESULT hr = S_OK;

					if (V_VT(&pvarComponents) != (VT_ARRAY | VT_BSTR) || // array of BSTRs
						SafeArrayGetDim(V_ARRAY(&pvarComponents)) != 1 || // one dimensional 
						FAILED(hr = SafeArrayGetLBound(V_ARRAY(&pvarComponents), 1, &lLBound)) ||
						FAILED(SafeArrayGetUBound(V_ARRAY(&pvarComponents), 1, &lUBound)) ||
						FAILED(SafeArrayAccessData(V_ARRAY(&pvarComponents), reinterpret_cast<void**>(&pbstr))))
					{
						// unable to unpack the array
						eAction = eUnpackArray;
						throw (_com_error(0,NULL));
					}
					else
					{

						LONG cElements = lUBound-lLBound+1;
						if (cElements > 0)
						{
			
							ptrLevel3Element = ptrDomOut->createElement("COMPONENTLIST");
							if (ptrLevel3Element == NULL)
							{
								throw (_com_error(0,NULL));
							}
							ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				
							for (int i = 0; i < cElements; i++)
							{
								if (pbstr[i] != NULL)
								{
									XmlNS::IXMLDOMElementPtr	ptrLevel4Element;
									ptrLevel4Element = ptrDomOut->createElement("COMPONENT");
									if (ptrLevel4Element == NULL)
									{
										throw (_com_error(0,NULL));
									}
									ptrLevel4Element->text = pbstr[i];
									ptrNode = ptrLevel3Element->appendChild(ptrLevel4Element);
								}
							}
						}
			
					}	
		
				}
				else
				{
					ptrLevel3Element->text = L"NOT STARTED";
					ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				}


				// check the number of threads running
	
				ptrLevel3Element = ptrDomOut->createElement("THREADSLIST");
				if (ptrLevel3Element == NULL)
				{	
					throw (_com_error(0,NULL));
				}
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				XmlNS::IXMLDOMElementPtr	ptrLevel4Element;
				ptrLevel4Element = ptrDomOut->createElement("THREADS");
				if (ptrLevel4Element == NULL)
				{
					throw (_com_error(0,NULL));
				}
				ptrNode = ptrLevel3Element->appendChild(ptrLevel4Element);

				XmlNS::IXMLDOMElementPtr	ptrLevel5Element;
				ptrLevel5Element = ptrDomOut->createElement("NUMBER");
				if (ptrLevel5Element == NULL)
				{
					throw (_com_error(0,NULL));
				}


		
				long nThreads;

		
				ThreadPoolManagerCommon1Ptr->get_NumberOfThreads(&nThreads);

		
				char buffer[20];
				ltoa(nThreads, buffer, 10);
			   	bstr_t bstrNumber(buffer);

		
				ptrLevel5Element->text = bstrNumber;
				ptrNode = ptrLevel4Element->appendChild(ptrLevel5Element);


				// iterate thru the schedule manager getting out info on the events
	
				XmlNS::IXMLDOMDocumentPtr	ptrDomEvent(__uuidof(XmlNS::DOMDocument));	

				m_ScheduleManagerDaily.GetEventInfo(ThreadPoolManagerCommon1Ptr, ptrDomEvent);


				//add in subelements - not the root one.

				bstr_t bstrElement1("THREADSLIST");
				XmlNS::IXMLDOMNodeListPtr ptrElementList;
				ptrElementList = ptrDomEvent->getElementsByTagName(bstrElement1);
				long lThreadList = ptrElementList->Getlength();
	
				// only ever 1 threadlist but sometimes we have no events
				if (lThreadList != 0)
				{

					ptrNode=ptrElementList->Getitem(0);

					// if node is ok then grab <QUEUE> 
				
					XmlNS::IXMLDOMNodeListPtr	ptrQueueElementList;
					ptrNode->get_childNodes(&ptrQueueElementList);
	
					long lThreads = ptrQueueElementList->Getlength();

		
					ptrNode = ptrLevel3Element->appendChild(ptrQueueElementList->Getitem(0));

		
					// if we have some loop around and get info out.
					for(int iThreadLoop = 1; iThreadLoop<lThreads; iThreadLoop++)
					{
						// unpack child
	
						ptrNode = ptrLevel3Element->appendChild(ptrQueueElementList->nextNode());
	
					}
				}

				// PSC 11/07/2006 CORE284 - Start
				ptrLevel3Element = ptrDomOut->createElement("TABLESUFFIX");
				if (ptrLevel3Element == NULL)
				{
					throw (_com_error(0,NULL));
				}

				BSTR bstrTableSuffix;
				ThreadPoolManagerCommon1Ptr->get_TableSuffix(&bstrTableSuffix);
				ptrLevel3Element->text = bstrTableSuffix;
				ptrNode = ptrLevel2Element->appendChild(ptrLevel3Element);
				// PSC 11/07/2006 CORE284 - End

			
				::SysFreeString(bstrQueueName);
				
				// PSC 11/07/2006 CORE284 - Start
				::SysFreeString(bstrQueueType);
				::SysFreeString(bstrTableSuffix);
				// PSC 11/07/2006 CORE284 - End

				theCommonIterator++;

			}
			else
			{
				// else stop the loop as we have no queues
				bLeaveLoop = true;
			}
				

		}
		// get the root element
		ptrReturnElement = ptrDomOut->GetdocumentElement();

#ifdef DEBUG
		_variant_t varPath = _T("c:\\xml.xml");

		ptrDomOut->save(varPath);
#endif
		XmlOut=ptrReturnElement->xml;
	}
	catch(_com_error comerr)
	{
		switch (eAction)
		{

			case eParseDom:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNABLETOPROCESSTAG, 
					L"CMessageQueueListener1::RequestAllQueueInfo", 
					L"Unable to process tag ",comerr);
				break;
			
			case eUnpackArray:
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_FAILEDTOUNPACKSAFEARRAY, 
					L"CMessageQueueListener1::RequestAllQueueInfo", 
					L"Unable to unpack safearray ",comerr);
				break;
			default :
				CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
					L"CMessageQueueListener1::RequestAllQueueInfo", 
					L"Undefined error ",comerr);
				break;
		}
	
	}
	catch(...)
	{
		CreateErrorInfo(ptrDomOut, L"APPERR", ERROR_UNDEFINEDERROR, 
			L"CMessageQueueListener1::RequestAllQueueInfo", 
			L"Undefined error ",(_com_error)NULL);
	}
	
	
	return XmlOut;
}




void CMessageQueueListener1::RemoveEvent(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, GUID gKey)
{

	bstr_t XmlOut;


	// call the schedule manager function

	
	// remove the scheduled events associated with the queue
	m_ScheduleManagerDaily.RemoveEvent(ptrDomOut,gKey);


//	return XmlOut;
}


IInternalThreadPoolManagerCommon1Ptr CMessageQueueListener1::GetThreadPoolManagerCommon1Ptr(
	QueueTypeEnum eQueueType, BSTR bstrQueueName)
{
	IInternalThreadPoolManagerCommon1Ptr ThreadPoolManagerCommon1Ptr;
	switch (eQueueType)
	{
		case eQueueTypeMSMQ1:
		case eQueueTypeOMMQ1:
		{
			MapQueueNameToCommon1::iterator theCommonIterator = 
				m_MapQueueNameToCommon1.find(bstrQueueName);
			if (theCommonIterator != m_MapQueueNameToCommon1.end() )
			{
				ThreadPoolManagerCommon1Ptr = (*theCommonIterator).second;
			}
			break;
		}
		default:
			_ASSERTE(0); // should not reach here 
			break;
	}
	return ThreadPoolManagerCommon1Ptr;
}

CMessageQueueListener1::QueueTypeEnum CMessageQueueListener1::GeteQueueType(BSTR bstrQueueType)
{
	QueueTypeEnum eQueueType = eQueueTypeUnknown;

	if (wcsicmp(bstrQueueType, L"MSMQ1") == 0)
	{
		eQueueType = eQueueTypeMSMQ1;
	}
	else 
	{
		eQueueType = eQueueTypeOMMQ1;
	}
	return eQueueType;
}

void CMessageQueueListener1::CreateErrorInfo(XmlNS::IXMLDOMDocument *ptrDomOut, BSTR strType, BSTR strNumber, BSTR strSource, BSTR strDescription,_com_error& comerr)
{
	XmlNS::IXMLDOMElementPtr	ptrReturnElement;
	XmlNS::IXMLDOMNodePtr		ptrNode;


	// create elements in the DOM

	//<RESPONSE TYPE="SYSERR|APPERR">
	//	<ERROR>
	//		<NUMBER></NUMBER>
	//		<SOURCE></SOURCE>
	//		<DESCRIPTION ></DESCRIPTION>
	//		<VERSION></VERSION>
	//	</ERROR>
	//</ RESPONSE>


	// check if dom already has root element - if it has attach new children
	// otherwise create it.


	bstr_t bstrElement3("ERROR");
	XmlNS::IXMLDOMNodeListPtr	ptrElementList;
	ptrElementList = ptrDomOut->getElementsByTagName(bstrElement3);
	long lOutList = ptrElementList->Getlength();


	if (lOutList == 0)
	{
		// new error
		ptrReturnElement = ptrDomOut->createElement("RESPONSE");
		ptrReturnElement->setAttribute("TYPE",strType);
		ptrNode = ptrDomOut->appendChild(ptrReturnElement);
	}
	else
	{
		ptrReturnElement = ptrDomOut->GetdocumentElement();

	}

	if (wcsicmp (strType, L"SUCCESS") != 0)
	{
		// it's an error
		
		XmlNS::IXMLDOMElementPtr	ptrSubElement;
		XmlNS::IXMLDOMElementPtr	ptrSubSubElement;

		ptrSubElement = ptrDomOut->createElement("ERROR");
	
		ptrNode = ptrReturnElement->appendChild(ptrSubElement);

		ptrSubSubElement = ptrDomOut->createElement("NUMBER");
		ptrSubSubElement->text = strNumber;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

		ptrSubSubElement = ptrDomOut->createElement("SOURCE");
		ptrSubSubElement->text = strSource;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);

		ptrSubSubElement = ptrDomOut->createElement("DESCRIPTION");
		ptrSubSubElement->text = strDescription;
		ptrNode = ptrSubElement->appendChild(ptrSubSubElement);
		
		// and also write the error to the event log
		_bstr_t bstrErrorMessage(strDescription);
		bstrErrorMessage += _T(" ");
		bstrErrorMessage += strSource;
		LogEventError(comerr, bstrErrorMessage);
	}
}


void CMessageQueueListener1::LogEventError(_com_error& comerr, LPCTSTR pszFormat, ...)
{
    try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		if (comerr.Description().length() == 0)
		{
			// output error message instead of description
			_Module.LogEventError(_T("%s (HResult = %d')"), chMsg, comerr.ErrorMessage());
		}
		else
		{
			// output description
			_Module.LogEventError(_T("%s (HResult = %d')"), chMsg, comerr.Description());
		}
	}
	catch(...)
	{
	}
}

int CMessageQueueListener1::GetDigit(LPSTR lpstr)
{ 
	char ch = *lpstr;
    if (ch >= '0' && ch <= '9')
		return(ch - '0');
	
	if (ch >= 'a' && ch <= 'f')
		return(ch - 'a' + 10);
	
	if (ch >= 'A' && ch <= 'F')
		return(ch - 'A' + 10);
	
	return(0);
} 


void CMessageQueueListener1::ConvertField(LPBYTE lpByte,LPSTR * ppStr,int iFieldSize,BOOL fRightToLeft)
{ 
	int i;
	for (i=0;i<iFieldSize ;i++ )
	{ // don't barf on the field separators 
		if ('-' == **ppStr)
			(*ppStr)++;
		if (fRightToLeft == TRUE)
		{ 
			// work from right to left within the byte stream 
			*(lpByte + iFieldSize - (i+1)) = 16*GetDigit(*ppStr) + GetDigit((*ppStr)+1);
		}
		else
		{ 
			// work from  left to right within the byte stream 
			*(lpByte + i) = 16*GetDigit(*ppStr) + GetDigit((*ppStr)+1);
		} 
		*ppStr+=2; 
		// get next two digit pair 
	} 
}

#define GUID_STRING_SIZE 64  

HRESULT CMessageQueueListener1::GUIDFromString(LPWSTR lpWStr, GUID * pGuid)
{ 
	BYTE * lpByte; // byte index into guid 
	int iFieldSize; // size of current field we're converting 
					// since its a guid, we can do a "brute force" conversion 
	char lpTemp[GUID_STRING_SIZE];
	char *lpStr = lpTemp;
	WideToAnsi(lpStr,lpWStr,GUID_STRING_SIZE);
	
	// make sure we have a {xxxx-...} type guid 
	if ('{' !=  *lpStr) 
		return E_FAIL; 
	lpStr++;  
	
	lpByte = (BYTE *)pGuid; 
	
	// data 1 
	iFieldSize = sizeof(unsigned long);
	ConvertField(lpByte,&lpStr,iFieldSize,TRUE);
	lpByte += iFieldSize;  
	
	// data 2 
	iFieldSize = sizeof(unsigned short); 
	ConvertField(lpByte,&lpStr,iFieldSize,TRUE);
	lpByte += iFieldSize;  
	
	// data 3 
	iFieldSize = sizeof(unsigned short); 
	ConvertField(lpByte,&lpStr,iFieldSize,TRUE);
	lpByte += iFieldSize;  
	
	// data 4 
	iFieldSize = 8*sizeof(unsigned char);
	ConvertField(lpByte,&lpStr,iFieldSize,FALSE);
	lpByte += iFieldSize;  
	
	// make sure we ended in the right place 
	if ('}' != *lpStr)  
	{ 
		memset(pGuid,0,sizeof(GUID));
		return E_FAIL;
	}  
	
	return S_OK;
}

int CMessageQueueListener1::WideToAnsi(LPSTR lpStr,LPWSTR lpWStr,int cchStr)
{
	int rval; 
	BOOL bDefault;  
	
	// use the default code page (CP_ACP) 
	// -1 indicates WStr must be null terminated 
	
	rval = WideCharToMultiByte(CP_ACP,0,lpWStr,-1,lpStr,cchStr,"-",&bDefault);
	return rval;
} 

bstr_t CMessageQueueListener1::LookupConnectionString( BSTR bstrDefaultConnStr )
{
	// Check the registry for the connection string data. The default is passed in, but 
	// only to provide backwards compatibility. The configuration data includes the path
	// to a repository file that is used to store the connection string in an encrypted 
	// format.

	// The path to the repository is stored in the registry at;
	// HKEY_LOCAL_MACHINE/SOFTWARE/OMIGA/Repository/Path
	// and the path to the rest of the configuration data is stored under the key;
	// HKEY_LOCAL_MACHINE/SOFTWARE/OMIGA/MessageQueueListener/Database Connection

	enum
	{
		edbpOracle,
		edbpSQLServer,
		edbpUnknown
	} edbpProvider = edbpUnknown;

    CRegKey keySoftware;
	bstr_t bstrRetVal(bstrDefaultConnStr);
	bstr_t bstrRepositoryPath = DEFAULT_REPOSITORY_PATH;
    LONG lRes;
	CComBSTR cbstrDBSchema;
	CComBSTR cbstrDBInstance;
	CComBSTR cbstrPassword;
	
	try
	{
		// Open the registry path to the configuration data
		lRes = keySoftware.Open(HKEY_LOCAL_MACHINE, _T("SOFTWARE"), KEY_READ);
		if (lRes == ERROR_SUCCESS)
		{
			CRegKey keyOmiga;
			lRes = keyOmiga.Open(keySoftware, _T("OMIGA"), KEY_READ);
			if (lRes == ERROR_SUCCESS)
			{
				CRegKey keyComponent;
				const long lValueLen = 500;
				TCHAR szValue[lValueLen];
				DWORD dwLen = lValueLen;

				// Check for the repository details first
				lRes = keyComponent.Open(keyOmiga, _T("Repository"), KEY_READ);
				if (lRes == ERROR_SUCCESS)
				{
					// Get the path to the repository
					lRes = keyComponent.QueryValue(szValue, _T("Path"), &dwLen);
					if (lRes == ERROR_SUCCESS) bstrRepositoryPath = szValue;

					keyComponent.Close();
				}

				// Now check for the connection string configuration data
				lRes = keyComponent.Open(keyOmiga, _T("MessageQueueListener"), KEY_READ);
				if (lRes == ERROR_SUCCESS)
				{
					CRegKey keyDB;
					lRes = keyDB.Open(keyComponent, _T("Database Connection"), KEY_READ);
					if (lRes == ERROR_SUCCESS)
					{
						// Reset the return value
						bstrRetVal = "";

						// First get the provider details
						lRes = keyDB.QueryValue(szValue, _T("Provider"), &dwLen);
						if (lRes == ERROR_SUCCESS)
						{
							if ( _tcsicmp(szValue,L"MSDAORA") == 0 || _tcsicmp(szValue,L"ORAOLEDB.ORACLE") == 0 )
							{
								edbpProvider = edbpOracle;
							}
							else if ( _tcsicmp(szValue,L"SQLOLEDB") == 0 )
							{
								edbpProvider = edbpSQLServer;
							}

							// Add this to the connection string
							bstrRetVal += _bstr_t("Provider=") + szValue + ";";
						}

						// Next get the user id for the access
						dwLen = lValueLen;
						lRes = keyDB.QueryValue(szValue, _T("User ID"), &dwLen);
						if (lRes == ERROR_SUCCESS)
						{
							cbstrDBSchema = szValue;
							if ( edbpProvider == edbpOracle )
								bstrRetVal += "User ID=" + _bstr_t(szValue) + ";";
							else
								bstrRetVal += "UID=" + _bstr_t(szValue) + ";";
						}

						// See if there is a default password in the registry
						dwLen = lValueLen;
						lRes = keyDB.QueryValue(szValue, _T("Password"), &dwLen);
						if (lRes == ERROR_SUCCESS)
						{
							cbstrPassword = szValue;
							if ( edbpProvider == edbpOracle )
								bstrRetVal += _bstr_t("Password=") + szValue + ";";
							else
								bstrRetVal += _bstr_t("PWD=") + szValue + ";";

						}

						// Now do the provider specific processing
						if ( edbpProvider == edbpOracle )
						{
							// Next find the source schema 
							dwLen = lValueLen;
							lRes = keyDB.QueryValue(szValue, _T("Data Source"), &dwLen);
							if (lRes == ERROR_SUCCESS)
							{
								cbstrDBInstance = szValue;
								bstrRetVal += "Data Source=" + _bstr_t(szValue) + ";";
							}

							// See if there are any additional options to the connection string
							dwLen = lValueLen;
							lRes = keyDB.QueryValue(szValue, _T("Additional_Options"), &dwLen);
							if (lRes == ERROR_SUCCESS)
							{
								bstrRetVal += _bstr_t(szValue) + ";";
							}
						}
						else
						{
							// Get the server name
							dwLen = lValueLen;
							lRes = keyDB.QueryValue(szValue, _T("Server"), &dwLen);
							if (lRes == ERROR_SUCCESS)
							{
								bstrRetVal += "Server=" + _bstr_t(szValue) + ";";
							}

							// Get the database name
							dwLen = lValueLen;
							lRes = keyDB.QueryValue(szValue, _T("Database Name"), &dwLen);
							if (lRes == ERROR_SUCCESS)
							{
								cbstrDBInstance = szValue;
								bstrRetVal += "Database=" + _bstr_t(szValue) + ";";
							}
						}

						// Close the registry key
						keyDB.Close();
					}

					keyComponent.Close();
				}

				keyOmiga.Close();
			}

			keySoftware.Close();
		}

		// If any values are missing then extract them from the connection string passed in
		bstr_t bstrDefConnStr = _tcsupr(bstrDefaultConnStr);

		if ( edbpProvider == edbpUnknown )
		{
			if (_tcsstr((LPCWSTR)bstrDefConnStr, L"SQLOLEDB") != NULL) edbpProvider = edbpSQLServer;
			else edbpProvider = edbpOracle;
		}

		if ( cbstrDBInstance.Length() == 0 )
		{
			cbstrDBInstance = (LPCSTR) GetConnStrSetting( bstrDefConnStr, (edbpProvider == edbpOracle) ? "DATA SOURCE=" : "DATABASE=" );
		}

		if ( cbstrDBSchema.Length == 0 )
		{
			cbstrDBSchema = (LPCSTR) GetConnStrSetting( bstrDefConnStr, (edbpProvider == edbpOracle) ? "USER ID=" : "UID=" );
		}

		// Now try and get a password from the repository
		bstr_t bstrFoundPassword = GetPasswordFromRepository( bstrRepositoryPath, cbstrDBInstance, cbstrDBSchema );

		// If a password has been returned
		if ( bstrFoundPassword.length() > 0 ) 
		{
			// Search for the password placeholder in the connection string
			string strVal = bstr_t( _tcsupr(bstrRetVal));
			string strPWDText = (edbpProvider == edbpOracle) ? "PASSWORD=" : "PWD=";
			int intStartIndex = strVal.find( strPWDText );

			if ( intStartIndex != -1 )
			{
				// Replace the password in the connection string
				intStartIndex += strPWDText.length();
				int intEndIndex = strVal.find( ";", intStartIndex );
				strVal.replace( intStartIndex, intEndIndex - intStartIndex, bstrFoundPassword );
				bstrRetVal = strVal.c_str();
			}
			else
			{
				// Append the password to the connection string
				bstrRetVal += strPWDText.c_str();
				bstrRetVal += bstrFoundPassword + ";";
			}
		}
	}
	catch(...)
	{
		// clean up and rethrow the error
		throw;
	}
	return bstrRetVal;
}

bstr_t CMessageQueueListener1::GetPasswordFromRepository( BSTR bstrRepositoryPath, 
														  BSTR bstrInstance, 
														  BSTR bstrSchema )
{
	bstr_t bstrResponse;
	DISPPARAMS dispparams;
	DISPID dispid;
	memset(&dispparams, 0, sizeof dispparams);
	BSTR bstrOpenMethod = SysAllocString(L"OpenRepository");
	BSTR bstrGetPwdMethod = SysAllocString(L"GetPassword");
	HRESULT hr = S_OK;
	CComPtr<IDispatch> pIDispRepository;
	CComBSTR bstrDecryptionKey = "{49DCCF38-84FC-4158-B099-33523A0E49B2}";
	CComBSTR bstrPasswordDecryptionKey = "{31FBDB54-494E-4738-AF2F-BEE63119BA08}";
	_variant_t varResult;

	try
	{
		// create the component
		hr = pIDispRepository.CoCreateInstance(_T("RepositoryReader.Reader1"));
		if ( SUCCEEDED( hr ) )
		{
			// get the dispid for the open function
			hr = pIDispRepository->GetIDsOfNames( IID_NULL, &bstrOpenMethod, 1, 
												  LOCALE_USER_DEFAULT, &dispid );
		}

		if ( SUCCEEDED( hr ) )
		{
			// Allocate memory for all VARIANTARG parameters.
			dispparams.cArgs = 2;
			dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
			memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);

			// Set up the arguments
			dispparams.rgvarg[0].vt = VT_BYREF|VT_BSTR;
			dispparams.rgvarg[0].pbstrVal = &bstrDecryptionKey.m_str;
			dispparams.rgvarg[1].vt = VT_BSTR;
			dispparams.rgvarg[1].bstrVal = bstrRepositoryPath;

			// Make the call.
			hr = pIDispRepository->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, 
						  				  DISPATCH_METHOD, &dispparams, &varResult, NULL, NULL);
		}

		if ( SUCCEEDED ( hr ) && varResult.boolVal == VARIANT_TRUE )
		{
			// Look up the method call to get the password
			hr = pIDispRepository->GetIDsOfNames( IID_NULL, &bstrGetPwdMethod, 1, 
												  LOCALE_USER_DEFAULT, &dispid );
		}

		// Was the repository opened successfully
		if ( SUCCEEDED ( hr ) && varResult.boolVal == VARIANT_TRUE )
		{
			// Allocate memory for all VARIANTARG parameters.
			dispparams.cArgs = 3;
			dispparams.rgvarg = new VARIANTARG[dispparams.cArgs];
			memset(dispparams.rgvarg, 0, sizeof(VARIANTARG) * dispparams.cArgs);

			// prepare the arguments
			dispparams.rgvarg[0].vt = VT_BSTR;
			dispparams.rgvarg[0].bstrVal = bstrPasswordDecryptionKey;
			dispparams.rgvarg[1].vt = VT_BSTR;
			CComBSTR cbstrSchema(bstrSchema);
			cbstrSchema.ToUpper();
			dispparams.rgvarg[1].bstrVal = cbstrSchema.m_str;
			dispparams.rgvarg[2].vt = VT_BSTR;
			CComBSTR cbstrInstance(bstrInstance);
			cbstrInstance.ToUpper();
			dispparams.rgvarg[2].bstrVal = cbstrInstance.m_str;
			VariantClear( &varResult );
    
			// Make the call.
			hr = pIDispRepository->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD,
				&dispparams, &varResult, NULL, NULL);
		}

		if ( SUCCEEDED ( hr ) )
		{
			// Set up the password in the connection string
			bstrResponse = varResult.bstrVal;
		}
	}
	catch(...)
	{
		// clean up and rethrow the error
		SysFreeString(bstrOpenMethod);
		SysFreeString(bstrGetPwdMethod);
		if (dispparams.rgvarg) delete[] dispparams.rgvarg;
		throw;
	}

	// Free any allocated resources
	SysFreeString(bstrOpenMethod);
	SysFreeString(bstrGetPwdMethod);
	if (dispparams.rgvarg) delete[] dispparams.rgvarg;

	return bstrResponse;
}

bstr_t CMessageQueueListener1::GetConnStrSetting( bstr_t bstrConnString, bstr_t bstrSetting )
{
	string strVal = bstrConnString;
	bstr_t bstrRetval;

	// Find the correct placeholder for this setting
	int intStartIndex = strVal.find( bstrSetting );
	if ( intStartIndex != -1 )
	{
		// Increment the pointer to the end of the placeholder
		intStartIndex += bstrSetting.length();

		// Find the terminator
		int intEndIndex = strVal.find( ";", intStartIndex );

		// Return the current value
		bstrRetval = strVal.substr( intStartIndex, intEndIndex - intStartIndex ).c_str();
	}
	return bstrRetval;
}
