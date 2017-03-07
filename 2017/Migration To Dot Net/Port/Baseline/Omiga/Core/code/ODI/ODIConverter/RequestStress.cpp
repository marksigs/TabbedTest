///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestStress.cpp
//	DESCRIPTION:	
//		Class for stress testing the CODIConverter1::Request function, which
//		handles calls to the ODIConverter.ODIConverter interface Request method.
//
//		This calls spawns the requested number of threads. Each thread then
//		loops, repeatedly calling CODIConverter1::Request. This is useful for
//		basic stress testing, without having to use dedicated load testing
//		software.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		22/10/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "RequestStress.h"


BOOL CRequestStress::s_bStressing = FALSE;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestStress::CRequestStress(LPCWSTR szType) :
	CRequest(szType),
	m_bstrRequest(L""),
	m_nMaxThreads(0),
	m_nSleep(0)
{
}

CRequestStress::~CRequestStress()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestStress::Execute
//	
//	Description:
//		Virtual function implementation.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode
//			XML that must contain the attributes:
//				REQUEST: The request to be stress tested.
//				MAXTHREADS: The number of threads to be spawned.
//				SLEEP: How long (in milliseconds) each thread should wait 
//				between iterations.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Empty response XML.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestStress::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	if (!s_bStressing)
	{
		::InterlockedExchange(reinterpret_cast<LPLONG>(&s_bStressing), TRUE);

		m_bstrRequest		= static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"REQUEST");
		m_nMaxThreads		= _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"MAXTHREADS"));
		m_nSleep			= _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"SLEEP"));

		for (int nThread = 0; nThread < m_nMaxThreads; nThread++)
		{
			DWORD dwThreadID = 0;
			::CreateThread(NULL, 0, RequestThread, this, 0, &dwThreadID);
			::Sleep(m_nSleep);
		}
	}
	else
	{
		::InterlockedExchange(reinterpret_cast<LPLONG>(&s_bStressing), FALSE);
	}

	MSXML::IXMLDOMDocumentPtr ptrResponseDoc(__uuidof(MSXML::DOMDocument));
	MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrResponseDoc->createElement(L"RESPONSE");
	ptrResponseDoc->appendChild(ptrXMLDOMElement);
	ptrXMLDOMElement->setAttribute(L"TYPE", L"SUCCESS");

	return ptrXMLDOMElement;
}

DWORD WINAPI CRequestStress::RequestThread(LPVOID lpParameter)
{
	BOOL bSuccess = TRUE;

#ifdef _DEBUG
    _ASSERTE(SUCCEEDED(::CoInitialize(NULL)));
#else
	::CoInitialize(NULL);
#endif

	try
	{
		CRequestStress* pRequestStress = reinterpret_cast<CRequestStress*>(lpParameter);

		_ASSERTE(pRequestStress != NULL);

		if (pRequestStress != NULL)
		{
			while (s_bStressing)
			{
				CODIConverter1::Request(pRequestStress->m_bstrRequest);
				::Sleep(pRequestStress->m_nSleep);
			}
		}
	}
	catch(_com_error& e)
	{
		CException Exception(E_GENERICERROR, e, __FILE__, __LINE__);
		CODIConverter1::LogError(Exception);
	}
	catch(CException& e)
	{
		CODIConverter1::LogError(e);
	}
	/*
	// AS 22/01/2007 VS2005 Port.
	// error C2316: 'exception &' : cannot be caught as the destructor and/or copy constructor are inaccessible.
	catch(exception& e)
	{
		CException Exception(E_GENERICERROR, e, __FILE__, __LINE__);
		CODIConverter1::LogError(Exception);
	}
	*/
	catch(...)
	{
		CException Exception(E_GENERICERROR, __FILE__, __LINE__, _T("Unknown error"));
		CODIConverter1::LogError(Exception);
	}

    ::CoUninitialize();

	return bSuccess;
}
