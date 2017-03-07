///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestThreadWait.cpp
//	DESCRIPTION:	
//
//		Class for testing ODIConverter concurrency, i.e., to check that
//		ODIConverter is spawning multiple threads to handle incoming 
//		calls to the ODIConverter.ODIConverter interface, Request method, and
//		that calls to the method are not being serialized.
//
//		This can be invoked using the following ASP VBScript:
//
//		<%
//		Dim objODIConverter 
//		Dim objResponse
//		Dim strRequest
//		Dim strResponse
//  
//		Set objODIConverter = Server.CreateObject("ODIConverter.ODIConverter")
//		Set objResponse = Server.CreateObject("MICROSOFT.XMLDOM")
//  
//		Response.Write "<TABLE border =1><TR><TD colspan=2 align=center><H3>MSDN Q216580<BR>ThreadWait 5 seconds</H3></TD></TR>"
//		Response.Write "<TR><TD>StartTime: </TD><TD>" & Now & "</TD></TR>"
//  
//		Response.Write "<TR><TD>ThreadID: </TD><TD>" 
//		strRequest = "<REQUEST OPERATION=""THREADWAIT"" WAIT=""5""/>"
//		strResponse = objODIConverter.Request(strRequest)
//		objResponse.loadXML strResponse
//		Response.Write objResponse.documentElement.getAttribute("THREADID")
//		Response.Write "</TD></TR>"
//  
//		Response.Write "<TR><TD>EndTime: </TD><TD>" & Now & "</TD></TR>"
//		Response.Write "<TR><TD>Session ID: </TD><TD>" & Session.SessionId & "</TD></TR></TABLE>"
//		Set objODIConverter = Nothing
//		Set objResponse = Nothing
//		%>
//
//		Run the ASP simultaneously from mutiple browers on different machines. 
//		If ODIConverter is multi threading correctly, then the browsers should display 
//		overlapping start and end times, and different thread ids.
//
//		See MSDN article Q216580 for more details: Blocking/Serialization When Using 
//		InProc Component (DLL) from ASP. Due to IIS application debugging being enabled.
//
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
#include "RequestThreadWait.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestThreadWait::CRequestThreadWait(LPCWSTR szType) :
	CRequest(szType)
{
}

CRequestThreadWait::~CRequestThreadWait()
{
}


///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestThreadWait::Execute
//	
//	Description:
//		Virtual function implementation.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode
//			XML that must contain the attribute:
//				WAIT: How long to wait before returning a response.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			Response XML of the form "<RESPONSE THREADID=123 />"
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestThreadWait::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL;

	try
	{
		MSXML::IXMLDOMDocumentPtr ptrResponseDoc(__uuidof(MSXML::DOMDocument));

		int nWait = _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"WAIT"));

		long lThreadId = ::GetCurrentThreadId();
		::Sleep(nWait * 1000);

		ptrXMLDOMElement = ptrResponseDoc->createElement(L"RESPONSE");
		ptrResponseDoc->appendChild(ptrXMLDOMElement);
		ptrXMLDOMElement->setAttribute(L"THREADID", lThreadId);
		ptrXMLDOMElement->setAttribute(L"TYPE", L"SUCCESS");
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

	return ptrXMLDOMElement;
}
