///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestConfigure.cpp
//	DESCRIPTION:	
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
#include "MetaData.h"
#include "ODIConverter.h"
#include "RequestConfigure.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestConfigure::CRequestConfigure(LPCWSTR szType) :
	CRequest(szType)
{ 
}

CRequestConfigure::~CRequestConfigure()
{
}

MSXML::IXMLDOMNodePtr CRequestConfigure::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		bool bSave = IsTrueString(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"SAVE"));

		CMetaData::InitMetaDataEnvs(ptrRequestNode);

		if (bSave)
		{
			_bstr_t bstrSaveAs = static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"SAVEAS");

			if (wcslen(bstrSaveAs) == 0)
			{
				bstrSaveAs = CMetaData::GetODIConverterXMLPath();
			}
			CMetaData::Save(bstrSaveAs);
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

	MSXML::IXMLDOMDocumentPtr ptrResponseDoc(__uuidof(MSXML::DOMDocument));
	MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrResponseDoc->createElement(L"RESPONSE");
	ptrResponseDoc->appendChild(ptrXMLDOMElement);
	ptrXMLDOMElement->setAttribute(L"TYPE", L"SUCCESS");

	return ptrResponseNode;
}

