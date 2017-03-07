///////////////////////////////////////////////////////////////////////////////
//	FILE:			ODIConverter1.cpp
//	DESCRIPTION:	
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
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "MetaData.h"
#include "Request.h"


/////////////////////////////////////////////////////////////////////////////
// CODIConverter1

STDMETHODIMP CODIConverter1::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IODIConverter
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CODIConverter1::Request(BSTR bstrRequest, BSTR *pbstrResponse)
{
	HRESULT hr = S_OK;

	*pbstrResponse = NULL;

	try
	{
		*pbstrResponse = Request(bstrRequest).copy();
	}
	catch(_com_error& e)
	{
		CException Exception(E_GENERICERROR, e, __FILE__, __LINE__);
		hr = LogError(Exception);
	}
	catch(CException& e)
	{
		hr = LogError(e);
	}
	/*
	// AS 22/01/2007 VS2005 Port.
	// error C2316: 'exception &' : cannot be caught as the destructor and/or copy constructor are inaccessible.
	catch(exception& e)
	{
		CException Exception(E_GENERICERROR, e, __FILE__, __LINE__);
		hr = LogError(Exception);
	}
	*/
	catch(...)
	{
		CException Exception(E_GENERICERROR, __FILE__, __LINE__, _T("Unknown error"));
		hr = LogError(Exception);
	}

	return hr;
}

_bstr_t CODIConverter1::Request(LPCWSTR szRequest)
{
	_bstr_t bstrResponse(L"");

	try
	{
		MSXML::IXMLDOMDocumentPtr ptrInDoc(__uuidof(MSXML::DOMDocument));
		MSXML::IXMLDOMNodePtr ptrRequestNode = NULL;
		MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

		ptrInDoc->async = false;

		// Load request.
		_variant_t varbLoaded = ptrInDoc->loadXML(szRequest);
		if (varbLoaded.boolVal == false)
		{
			throw CException(E_LOADXML, __FILE__, __LINE__, _T("Unable to load ODIConverter request XML"));
		}

		// Check this is a REQUEST node.
		ptrRequestNode = ptrInDoc->selectSingleNode(L"//REQUEST");
		if (ptrRequestNode == NULL)
		{
			throw CException(E_INVALIDREQUEST, __FILE__, __LINE__, _T("Invalid ODIConverter request XML"));
		}

		if (ptrRequestNode->Getattributes() != NULL && ptrRequestNode->Getattributes()->getNamedItem(L"OPERATION") != NULL)
		{
			// OPERATION attribute on REQUEST node. Do single operation.
			ptrResponseNode = DoRequest(ptrRequestNode);
		}
		else
		{
			// Multiple operations.
			MSXML::IXMLDOMNodeListPtr ptrOperationNodeList = ptrInDoc->selectNodes(L"//REQUEST/OPERATION");

			for (long lNode = 0; lNode < ptrOperationNodeList->length; lNode++)
			{
				MSXML::IXMLDOMNodePtr ptrOperationNode = ptrOperationNodeList->item[lNode];

				// Copy attributes from REQUEST node to OPERATION node.
				MSXML::IXMLDOMNamedNodeMapPtr ptrAttributes = ptrRequestNode->Getattributes();
				for (long lAttribute = 0; lAttribute < ptrAttributes->length; lAttribute++)
				{
					MSXML::IXMLDOMAttributePtr ptrAttribute = ptrAttributes->item[lAttribute];

					ptrOperationNode->Getattributes()->setNamedItem(ptrAttribute->cloneNode(true));
				}
				ptrResponseNode = DoRequest(ptrOperationNode);
			}
		}

		if (ptrResponseNode != NULL)
		{
			bstrResponse = ptrResponseNode->xml;
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

	return bstrResponse;
}

MSXML::IXMLDOMNodePtr CODIConverter1::DoRequest(MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		LPCWSTR szOperation = L"";

		if (_wcsicmp(ptrRequestNode->nodeName, L"REQUEST") == 0)
		{
			// Get operation (function) name from REQUEST/@OPERATION.
			szOperation = static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"OPERATION");
		}
		else
		{
			// Get operation (function) name from OPERATION/@NAME.
			szOperation = static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"NAME");
		}

		if (wcslen(szOperation) > 0)
		{
			CRequest* pRequest = CMetaData::GetRequest(szOperation);

			if (pRequest != NULL)
			{
				ptrResponseNode = pRequest->Execute(ptrRequestNode);
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

	return ptrResponseNode;
}


HRESULT CODIConverter1::LogError(CException& e)
{
	long lLogError = e.GetError();
	long lClientError = lLogError;
	HRESULT hLogResult = e.GetHResult();
	HRESULT hClientResult = hLogResult;
	if (lLogError >= WSABASEERR && lLogError <= WSA_QOS_GENERIC_ERROR)
	{
		// The event log will show the actual socket error.
		// Return a generic socket error (E_WSASOCKET) to the client.
		// Thus we do not have to define all the WINSOCK2.h errors in the Omiga IV database.
		lClientError = E_WSASOCKET;
		hClientResult = MAKE_HRESULT(SEVERITY_ERROR, FACILITY_ITF, lClientError);
	}

	LPCWSTR lpszFile	= e.GetFile();
	int nLine			= e.GetLine();
	if (lpszFile != NULL && wcslen(lpszFile) > 0 && nLine > 0)
	{
		_Module.LogEventError(_T("%s(%d) : %s exception: %s(%d)"), lpszFile, nLine, e.GetType(), e.GetDescription(), lLogError);
	}
	else
	{
		_Module.LogEventError(_T("%s exception: %s(%d)"), e.GetType(), e.GetDescription(), lLogError);
	}

	return Error(e.GetDescription(), e.GetIID() != GUID_NULL ? e.GetIID() : IID_IODIConverter, hClientResult);
}
