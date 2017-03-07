///////////////////////////////////////////////////////////////////////////////
//	FILE:			Request.cpp
//	DESCRIPTION:	
//		Base class for request classes. Each derived class implements a 
//		a type of request that can be called by the Request method on the 
//		ODIConverter.ODIConverter interface.
//
//		The derived request class must implement the virtual Execute function.
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
#include "Request.h"
#include "RequestOptimus.h"
#include "RequestGetCodePageTable.h"
#include "RequestDecodeOptimusStream.h"
#include "RequestStress.h"
#include "RequestThreadWait.h"
#include "RequestTransact.h"
#include "RequestConfigure.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequest::CRequest(LPCWSTR szName)
{ 
	SetName(szName); 
}

CRequest::~CRequest()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequest::Create
//	
//	Description:
//		Static virtual constructor for derived request classes.
//	
//	Parameters:
//		LPCWSTR szType:
//			The type of derived request object to construct.
//	
//	Return:
//		CRequest*: 	
//			Pointer to derived request object.
///////////////////////////////////////////////////////////////////////////////
CRequest* CRequest::Create(LPCWSTR szType)
{
	CRequest* pRequest = NULL;

	try
	{
		static const LPCWSTR pszRequestOptimus				= L"OPTIMUSREQUEST";
		static const LPCWSTR pszRequestTransact				= L"TRANSACTREQUEST";
		static const LPCWSTR pszRequestGetCodePageTable		= L"GETCODEPAGETABLE";
		static const LPCWSTR pszRequestDecodeOptimusStream	= L"DECODEOPTIMUSSTREAM";
		static const LPCWSTR pszRequestThreadWait			= L"THREADWAIT";
		static const LPCWSTR pszRequestStress				= L"STRESS";
		static const LPCWSTR pszRequestConfigure			= L"CONFIGURE";

		if (wcsicmp(szType, pszRequestOptimus) == 0)
		{
			pRequest = new CRequestOptimus(szType);
		}
		else if (wcsicmp(szType, pszRequestTransact) == 0)
		{
			pRequest = new CRequestTransact(szType);
		}
		else if (wcsicmp(szType, pszRequestGetCodePageTable) == 0)
		{
			pRequest = new CRequestGetCodePageTable(szType);
		}
		else if (wcsicmp(szType, pszRequestDecodeOptimusStream) == 0)
		{
			pRequest = new CRequestDecodeOptimusStream(szType);
		}
		else if (wcsicmp(szType, pszRequestThreadWait) == 0)
		{
			pRequest = new CRequestThreadWait(szType);
		}
		else if (wcsicmp(szType, pszRequestStress) == 0)
		{
			pRequest = new CRequestStress(szType);
		}
		else if (wcsicmp(szType, pszRequestConfigure) == 0)
		{
			pRequest = new CRequestConfigure(szType);
		}
		// ADDHERE Contruction of other request objects.
		else
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid request type: %s"), szType);
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
		throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid request type: %s"), szType);
	}

	return pRequest;
}

void CRequest::WriteBuffer(LPCTSTR lpszFName, LPBYTE lpBuffer, int nBufferSize)
{
	try
	{
		_ASSERTE(lpBuffer != NULL);

		if (lpBuffer != NULL)
		{
			_TCHAR szPath[_MAX_PATH] = _T("");
			_stprintf(szPath, _T("%s%s%s.dat"), _Module.m_szDrive, _Module.m_szDir, lpszFName);
			FILE* fp = _tfopen(szPath, _T("wb"));
			fwrite(lpBuffer, sizeof(BYTE), nBufferSize, fp);
			fclose(fp);
		}
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}
}

void CRequest::WriteXML(LPCTSTR lpszFName, MSXML::IXMLDOMNodePtr ptrNode)
{
	try
	{
		_ASSERTE(ptrNode != NULL);

		if (ptrNode != NULL)
		{
			MSXML::IXMLDOMDocumentPtr ptrDOMDocument(__uuidof(MSXML::DOMDocument));;
			_TCHAR szXMLPath[_MAX_PATH] = _T("");
			_stprintf(szXMLPath, _T("%s%s%s.xml"), _Module.m_szDrive, _Module.m_szDir, lpszFName);
			ptrDOMDocument->appendChild(ptrNode);
			ptrDOMDocument->save(szXMLPath);
		}
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}
}

void CRequest::LogProfile(const MSXML::IXMLDOMNodePtr ptrNode)
{
	try
	{
		MSXML::IXMLDOMElementPtr ptrElement(ptrNode->selectSingleNode(L"//RESPONSE/PROFILE"));

		if (ptrElement != NULL)
		{
			_Module.LogDebug(_T("Profile:\n"));
			_Module.LogDebug(_T("Tag,Hits,Time(ms)\n"));

			MSXML::IXMLDOMNodeListPtr ptrNodeList = ptrElement->GetchildNodes();
			for (int nChild = 0; nChild < ptrNodeList->Getlength(); nChild++)
			{
				MSXML::IXMLDOMElementPtr ptrElementChild = ptrNodeList->Getitem(nChild);
				_bstr_t bstrTagName = ptrElementChild->GettagName();

				MSXML::IXMLDOMElementPtr ptrElementHits = ptrElementChild->selectSingleNode(L"HITS");
				_bstr_t bstrHits = ptrElementHits->Gettext();

				MSXML::IXMLDOMElementPtr ptrElementTime = ptrElementChild->selectSingleNode(L"TOTALTIME");
				_bstr_t bstrTime = ptrElementTime->Gettext();

				_Module.LogDebug(_T("%s,%s,%s\n"), static_cast<LPCWSTR>(bstrTagName), static_cast<LPCWSTR>(bstrHits), static_cast<LPCWSTR>(bstrTime));
			}
		}
	}
	catch(_com_error& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}
}


