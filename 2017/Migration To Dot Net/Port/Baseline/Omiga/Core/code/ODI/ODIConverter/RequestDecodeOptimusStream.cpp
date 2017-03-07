///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestDecodeOptimusStream.cpp
//	DESCRIPTION:	
//		Implements a method for converting a stream of serialized Optimus 
//		objects into the equivalent XML. This is useful where the stream 
//		has not been generated as a result of ODI calling into Optimus, but
//		has come from elsewhere, and you need decode the steam into XML.
//		
//		For example, Microsoft's Net Monitor can be used to capture the 
//		traffic between OSG and Optimus, including the request and response
//		streams (which contain serialized objects). You can take one of these
//		streams and decode it into XML, and then compare the result against
//		the XML ODI produces for the same type of requests/responses.
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
#include "odiconverter.h"
#include "ODIConverter1.h"
#include "MetaData.h"
#include "MetaDataEnvOptimus.h"
#include "OptimusResponse2XML.h"
#include "RequestDecodeOptimusStream.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestDecodeOptimusStream::CRequestDecodeOptimusStream(LPCWSTR szType) :
	CRequest(szType)
{
}

CRequestDecodeOptimusStream::~CRequestDecodeOptimusStream()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestDecodeOptimusStream::Execute
//	
//	Description:
//		Implementation of virtual method.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			XML. Must contain the attributes:
//				STREAM: The byte stream as comma/space separated two digit 
//				hexadecimal numbers, e.g., "00 00 00 CC FF FF F1...."
//				ODIENVIRONMENT: The name of the ODI environment to use for
//				decoding the byte stream. The ODI environment determines which
//				code page is used for converting payload data from Ebcdic to 
//				Unicode.
//	
//	Return:
//		IXMLDOMNodePtr: 
//			XML equivalent of the byte stream.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestDecodeOptimusStream::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		LPWSTR szStream			= static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"STREAM");
		LPCWSTR szMetaDataEnv	= static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getAttributeText(L"ODIENVIRONMENT");

		if (wcslen(szMetaDataEnv) == 0)
		{
			szMetaDataEnv = L"DEFAULT";
		}

		BYTE* pStream = new BYTE[wcslen(szStream)];

		if (pStream != NULL)
		{
			BYTE* pStreamPtr = pStream;
			int nTotalBytes = 0;
			static const wchar_t* pszDelimiters = L", \r\n\t";

			wchar_t* token = wcstok(szStream, pszDelimiters);
			while (token != NULL)
			{
				*pStreamPtr = hwtoi(token);
				pStreamPtr++;
				nTotalBytes++;
				token = wcstok(NULL, pszDelimiters);
			}

			CMetaDataEnvOptimus* pMetaDataEnv = reinterpret_cast<CMetaDataEnvOptimus*>(CMetaData::GetMetaDataEnv(szMetaDataEnv, L"OPTIMUS"));
			// Put a read lock on the meta data to prevent it changing for the duration of this request.
			namespaceMutex::CReadLock lckReadLock(pMetaDataEnv->GetSharedLock());
			ptrResponseNode = COptimusResponse2XML::Convert(pStream, nTotalBytes, pMetaDataEnv, pMetaDataEnv->GetCodePages().GetCodePage());

			delete []pStream;
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
