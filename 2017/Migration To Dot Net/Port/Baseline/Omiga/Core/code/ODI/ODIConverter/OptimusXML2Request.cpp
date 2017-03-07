///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusXML2Request.cpp
//	DESCRIPTION:	
//		This class converts an XML request in an Optimus byte stream. All
//		members are static; this class should not be instantiated.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      08/08/01    Initial version
//  STB		28/05/02	SYS4558 Re-calculate the length of the ACTUAL data.
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CodePage.h"
#include "Exception.h"
#include "LookUpTable.h"
#include "MetaDataEnvOptimus.h"
#include "OptimusMetaData.h"
#include "OptimusObject.h"
#include "ODIConverter.h"
#include "Profiler.h"
#include "OptimusXML2Request.h"


///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusXML2Request::Convert
//	
//	Description:
//		Converts an XML request into an Optimus byte stream.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			The XML request node.
//		BYTE* pbRequest:
//			Buffer to receive Optimus byte stream.
//		const size_t nBufferSize:
//			Maximum size of pbRequest.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property names used in the XML to the ids used in the byte stream 
//			sent to Optimus.
//		CCodePage* pCodePage:
//			The code page to be used in converting the Unicode payload data in 
//			the XML	request into Ebcdic in the byte stream.
//	
//	Return:
//		size_t: 	
//			The actual length of the resulting Optimus byte stream.
///////////////////////////////////////////////////////////////////////////////
size_t COptimusXML2Request::Convert(
	const MSXML::IXMLDOMNodePtr ptrRequestNode, 
	BYTE* pbRequest, 
	const size_t nBufferSize, 
	CMetaDataEnvOptimus* pMetaDataEnv,
	CCodePage* pCodePage)
{
	size_t nRequestSize;

	try
	{
		if (pMetaDataEnv->GetLogDebug())
		{
			_Module.LogDebug(_T("Request:\n"));
			COptimusObject::LogDebugHeader();
		}

		long lNumbering = 0;
		BYTE* pbRequestSave = pbRequest;
		nRequestSize = ConvertDOMElement(ptrRequestNode, pbRequest, nBufferSize, -1, lNumbering, -1, pMetaDataEnv, pCodePage);
		pbRequest = pbRequestSave;
		pbRequest[nRequestSize++] = 0x0D; // Terminator required by Optimus.
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nRequestSize;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusXML2Request::ConvertDOMElement
//	
//	Description:
//		Converts a single element in the request XML into a record appended
//		onto the end of the existing Optimus byte stream.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			The current element in the XML request.
//		BYTE* pbRequest:
//			Buffer to receive Optimus byte stream.
//		const size_t nBufferSize:
//			Maximum size of pbRequest.
//		long lParentNumber:
//			Number of the parent object for the current object/property.
//		long& lNumbering:
//			Unique number for the new object/property record that will be
//			appended onto the byte stream. Each record is given an unique 
//			number within the byte stream. The number is incremented by one
//			for each successive record in the byte stream.
//		short nParentId:
//			The unique Optimus id of the parent object for this element. The
//			parent id of the root object in the request is -1.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property names used in the XML to the ids used in the byte stream 
//			sent to Optimus.
//		CCodePage* pCodePage:
//			The code page to be used in converting the Unicode payload data in 
//			the XML	request into Ebcdic in the byte stream.
//	
//	Return:
//		size_t: 
//			The length of the Optimus byte stream after converting the element.
///////////////////////////////////////////////////////////////////////////////
size_t COptimusXML2Request::ConvertDOMElement(
	const MSXML::IXMLDOMElementPtr ptrXMLDOMElement, 
	BYTE*& pbRequest, 
	const size_t nBufferSize, 
	long lParentNumber, 
	long& lNumbering, 
	short nParentId,
	CMetaDataEnvOptimus* pMetaDataEnv,
	CCodePage* pCodePage)
{
	size_t nRequestSize	= 0;

	try
	{
		short nLength		= 0;
		long lNumber		= lNumbering++;
		short nId			= 0;
		short nNewParentId	= 0;

		if (ptrXMLDOMElement == NULL)
		{
			throw CException(E_NULLELEMENT, __FILE__, __LINE__, _T("%s"), _T("Element is NULL"));
		}

		nLength = sizeof(nLength) + sizeof(lNumber) + sizeof(lParentNumber) + sizeof(nId);

		if (nBufferSize < nLength)
		{
			throw CException(E_BUFFEROVERRUN, __FILE__, __LINE__, _T("%s"), _T("Request buffer is too small"));
		}

		// Save pointer, as put functions move it.
		BYTE* pbRequestStart = pbRequest;
		PutShort(pbRequest, nLength);
		PutLong(pbRequest, lNumber);
		PutLong(pbRequest, lParentNumber);

		// Need to copy tag name otherwise it will be overwritten.
		_bstr_t bstrTag = ptrXMLDOMElement->GettagName();
		LPCWSTR szTag	= bstrTag; 
		bool bProperty	= false;
		CLookUpTable* pLookUpTable = NULL;

		/*
		// AS 26/02/02 This is erroneously picking up some properties as objects, e.g.,
		// TARGETKEY in CUSTOMERINVOLVEMENTPATTERN, i.e., properties CAN have child nodes
		// in the request XML.
		if (ptrXMLDOMElement->firstChild == NULL)
		{
			// Element does not have any children, so it could be a childless object or a property.
			// If the element did have children it would definitely be an object and not a property.
			// See if the element is in the property map.
			nId = nNewParentId = pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyIdPropertyId(pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyId(nParentId, szTag, pLookUpTable));
		}
		*/

		// See if the element is in the property map.
		nId = nNewParentId = pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyIdPropertyId(pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyId(nParentId, szTag, pLookUpTable));

		if (nId == 0)
		{
			// Element is not a property so look for it in the object map.
			nId = nNewParentId = pMetaDataEnv->GetOptimusMetaData().GetObjectId(szTag);
			if (nId == 0)
			{
				// Element is not in the object map. However, some objects only appear as properties, 
				// so check the property map. We only need to check for elements that have children
				// (if it didn't have children the property map will already have been searched).
				if (ptrXMLDOMElement->firstChild != NULL)
				{
					nId = nNewParentId = pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyIdPropertyId(pMetaDataEnv->GetOptimusMetaData().GetObjectPropertyId(nParentId, szTag, pLookUpTable));
				}
				else
				{
					throw CException(E_INVALIDTYPE, __FILE__, __LINE__, _T("Invalid id %d for object %s"), nId, szTag);
				}
			}
		}
		else
		{
			bProperty = true;

			// Element is a property, but check if it is also an object, in which case the object id 
			// is also the parent id for properties of this object.
			nNewParentId = pMetaDataEnv->GetOptimusMetaData().GetObjectId(szTag);
			if (nNewParentId == 0)
			{
				// Element is not an object, so the parent id is the property id.
				nNewParentId = nId;
			}
		}
		PutShort(pbRequest, nId);

		_bstr_t bstrData = L"";
		if (bProperty)
		{
			short nLengthEbcdic		= 0;
			short nLengthUnicode	= 0;

			_variant_t varData = ptrXMLDOMElement->getAttribute(L"DATA");
			if (varData.vt != VT_NULL)
			{
				bstrData = varData.bstrVal;
			}

			nLengthUnicode = wcslen(bstrData);
			if (nLengthUnicode > 0)
			{
				if (pLookUpTable != NULL)
				{
					// This item is a property with a look up table associated with it,
					// so look up conversion value for the value of the property.
					LPCWSTR szValue = pLookUpTable->GetValue(bstrData, CLookUpTable::dirSend);
					if (wcslen(szValue) > 0)
					{
						// There is an entry in the look up table for the current value of this property,
						// so replace the current value with the value from the look up table.
						bstrData = szValue;

						// SYS4558 - Re-calculate the length of the ACTUAL data.
						nLengthUnicode = wcslen(bstrData);
					}
					else
					{
						_Module.LogEventWarning(_T("Lookup failure: LOOKUPTABLE\\@NAME=\"%s\" ITEM\\@SENDIN=\"%s\""), static_cast<LPWSTR>(pLookUpTable->GetName()), static_cast<LPWSTR>(bstrData));
					}
				}

				if (nBufferSize < nLength + (nLengthUnicode / 2))
				{
					throw CException(E_BUFFEROVERRUN, __FILE__, __LINE__, _T("Request buffer is too small"));
				}

				nLengthEbcdic = pCodePage->SendW2MBC(bstrData, nLengthUnicode, reinterpret_cast<LPSTR>(pbRequest), nBufferSize - nLength);

				if (nLengthEbcdic <= 0)
				{
					throw CException(E_UNICODE2EBCDIC, __FILE__, __LINE__, _T("Could not convert UNICODE to EBCDIC"));
				}
			}

			nLength += nLengthEbcdic;
			pbRequest = pbRequestStart;
			PutShort(pbRequest, nLength);
		}

		if (pMetaDataEnv->GetLogDebug())
		{
			COptimusObject OptimusObject(pbRequestStart, szTag, bstrData);
			OptimusObject.LogDebug(pMetaDataEnv->GetLogHex());
		}

		pbRequest = pbRequestStart + nLength;

		MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMElement->GetchildNodes();

		nRequestSize = nLength;
		for (int nChild = 0; nChild < ptrXMLDOMNodeList->Getlength(); nChild++)
		{
			nRequestSize += ConvertDOMElement(ptrXMLDOMNodeList->Getitem(nChild), pbRequest, nBufferSize, lNumber, lNumbering, nNewParentId, pMetaDataEnv, pCodePage);
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

	return nRequestSize;
}
