///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusResponse2XML.cpp
//	DESCRIPTION:	
//		This class converts an Optimus response (a byte stream) into XML. All
//		members are static; this class should not be instantiated.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      02/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <vector>
#include "CodePage.h"
#include "Exception.h"
#include "LookUpTable.h"
#include "MetaDataEnvOptimus.h"
#include "OptimusMetaData.h"
#include "OptimusObject.h"
#include "OptimusResponse2XML.h"
#include "ODIConverter.h"
#include "Profiler.h"


//#define DEBUG_ATTRIBUTES
#ifdef DEBUG_ATTRIBUTES
	#define DEBUG_SETNUMBERING(ptrXMLDOMElement, lNumber) ptrXMLDOMElement->setAttribute(L"DEBUG_NUMBERING", lNumber)
	#define DEBUG_SETOFFSET(ptrXMLDOMElement, lOffset) ptrXMLDOMElement->setAttribute(L"DEBUG_OFFSET", lOffset)
	#define DEBUG_SETPARENT(ptrXMLDOMElement, lParent) ptrXMLDOMElement->setAttribute(L"DEBUG_PARENT", lParent)
#else
	#define DEBUG_SETNUMBERING(ptrXMLDOMElement, lNumber) void(0)
	#define DEBUG_SETOFFSET(ptrXMLDOMElement, lOffset) void(0)
	#define DEBUG_SETPARENT(ptrXMLDOMElement, lParent) void(0)
#endif

///////////////////////////////////////////////////////////////////////////////

// Class wrapper for vector of IXMLDOMElementPtr; ensures interface pointers are
// dereferenced properly, especially if an exception is thrown.
class CDOMElements
{
public:
	typedef std::vector<MSXML::IXMLDOMElementPtr> vecXMLDOMElementPtrType;

private:
	vecXMLDOMElementPtrType m_vecXMLDOMElementPtr;

public:
	~CDOMElements() 
	{
		// Clear down the vector (must set interface pointers to NULL).
		for (vecXMLDOMElementPtrType::iterator it = m_vecXMLDOMElementPtr.begin(); it < m_vecXMLDOMElementPtr.end(); it++)
		{
			 *it = NULL;
		}
		m_vecXMLDOMElementPtr.clear();
	}
	vecXMLDOMElementPtrType& GetDOMElements() { return m_vecXMLDOMElementPtr; }
};


///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusResponse2XML::Convert
//	
//	Description:
//		Converts an Optimus response byte stream into XML.
//	
//	Parameters:
//		BYTE* pbResponse:
//			Buffer containing Optimus byte stream response.
//		size_t nSize:
//			Actual length of the byte stream.
//		CMetaDataEnvOptimus* pMetaDataEnv:
//			The ODIConverter meta data environment. In particular, this holds
//			the Optimus object map meta data, that maps Optimus object and 
//			property id used in the byte stream to the names used in the XML
//			returned from ODI.
//		CCodePage* pCodePage:
//			The code page to be used in converting the Ebcdic payload data in 
//			the byte stream response into Unicode in the XML response.
//	
//	Return:
//		IXMLDOMNodePtr: 	
//			XML response.
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr COptimusResponse2XML::Convert(
	BYTE* pbResponse, 
	size_t nSize, 
	CMetaDataEnvOptimus* pMetaDataEnv,
	CCodePage* pCodePage)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	// Vectors which contain xml elements / types and is indexed by lNumber (see below).
	size_t nToReserve = 300; 

	MSXML::IXMLDOMDocumentPtr ptrResponseDoc(__uuidof(MSXML::DOMDocument));
	CDOMElements DOMElements;
	CDOMElements::vecXMLDOMElementPtrType vecXMLDOMElementPtr = DOMElements.GetDOMElements();
	vecXMLDOMElementPtr.reserve(nToReserve);

	typedef std::vector<short> vecTypeType;
	vecTypeType vecType;
	vecType.reserve(nToReserve);

	if (pMetaDataEnv->GetLogDebug())
	{
		_Module.LogDebug(_T("Response:\n"));
		COptimusObject::LogDebugHeader();
	}

	try
	{
		BYTE* pbCurrent = pbResponse;
		BYTE* pbEnd = pbResponse + nSize - 1;
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL;

		// buffer used for all conversion from EBCDIC to unicode
		const int buffUnicodeDataSize = 0xFFF;
		wchar_t buffUnicodeData[buffUnicodeDataSize];

		while (pbCurrent < pbEnd)
		{
			// first four values are short, long, long, short
			const int nMinLength = sizeof(short) + sizeof(long) + sizeof(long) + sizeof(short);
#ifdef DEBUG_ATTRIBUTES
			long lDebugOffset = pbCurrent - pbResponse;
#endif	
			BYTE* pbCurrentStart = pbCurrent;
			short nLength = GetShort(pbCurrent);
#ifdef _DEBUG
			// _DEBUG required to prevent release build warning.
			long lNumber = GetLong(pbCurrent);
#else
			GetLong(pbCurrent);
#endif
			long lParent = GetLong(pbCurrent);
			short nType = GetShort(pbCurrent);
			LPCWSTR szTag = L"";
			CLookUpTable* pLookUpTable = NULL;

			// create element from the data (either an object or an argument)
			if (nType > 0)
			{
				szTag = pMetaDataEnv->GetOptimusMetaData().GetObjectName(nType);
			}
			else
			{
 				szTag = pMetaDataEnv->GetOptimusMetaData().GetPropertyName(vecType[lParent], nType, pLookUpTable);
			}
			ptrXMLDOMElement = ptrResponseDoc->createElement(szTag);

			DEBUG_SETNUMBERING(ptrXMLDOMElement, lNumber);
			DEBUG_SETOFFSET(ptrXMLDOMElement, lDebugOffset);
			DEBUG_SETPARENT(ptrXMLDOMElement, lParent);
			// ... add it into the vectors
			vecXMLDOMElementPtr.push_back(ptrXMLDOMElement);
			vecType.push_back(nType);

#ifdef _DEBUG
			// _DEBUG required to prevent release build warning.
			_ASSERTE(lNumber + 1 == vecXMLDOMElementPtr.size());
			_ASSERTE(lNumber + 1 == vecType.size());
#endif
			// ... and add it to the DOM
			if (lParent == -1) // top
			{
				ptrResponseDoc->appendChild(ptrXMLDOMElement);
				ptrResponseNode = ptrXMLDOMElement;
			}
			else
			{
				vecXMLDOMElementPtr[lParent]->appendChild(ptrXMLDOMElement);
			}

			// ... fifth value is EBCDIC data (if any)
			short nLengthEBCDIC = nLength - nMinLength;
			buffUnicodeData[0] = L'\0';
			if (nLengthEBCDIC > 0)
			{
				int nResult = 0;
				if (pCodePage != NULL)
				{
					nResult = pCodePage->RecvMBC2W(reinterpret_cast<LPCSTR>(pbCurrent), nLengthEBCDIC, buffUnicodeData, buffUnicodeDataSize);
				}

				if (nResult > 0)
				{
					buffUnicodeData[nResult] = L'\0'; // must null terminate

					_bstr_t bstrData = buffUnicodeData;
					if (pLookUpTable != NULL)
					{
						// This item is a property with a look up table associated with it,
						// so look up conversion value for the value of the property.
						LPCWSTR szValue = pLookUpTable->GetValue(bstrData, CLookUpTable::dirRecv);
						if (wcslen(szValue) > 0)
						{
							// There is an entry in the look up table for the current value of this property,
							// so replace the current value with the value from the look up table.
							bstrData = szValue;
						}
						else
						{
							_Module.LogEventWarning(_T("Lookup failure: LOOKUPTABLE\\@NAME=\"%s\" ITEM\\@RECVIN=\"%s\""), static_cast<LPWSTR>(pLookUpTable->GetName()), static_cast<LPWSTR>(bstrData));
						}
					}

					ptrXMLDOMElement->setAttribute(L"DATA", bstrData);
				}
				else
				{
					throw CException(E_EBCDIC2UNICODE, __FILE__, __LINE__, _T("%s"), _T("Could not convert EBCDIC to UNICODE"));
				}

				pbCurrent += nLengthEBCDIC;
			}

			if (pMetaDataEnv->GetLogDebug())
			{
				COptimusObject OptimusObject(pbCurrentStart, szTag, buffUnicodeData);
				OptimusObject.LogDebug(pMetaDataEnv->GetLogHex());
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
