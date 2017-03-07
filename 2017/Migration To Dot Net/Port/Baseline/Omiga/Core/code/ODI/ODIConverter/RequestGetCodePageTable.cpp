///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestGetCodePageTable.cpp
//	DESCRIPTION:	
//		For a given Microsoft code page number, this class returns a code page
//		table that can be used by ODIConverter for converting data to and from the 
//		code page.
//
//		This request type is useful for generating a code page table on a
//		machine on which the code page is installed in the Windows operating
//		system. The code page table can then be added to the ODIConverter meta 
//		data (as a CodePageMap XML), and used on a machine which does not have 
//		the code page installed in the Windows operating system.
//
//		For example, the production server may not have the necessary Ebcdic
//		code pages installed in Windows in order for ODIConverter to do the
//		necessary conversion from Unicode to Ebcdic that is required by Optimus.
//		If the code pages are installed ODIConverter will use them. However, if
//		the code pages are not installed, ODIConverter can use a code page table
//		stored in its CodePageMap meta data, removing any reliance on Windows.
//		The CodePageMap will have to have been generated on a machine that did
//		have the correct code pages installed in Windows.
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
#include "CodePage.h"
#include "Exception.h"
#include "ODIConverter.h"
#include "ODIConverter1.h"
#include "RequestGetCodePageTable.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRequestGetCodePageTable::CRequestGetCodePageTable(LPCWSTR szType) :
	CRequest(szType)
{
}

CRequestGetCodePageTable::~CRequestGetCodePageTable()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CRequestGetCodePageTable::Execute
//	
//	Description:
//		Virtual function implementation.
//	
//	Parameters:
//		const IXMLDOMNodePtr ptrRequestNode:
//			XML that must contain the attributes:
//				CODEPAGE: The number of Microsoft code page for which the code
//				page table will be generated.
//				CODEPAGESIZE: The maximum size of the code page table.
//	
//	Return:
//		IXMLDOMNodePtr: 	
//			XML response, in the form, e.g.,
//
//				<CODEPAGE MSCODEPAGE="37">00 01 02 03 37...</CODEPAGE>
//
//			To use this XML in the ODIConverter CodePageMap meta data, add
//			COUNTRY and LANGUAGE attributes, which correspond to the Optimus
//			LOCALES for which you want this code page table to be used, e.g.,
//
//				<CODEPAGE COUNTRY="UK" LANGUAGE="EN" MSCODEPAGE="37">00 01 02 03 37...</CODEPAGE>
///////////////////////////////////////////////////////////////////////////////
MSXML::IXMLDOMNodePtr CRequestGetCodePageTable::Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode)
{
	MSXML::IXMLDOMNodePtr ptrResponseNode = NULL;

	try
	{
		int nCodePage		= _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"CODEPAGE"));
		int nCodePageSize	= _wtoi(static_cast<IXMLAssistDOMNodePtr>(ptrRequestNode)->getMandatoryAttributeText(L"CODEPAGESIZE"));

		if (nCodePage != 0 && nCodePageSize != 0)
		{
			unsigned short* pCodePageTable = new unsigned short[nCodePageSize];

			if (pCodePageTable != NULL)
			{
				CCodePage::GetCodePageTable(nCodePage, pCodePageTable, nCodePageSize);

				MSXML::IXMLDOMDocumentPtr ptrDOMDocumentCodePageTable(__uuidof(MSXML::DOMDocument));

				if (ptrDOMDocumentCodePageTable != NULL)
				{
					MSXML::IXMLDOMElementPtr ptrXMLDOMElementResponse = ptrDOMDocumentCodePageTable->createElement(L"RESPONSE");
					ptrDOMDocumentCodePageTable->appendChild(ptrXMLDOMElementResponse);
					ptrXMLDOMElementResponse->setAttribute(L"TYPE", L"SUCCESS");

					MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrDOMDocumentCodePageTable->createElement(L"CODEPAGE");
					ptrXMLDOMElementResponse->appendChild(ptrXMLDOMElement);
					wchar_t szCodePage[16];
					swprintf(szCodePage, L"%d", nCodePage);
					ptrXMLDOMElement->setAttribute(L"MSCODEPAGE", szCodePage);

					// Each character code can take up to 4 bytes (maximum unsigned short is FFFF), plus
					// extra character for comma separator.
					wchar_t* pszCodePageTable = new wchar_t[(nCodePageSize * 5) + 1];

					if (pszCodePageTable != NULL)
					{
						pszCodePageTable[0] = L'\0';

						for (int nChar = 0; nChar < nCodePageSize; nChar++)
						{
							wchar_t szChar[16];
							swprintf(szChar, L"%02X", pCodePageTable[nChar]);
							if (wcslen(pszCodePageTable) > 0)
							{
								wcscat(pszCodePageTable, L" ");
							}
							wcscat(pszCodePageTable, szChar);
						}
						ptrXMLDOMElement->text = pszCodePageTable;

						delete []pszCodePageTable;

						ptrResponseNode = ptrDOMDocumentCodePageTable->documentElement;
					}
				}

				delete []pCodePageTable;
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

