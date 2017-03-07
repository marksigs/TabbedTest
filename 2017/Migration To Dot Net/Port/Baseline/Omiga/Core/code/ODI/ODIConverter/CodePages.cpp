// CodePages.cpp: implementation of the CCodePages class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "ODIConverter.h"
#include "CodePages.h"
#include "CodePage.h"

static LPCWSTR g_pszEnableEuroKey						= L"//CODEPAGES/ENABLEEURO";
static LPCWSTR g_pszCodePageKey							= L"//CODEPAGES/CODEPAGE";
static LPCWSTR g_pszCountry								= L"COUNTRY";
static LPCWSTR g_pszLanguage							= L"LANGUAGE";
static LPCWSTR g_pszEuro								= L"EURO";
static LPCWSTR g_pszDefault								= L"DEFAULT";
static LPCWSTR g_pszMSCodePage							= L"MSCODEPAGE";
static LPCWSTR g_pszIBMCCSID							= L"IBMCCSID";
static LPCWSTR g_pszOperation							= L"OPERATION";
static LPCWSTR g_pszSet									= L"SET";
static LPCWSTR g_pszInsert								= L"INSERT";
static LPCWSTR g_pszUpdate								= L"UPDATE";
static LPCWSTR g_pszDelete								= L"DELETE";

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCodePages::CCodePages() :
	m_bEnableEuro(false),
	m_bFoundDefault(false),
	m_bstrCodePagesPath(L"")
{
}

CCodePages::~CCodePages()
{
	// Free all allocated objects.
	if (m_MapLocaleToCodePage.size() > 0)
	{
		MapLocaleToCodePageType::iterator it;
		for (it = m_MapLocaleToCodePage.begin(); it != m_MapLocaleToCodePage.end(); it++)
		{
			CCodePage* pCodePage = (*it).second;

			delete pCodePage;
		}
		m_MapLocaleToCodePage.clear();
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePages::InitCodePageMap 
//	 
//	Description:  
//		Initialises the code page information for this meta data environment.
//	 
//	Parameters:  
//		LPCWSTR szCodePagesPath:    
//			File name of the XML file that contains the code page information.   
//			e.g.,
//
//			<CODEPAGES>
//				<ENABLEEURO>N</ENABLEEURO>
//				<CODEPAGE COUNTRY="UK" LANGUAGE="EN" EURO="Y" MSCODEPAGE="1146" IBMCCSID="1146"/>
//					00 01 02 03 37 2D 2E 2F 16 05 25 0B 0C 0D 0E 0F 10 11 12 13
//				</CODEPAGE>
//			</CODEPAGES>
//
//			Code pages are used to convert all data being sent to and from the 
//			remote machine. For example, Optimus expects data in Ebcdic format,
//			so ODIConverter needs to convert the Unicode data sent from Omiga
//			into Ebcdic. This conversion is done using Microsoft's Ebcdic code
//			pages.
//
//			CODEPAGES/ENABLEEURO:
//				If Y then only code pages with the EURO attribute set to "Y" 
//				will be used by ODIConverter. Thus ENABLEEURO is a way of
//				globally switching on support for the Euro in this meta data
//				environment.
//			CODEPAGES/CODEPAGE:
//				Define a CODEPAGE node for each locale. The locale consists of
//				the country code and the language code. The associated code page
//				defines the conversion to be applied for data in that locale.
//			CODEPAGE/@COUNTRY:
//				The country code for the locale.
//			CODEPAGE/@LANGUAGE:
//				The language code for the locale.
//			CODEPAGE/@EURO:
//				If "Y", then this code page has the Euro defined.
//			CODEPAGE/@MSCODEPAGE:
//				The Microsoft Code Page number for this locale.
//			CODEPAGE/@IBMCCSID:
//				The IBM number for this locale. This attribute is ignored by 
//				ODIConverter.
//			CODEPAGE/@DEFAULT:
//				If "Y", then this is the default code page that will be used
//				if none of the code pages matches the locale.
//			CODEPAGE/TEXT:
//				Optional translation table; useful when the MSCODEPAGE is not
//				installed on the run time machine. See CodePage.cpp for details.
//	 
//	Return:  
//		None.
///////////////////////////////////////////////////////////////////////////////
long CCodePages::Init(LPCWSTR szCodePagesPath)
{
	long lCodePages = 0;

	try
	{
		// initialise from XML file
		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement;

		if (_taccess(szCodePagesPath, 00) != 0)
		{
			throw CException(E_FILEDOESNOTEXIST, __FILE__, __LINE__, _T("File does not exist: %s"), szCodePagesPath);
		}

		ptrXMLDOMDocument->async = false;
		_variant_t varLoaded = ptrXMLDOMDocument->load(szCodePagesPath);

		if (varLoaded.boolVal == false)
		{
			throw CException(E_INVALIDMETADATAFILE, __FILE__, __LINE__, _T("Invalid meta data file: %s"), szCodePagesPath);
		}

		m_bstrCodePagesPath = szCodePagesPath;

		lCodePages = InitNode(ptrXMLDOMDocument->documentElement, szCodePagesPath);
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

	return lCodePages;
}

long CCodePages::InitNode(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szCodePagesPath)
{
	long lIndex = 0;

	try
	{
		if (szCodePagesPath == NULL)
		{
			szCodePagesPath = L"Unknown source";
		}

		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMNode->selectSingleNode(g_pszEnableEuroKey);
		if (ptrXMLDOMElement != NULL)
		{
			m_bEnableEuro = IsTrueString(ptrXMLDOMElement->Gettext());
		}

		MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMNode->selectNodes(g_pszCodePageKey);

		long lLength = ptrXMLDOMNodeList->length;

		while (lIndex < lLength)
		{
			ptrXMLDOMElement = ptrXMLDOMNodeList->item[lIndex];

			if (ptrXMLDOMElement != NULL) 
			{
				_variant_t varCountry			= ptrXMLDOMElement->getAttribute(g_pszCountry);
				_variant_t varLanguage			= ptrXMLDOMElement->getAttribute(g_pszLanguage);
				_variant_t varEuro				= ptrXMLDOMElement->getAttribute(g_pszEuro);
				_variant_t varMSCodePage		= ptrXMLDOMElement->getAttribute(g_pszMSCodePage);
				_variant_t varIBMCCSID			= ptrXMLDOMElement->getAttribute(g_pszIBMCCSID);
				_variant_t varDefault			= ptrXMLDOMElement->getAttribute(g_pszDefault);
				_variant_t varOp				= ptrXMLDOMElement->getAttribute(g_pszOperation);
				_variant_t varTranslationTable	= ptrXMLDOMElement->text;

				_bstr_t bstrCountry				= L"";
				_bstr_t bstrLanguage			= L"";
				bool bEuro						= false;
				int nMSCodePage					= 0;
				int nIBMCCSID					= 0;
				bool bDefault					= false;
				_bstr_t bstrOp					= g_pszSet;	// Default operation is to set.
				_bstr_t bstrTranslationTable	= L"";

				if (varCountry.vt != VT_NULL && wcslen(varCountry.bstrVal) > 0)
				{
					bstrCountry = varCountry.bstrVal;
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing CODEPAGE/@%s: %s"), g_pszCountry, szCodePagesPath);
				}

				if (varLanguage.vt != VT_NULL && wcslen(varLanguage.bstrVal) > 0)
				{
					bstrLanguage = varLanguage.bstrVal;
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing CODEPAGE/@%s: %s"), g_pszLanguage, szCodePagesPath);
				}

				bEuro = varEuro.vt != VT_NULL && IsTrueString(varEuro.bstrVal);

				if (varMSCodePage.vt != VT_NULL && _wtoi(varMSCodePage.bstrVal) > 0)
				{
					nMSCodePage = _wtoi(varMSCodePage.bstrVal);
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing CODEPAGE/@%s: %s"), g_pszMSCodePage, szCodePagesPath);
				}

				if (varIBMCCSID.vt != VT_NULL)
				{
					nIBMCCSID = _wtoi(varIBMCCSID.bstrVal);
				}

				bDefault = varDefault.vt != VT_NULL && IsTrueString(varDefault.bstrVal);

				if (bDefault)
				{
					m_bFoundDefault = true;
				}

				if (varTranslationTable.vt != VT_NULL && wcslen(varTranslationTable.bstrVal) > 0)
				{
					bstrTranslationTable = varTranslationTable.bstrVal;
				}

				_bstr_t bstrLocaleKey = L"";
				if (bDefault)
				{
					bstrLocaleKey = GetLocaleKey(g_pszDefault, L"");
				}
				else
				{
					bstrLocaleKey = GetLocaleKey(bstrCountry, bstrLanguage, bEuro);
				}

				if (varOp.vt != VT_NULL && wcslen(varOp.bstrVal) > 0)
				{
					bstrOp = varOp.bstrVal;
				}

				CCodePage* pCodePage = NULL;
				MapLocaleToCodePageType::iterator it = m_MapLocaleToCodePage.find(bstrLocaleKey);
				if (it != m_MapLocaleToCodePage.end())
				{
					pCodePage = (*it).second;
				}
				
				EOp Op = opInsert;
				if (wcsicmp(bstrOp, g_pszSet) == 0)
				{
					if (pCodePage == NULL)
					{
						Op = opInsert;
					}
					else
					{
						Op = opUpdate;
					}
				}
				else if (wcsicmp(bstrOp, g_pszInsert) == 0)
				{
					if (pCodePage == NULL)
					{
						Op = opInsert;
					}
					else
					{
						throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - inserting an existing CODEPAGE %s: %s"), (LPCWSTR)bstrLocaleKey, szCodePagesPath);
					}
				}
				else if (wcsicmp(bstrOp, g_pszUpdate) == 0)
				{
					if (pCodePage != NULL)
					{
						Op = opUpdate;
					}
					else
					{
						throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - updating an non-existant CODEPAGE %s: %s"), (LPCWSTR)bstrLocaleKey, szCodePagesPath);
					}
				}
				else if (wcsicmp(bstrOp, g_pszDelete) == 0)
				{
					if (pCodePage != NULL)
					{
						Op = opDelete;
					}
					else
					{
						throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - deleting an non-existant CODEPAGE %s: %s"), (LPCWSTR)bstrLocaleKey, szCodePagesPath);
					}
				}

				if (Op == opInsert && pCodePage == NULL)
				{
					m_MapLocaleToCodePage.insert(MapLocaleToCodePageType::value_type(
						bstrLocaleKey,
						new CCodePage(bstrCountry, bstrLanguage, bEuro, nMSCodePage, nIBMCCSID, bDefault, bstrTranslationTable)));
				}
				else if (Op == opUpdate && pCodePage != NULL)
				{
					pCodePage->SetCountry(bstrCountry);
					pCodePage->SetLanguage(bstrLanguage);
					pCodePage->SetEuro(bEuro);
					pCodePage->SetMSCodePage(nMSCodePage);
					pCodePage->SetIBMCCSID(nIBMCCSID);
					pCodePage->SetDefault(bDefault);
					pCodePage->SetTranslationTable(bstrTranslationTable);
				}
				else if (Op == opDelete && pCodePage != NULL)
				{
					delete pCodePage;
					m_MapLocaleToCodePage.erase(it);
				}
			}

			lIndex++;
		}

		if (!m_bFoundDefault && lLength > 0)
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: no default code page: %s"), szCodePagesPath);
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

	return lIndex;
}

void CCodePages::Save(LPCWSTR szCodePagesPath)
{
	try
	{
		if (szCodePagesPath != NULL)
		{
			m_bstrCodePagesPath = szCodePagesPath;
		}

		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));

		MSXML::IXMLDOMElementPtr ptrXMLDOMElementParent = ptrXMLDOMDocument->createElement(L"CODEPAGES");
		ptrXMLDOMDocument->appendChild(ptrXMLDOMElementParent);

		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"ENABLEEURO");
		ptrXMLDOMElementParent->appendChild(ptrXMLDOMElement);
		ptrXMLDOMElement->text = m_bEnableEuro ? L"Y" : L"N";

		if (m_MapLocaleToCodePage.size() > 0)
		{
			MapLocaleToCodePageType::iterator it;
			for (it = m_MapLocaleToCodePage.begin(); it != m_MapLocaleToCodePage.end(); it++)
			{
				_bstr_t bstrLocaleKey = (*it).first;
				CCodePage* pCodePage = (*it).second;

				_ASSERTE(pCodePage  != NULL);

				if (pCodePage != NULL)
				{
					MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"CODEPAGE");
					ptrXMLDOMElementParent->appendChild(ptrXMLDOMElement);
					pCodePage->Save(ptrXMLDOMElement);
				}
			}
		}

		ptrXMLDOMDocument->save(m_bstrCodePagesPath);
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
}

_bstr_t CCodePages::GetLocaleKey(LPCWSTR szCountry, LPCWSTR szLanguage)
{
	return GetLocaleKey(szCountry, szLanguage, m_bEnableEuro);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnvOptimus::GetLocaleKey 
//	 
//	Description:  
//		Creates a locale name from the country and language code, and the Euro  
//		flag. E.g., for country "UK" and language "EN" with Euro support, the  
//		locale key is "UKENEURO".   
//
//		Optimus returns information about the supported and preferred locales. 
//		Each supported/preferred locale has a country and language code, as 
//		returned from Optimus.
//	 
//	Parameters:  
//		LPCWSTR szCountry:    
//			The country code for this locale.   
//		LPCWSTR szLanguage:    
//			The language code for this locale.   
//		bool bEuro:    
//			If true then the locale supports the Euro.   
//	 
//	Return:  
//		_bstr_t:    
//			The locale key. 
///////////////////////////////////////////////////////////////////////////////
_bstr_t CCodePages::GetLocaleKey(LPCWSTR szCountry, LPCWSTR szLanguage, bool bEuro)
{
	_bstr_t bstrLocale(L"");

	try
	{
		bstrLocale = szCountry;
		bstrLocale += szLanguage;

		if (bEuro)
		{
			bstrLocale += g_pszEuro;
		}

		_wcsupr(bstrLocale);
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return bstrLocale;
}

CCodePage* CCodePages::GetCodePage(LPCWSTR szCountry, LPCWSTR szLanguage)
{
	CCodePage* pCodePage = NULL;

	try
	{
		if (wcslen(szCountry) == 0 || wcslen(szLanguage) == 0)
		{
			szCountry = g_pszDefault;
		}

		MapLocaleToCodePageType::iterator it = m_MapLocaleToCodePage.find(GetLocaleKey(szCountry, szLanguage));

		if (it == m_MapLocaleToCodePage.end())
		{
			it = m_MapLocaleToCodePage.find(GetLocaleKey(g_pszDefault, L""));
			if (it == m_MapLocaleToCodePage.end())
			{
				throw CException(E_INVALIDCODEPAGE, __FILE__, __LINE__, _T("Invalid code page: no default code page"));
			}
			else
			{
				pCodePage = (*it).second;
			}
		}
		else
		{
			pCodePage = (*it).second;
		}

		if (pCodePage == NULL)
		{
			throw CException(E_INVALIDCODEPAGE, __FILE__, __LINE__, _T("Invalid code page"));
		}
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return pCodePage;
}
