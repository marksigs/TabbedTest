///////////////////////////////////////////////////////////////////////////////
//	FILE:			CodePage.cpp
//	DESCRIPTION:	Encapsulates single code page.
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		28/09/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CodePage.h"
#include "Exception.h"
#include "ODIConverter.h"

static LPCWSTR g_pszCountry								= L"COUNTRY";
static LPCWSTR g_pszLanguage							= L"LANGUAGE";
static LPCWSTR g_pszEuro								= L"EURO";
static LPCWSTR g_pszDefault								= L"DEFAULT";
static LPCWSTR g_pszMSCodePage							= L"MSCODEPAGE";
static LPCWSTR g_pszIBMCCSID							= L"IBMCCSID";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCodePage::CCodePage(
	LPCWSTR szCountry,
	LPCWSTR szLanguage,
	bool bEuro, 
	int nMSCodePage, 
	int nIBMCCSID, 
	bool bDefault,
	LPCWSTR szTranslationTable) :
		m_bstrCountry(szCountry),
		m_bstrLanguage(szLanguage),
		m_bEuro(bEuro),
		m_nMSCodePage(nMSCodePage),
		m_nIBMCCSID(nIBMCCSID),
		m_bDefault(bDefault),
		m_pTranslationTableSend(NULL),
		m_pTranslationTableRecv(NULL)
{
	SetTranslationTable(szTranslationTable);
}

CCodePage::~CCodePage()
{
	InitTranslationTables();
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::InitTranslationTables 
//	 
//	Description:  
//		Initialises translation tables.   
//	 
//	Parameters:  
//		None 
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
void CCodePage::InitTranslationTables()
{
	// Clean up.
	delete []m_pTranslationTableSend;
	m_pTranslationTableSend = NULL;
	delete []m_pTranslationTableRecv;
	m_pTranslationTableRecv = NULL;
	m_nTranslationTableSize = 0;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::SetTranslationTable
//
//	Description:
//		Creates send and receive translation tables from a comma/space separated 
//		string of hex numbers, e.g., "00 01 02 03 37 2D 2E 2F 16,..."
//
//		The translation tables are used to convert send and receive buffers when
//		the required Windows code page is not installed on the run time machine.
//		This particularly may be the case when using EBCDIC code pages.
//
//		Each item in the string becomes the value of the nth element in the
//		send translation table. The receive translation table is the inverse of the
//		send translation table.
//
//		The nth value in the translation table is the code to which the character 
//		with code n will be coverted.
//
//		For example, the ASCII code for the character 'A' is 65. The EBCDIC code 
//		for 'A' is 193. An ASCII to EBCDIC translation table would have the value
//		C1 (193 in hex) as its 65th element.  
//
//	Parameters:
//	  const LPCWSTR szTranslationTable:
//			Comma or space separated translation table, e.g, 
//			"00 01 02 03 37 2D 2E 2F 16,...".
//
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////

static const wchar_t* g_pszDelimiters = L", \r\n\t";

void CCodePage::SetTranslationTable(LPCWSTR szTranslationTable)
{
	try
	{
		m_bstrTranslationTable = szTranslationTable;

		InitTranslationTables();

		m_nTranslationTableSize = CalculateTranslationTableSize(szTranslationTable);

		if (m_nTranslationTableSize > 0)
		{
			m_pTranslationTableSend = new unsigned short[m_nTranslationTableSize];
			m_pTranslationTableRecv = new unsigned short[m_nTranslationTableSize];

			if (m_pTranslationTableSend != NULL && m_pTranslationTableRecv!= NULL)
			{
				int nTokens = 0;
				LPWSTR pszTransTable = _wcsdup(szTranslationTable);
				if (pszTransTable != NULL)
				{
					wchar_t* token = wcstok(pszTransTable, g_pszDelimiters);
					while (token != NULL && nTokens < m_nTranslationTableSize)
					{
						m_pTranslationTableSend[nTokens++] = hwtoi(token);
						token = wcstok(NULL, g_pszDelimiters);
					}
					free(pszTransTable);
				}

				// Receive translation table is inverse of send translation table.
				for (int nIndex = 0; nIndex < m_nTranslationTableSize; nIndex++)
				{
					m_pTranslationTableRecv[m_pTranslationTableSend[nIndex]] = nIndex;
				}
			}
			else
			{
				InitTranslationTables();
			}
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
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::CalculateTranslationTableSize 
//	 
//	Description:  
//		Calculates the size of the translation table as the number of comma  
//		separated items in the input string + 1.   
//	 
//	Parameters:  
//		LPCWSTR szTranslationTable:    
//			Translation table.   
//	 
//	Return:  
//		int:    
//			Number of items in translation table. 
///////////////////////////////////////////////////////////////////////////////
int CCodePage::CalculateTranslationTableSize(LPCWSTR szTranslationTable) const
{
	int nTokens = 0;

	try
	{
		LPWSTR pszTransTable = _wcsdup(szTranslationTable);
		if (pszTransTable != NULL)
		{
			wchar_t* token = wcstok(pszTransTable, g_pszDelimiters);
			while (token != NULL)
			{
				nTokens++;
				token = wcstok(NULL, g_pszDelimiters);
			}
			free(pszTransTable);
		}
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nTokens;
}

int CCodePage::SendW2MBC(LPCWSTR lpWide, LPSTR lpMBC) 
{
	return SendW2MBC(lpWide, -1, lpMBC, wcslen(lpWide) + 1);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::SendW2MBC 
//	 
//	Description:  
//		Converts a send buffer from Unicode to a multi byte string, using a  
//		Windows code page or user supplied translation table if the code page  
//		is not available.   
//	 
//	Parameters:  
//		LPCWSTR lpWide:    
//			The Unicode string to be converted.   
//		int nMaxSizeWide:    
//			The maximum size of the Unicode string.   
//		LPSTR lpMBC:    
//			Buffer to receive the resulting multi byte string.   
//		int nMaxSizeMBC:    
//			Maximum size of the multi byte string buffer.   
//	 
//	Return:  
//		int:    
//			The number of bytes converted. 
///////////////////////////////////////////////////////////////////////////////
int CCodePage::SendW2MBC(LPCWSTR lpWide, int nMaxSizeWide, LPSTR lpMBC, int nMaxSizeMBC) 
{
	int nResult = 0;

	try
	{
		_ASSERTE(lpMBC != NULL);
		_ASSERTE(lpWide != NULL);
		
		lpMBC[0] = '\0';
		if (::IsValidCodePage(m_nMSCodePage))
		{
			// Code page is installed on this machine.
			nResult = ::WideCharToMultiByte(m_nMSCodePage, 0, lpWide, nMaxSizeWide, lpMBC, nMaxSizeMBC, NULL, NULL);
		}
		else if (m_pTranslationTableSend != NULL)
		{
			// Code page is not available so use user supplied translation table.
			nResult = 
				Translate(
					reinterpret_cast<const BYTE*>(lpWide), 
					nMaxSizeWide,
					reinterpret_cast<BYTE*>(lpMBC), 
					nMaxSizeMBC,
					m_pTranslationTableSend, 
					m_nTranslationTableSize,
					TRUE,
					FALSE);
		}
		else
		{
			throw CException(E_INVALIDCODEPAGE, __FILE__, __LINE__, _T("Invalid code page: %d"), m_nMSCodePage);
		}

		_ASSERTE(nResult);

	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nResult;
}

int CCodePage::RecvMBC2W(LPCSTR lpMBC, LPWSTR lpWide) 
{
	return RecvMBC2W(lpMBC, -1, lpWide, (strlen(lpMBC) * 2) + 2);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::RecvMBC2W 
//	 
//	Description:  
//		Converts a receive buffer from a multi byte string to Unicode, using a  
//		Windows code page or user supplied translation table if the code page  
//		is not available.   
//	 
//	Parameters:  
//		LPCSTR lpMBC:    
//			The multi byte string to be converted.   
//		int nMaxSizeMBC:    
//			Maximum size of the mult byte string.   
//		LPWSTR lpWide:    
//			Buffer to receive the Unicode string.   
//		int nMaxSizeWide:    
//			Maximum size of the Unicode buffer.   
//	 
//	Return:  
//		int:    
//			Number of characters converted. 
///////////////////////////////////////////////////////////////////////////////
int CCodePage::RecvMBC2W(LPCSTR lpMBC, int nMaxSizeMBC, LPWSTR lpWide, int nMaxSizeWide) 
{
	int nResult = 0;

	try
	{
		_ASSERTE(lpMBC != NULL);
		_ASSERTE(lpWide != NULL);

		lpWide[0] = L'\0';
		if (::IsValidCodePage(m_nMSCodePage))
		{
			// Code page is installed on this machine.
			nResult = ::MultiByteToWideChar(m_nMSCodePage, 0, lpMBC, nMaxSizeMBC, lpWide, nMaxSizeWide);
		}
		else if (m_pTranslationTableRecv != NULL)
		{
			// Code page is not available so use user supplied translation table.
			nResult = 
				Translate(
					reinterpret_cast<const BYTE*>(lpMBC), 
					nMaxSizeMBC,
					reinterpret_cast<BYTE*>(lpWide), 
					nMaxSizeWide,
					m_pTranslationTableRecv, 
					m_nTranslationTableSize,
					FALSE,
					TRUE);
		}
		else
		{
			throw CException(E_INVALIDCODEPAGE, __FILE__, __LINE__, _T("Invalid code page: %d"), m_nMSCodePage);
		}

		_ASSERTE(nResult != 0);
	}
	catch(CException&)
	{
		throw;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nResult;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::Translate 
//	 
//	Description:  
//		Converts a byte stream to another byte stream using a translation  
//		table.   
//	 
//	Parameters:  
//		const BYTE* pbySrc:    
//			Byte steam to convert.
//		const int nSrcLen:    
//			Maximum size of the byte stream to convert.
//		BYTE* pbyTgt:    
//			Buffer to receive converted byte stream.
//		const int nTgtLen:    
//			Maximum size of the receive buffer.
//		const unsigned short* TransTable:    
//			Translation table for conversion.
//		const int nMaxSizeTransTable:    
//			Size of the translation table.
//		const BOOL bSrcW:    
//			If TRUE, then source byte stream is wide.
//		const BOOL bTgtW:    
//			If TRUE, then the target byte stream is wide.
//
//	Return:  
//		int:  
//			Number of characters converted.
///////////////////////////////////////////////////////////////////////////////
int CCodePage::Translate(
	const BYTE* pbySrc, 
	const int nSrcLen,
	BYTE* pbyTgt, 
	const int nTgtLen,
	const unsigned short* TransTable, 
	const int nMaxSizeTransTable,
	const BOOL bSrcW,
	const BOOL bTgtW)
{
	int nResult = 0;

	try
	{
		int nSrcByteLen = 0;
		int nSrcByteInc = 0;
		int nTgtByteLen = 0;
		int nTgtByteInc = 0;
		int nActualSrcLen = nSrcLen;

		if (bSrcW)
		{
			if (nActualSrcLen == -1)
			{
				nActualSrcLen = wcslen(reinterpret_cast<LPCWSTR>(pbySrc)) + 1;
			}
			nSrcByteLen = nActualSrcLen * 2;
			nSrcByteInc = 2;
			if (bTgtW)
			{
				nTgtByteLen = nSrcByteLen;
			}
			else
			{
				nTgtByteLen = nSrcByteLen / 2;
			}
		}
		else
		{
			if (nActualSrcLen == -1)
			{
				nActualSrcLen = strlen(reinterpret_cast<LPCSTR>(pbySrc)) + 1;
			}
			nSrcByteLen = nActualSrcLen;
			nSrcByteInc = 1;
			if (bTgtW)
			{
				nTgtByteLen = nSrcByteLen * 2;
			}
			else
			{
				nTgtByteLen = nSrcByteLen;
			}
		}

		if (bTgtW)
		{
			nTgtByteInc = 2;
		}
		else
		{
			nTgtByteInc = 1;
		}

		_ASSERTE(pbySrc != NULL);
		_ASSERTE(pbyTgt != NULL);

		int nTgtIndex = 0;
		if (pbySrc != NULL && pbyTgt != NULL)
		{
			::ZeroMemory(pbyTgt, nTgtByteLen + nTgtByteInc);
			for (
					int nSrcIndex = 0; 
					nSrcIndex < nSrcByteLen && nTgtIndex < nTgtByteLen; 
					nSrcIndex += nSrcByteInc,
					nTgtIndex += nTgtByteInc)
			{
				// Note we ignore the second byte for wide characters - assumes ANSII character set.
				BYTE ch = *(pbySrc + nSrcIndex);

				_ASSERTE(ch < nMaxSizeTransTable);

				if (ch < nMaxSizeTransTable)
				{
					*(pbyTgt + nTgtIndex) = static_cast<BYTE>(TransTable[ch]);
				}
			}
		}

		nResult = nTgtIndex > 0 ? nTgtIndex / nTgtByteInc : 0;
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nResult;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CCodePage::GetCodePageTable 
//	 
//	Description:  
//		This function is useful for creating meta data translation tables on  
//		one machine (which has the code page) for use on another machine  
//		(which does not have the code page).   
//	 
//	Parameters:  
//		int nCodePage:    
//			Windows code page.   
//		unsigned short* CodePageTable:    
//			Buffer to receive translated table.   
//		int nMaxSizeCodePageTable:    
//			Maximum size of translation table buffer.   
//	 
//	Return:  
//		int:    
//			Number of characters converted in translation table. 
///////////////////////////////////////////////////////////////////////////////
int CCodePage::GetCodePageTable(int nCodePage, unsigned short* CodePageTable, int nMaxSizeCodePageTable)
{
	int nResult = 0;

	try
	{
		// Code page must exist!
		_ASSERTE(::IsValidCodePage(nCodePage));

		if (::IsValidCodePage(nCodePage))
		{
			::ZeroMemory(CodePageTable, nMaxSizeCodePageTable);
			unsigned short* pCodePageTable = CodePageTable;
			for (int nChar = 0; nChar < nMaxSizeCodePageTable; nChar++)
			{
				wchar_t lpWide[2];
				lpWide[0] = nChar;
				lpWide[1] = L'\0';
				BOOL bUsedDefaultChar = FALSE;

				nResult += ::WideCharToMultiByte(nCodePage, 0, lpWide, wcslen(lpWide) + 1, reinterpret_cast<LPSTR>(pCodePageTable), sizeof(unsigned short), "\0", &bUsedDefaultChar);

				pCodePageTable++;
			}
		}
	}
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return nResult;
}

void CCodePage::Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode)
{
	_ASSERTE(ptrXMLDOMNode != NULL);

	if (ptrXMLDOMNode != NULL)
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMNode;
		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument = ptrXMLDOMNode->ownerDocument;

		ptrXMLDOMElement->setAttribute(g_pszCountry, m_bstrCountry);
		ptrXMLDOMElement->setAttribute(g_pszLanguage, m_bstrLanguage);
		
		wchar_t szMSCodePage[16] = L"";
		swprintf(szMSCodePage, L"%d", m_nMSCodePage);
		ptrXMLDOMElement->setAttribute(g_pszMSCodePage, szMSCodePage);

		wchar_t szIBMCCSID[16] = L"";
		swprintf(szIBMCCSID, L"%d", m_nIBMCCSID);
		ptrXMLDOMElement->setAttribute(g_pszIBMCCSID, szIBMCCSID);

		ptrXMLDOMElement->setAttribute(g_pszEuro, m_bEuro ? L"Y" : L"N");
		ptrXMLDOMElement->setAttribute(g_pszDefault, m_bDefault ? L"Y" : L"N");
		ptrXMLDOMElement->text = m_bstrTranslationTable;
	}
}