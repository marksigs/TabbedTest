///////////////////////////////////////////////////////////////////////////////
//	FILE:			LookUpTable.cpp
//	DESCRIPTION:	
//		Encapsulates send and receive look up tables for a property.
//		The look up tables enable values for the property to be automatically
//		converted from one value to another when the property is sent to the
//		remote machine, or received from the remote machine.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		30/10/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "ODIConverter.h"
#include "LookUpTable.h"


static LPCWSTR g_pszItems								= L".//ITEM";
static LPCWSTR g_pszName								= L"NAME";
static LPCWSTR g_pszSendIn								= L"SENDIN";
static LPCWSTR g_pszSendOut								= L"SENDOUT";
static LPCWSTR g_pszRecvIn								= L"RECVIN";
static LPCWSTR g_pszRecvOut								= L"RECVOUT";
static LPCWSTR g_pszDefault								= L"DEFAULT";
static LPCWSTR g_pszSendSrc								= L"SENDSRC";
static LPCWSTR g_pszRecvSrc								= L"RECVSRC";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CLookUpTable::CLookUpTable(LPCWSTR szName, const MSXML::IXMLDOMElementPtr ptrXMLDOMElement) :
	m_bstrName(szName),
	m_bstrSendSrc(L""),
	m_bstrRecvSrc(L"")
{
	Init(ptrXMLDOMElement);
}


CLookUpTable::~CLookUpTable()
{
}


///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTable::Init
//	
//	Description:
//		Initialise look up table.
//	
//	Parameters:
//		const IXMLDOMElementPtr ptrXMLDOMElement:
//			XML DOM element that describes the look up table. 
//			See LookUpTables.cpp for a description.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void CLookUpTable::Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode)
{
	_ASSERTE(ptrXMLDOMNode != NULL);

	if (ptrXMLDOMNode != NULL)
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMNode;
		_variant_t varSendSrc	= ptrXMLDOMElement->getAttribute(g_pszSendSrc);
		_variant_t varRecvSrc	= ptrXMLDOMElement->getAttribute(g_pszRecvSrc);
		
		if (varSendSrc.vt != VT_NULL && wcslen(varSendSrc.bstrVal) > 0)
		{
			m_bstrSendSrc = varSendSrc.bstrVal;
		}
		if (varRecvSrc.vt != VT_NULL && wcslen(varRecvSrc.bstrVal) > 0)
		{
			m_bstrRecvSrc = varRecvSrc.bstrVal;
		}

		MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMElement->selectNodes(g_pszItems);

		long lLength = ptrXMLDOMNodeList->length;
		long lIndex = 0;

		// Assume look up tables includes all items.
		m_MapSendInToSendOut.clear();
		m_MapRecvInToRecvOut.clear();
		while (lIndex < lLength)
		{
			MSXML::IXMLDOMElementPtr ptrXMLDOMItem = ptrXMLDOMNodeList->item[lIndex];

			if (ptrXMLDOMItem != NULL) 
			{
				_variant_t varName		= ptrXMLDOMItem->getAttribute(g_pszName);
				_variant_t varSendIn	= ptrXMLDOMItem->getAttribute(g_pszSendIn);
				_variant_t varSendOut	= ptrXMLDOMItem->getAttribute(g_pszSendOut);
				_variant_t varRecvIn	= ptrXMLDOMItem->getAttribute(g_pszRecvIn);
				_variant_t varRecvOut	= ptrXMLDOMItem->getAttribute(g_pszRecvOut);

				_bstr_t bstrName		= L"";
				_bstr_t bstrSendIn		= L"";
				_bstr_t bstrSendOut		= L"";
				_bstr_t bstrRecvIn		= L"";
				_bstr_t bstrRecvOut		= L"";

				if (varName.vt != VT_NULL && wcslen(varName.bstrVal) > 0)
				{
					bstrName = varName.bstrVal;
				}

				if (varSendIn.vt != VT_NULL && wcslen(varSendIn.bstrVal) > 0)
				{
					// SendIn value specified.
					bstrSendIn = varSendIn.bstrVal;

					if (varSendOut.vt != VT_NULL && wcslen(varSendOut.bstrVal) > 0)
					{
						bstrSendOut = varSendOut.bstrVal;
					}

					MapInToOutType::iterator it = m_MapSendInToSendOut.find(bstrSendIn);
					if (it != m_MapSendInToSendOut.end())
					{
						(*it).second.first	= bstrSendOut;
						(*it).second.second	= bstrName;
					}
					else
					{
						m_MapSendInToSendOut.insert(MapInToOutType::value_type(bstrSendIn, PairMapInToOutType(bstrSendOut, bstrName)));
					}
				}
				if (varRecvIn.vt != VT_NULL && wcslen(varRecvIn.bstrVal) > 0)
				{
					// RecvIn value specified.
					bstrRecvIn = varRecvIn.bstrVal;

					if (varRecvOut.vt != VT_NULL && wcslen(varRecvOut.bstrVal) > 0)
					{
						bstrRecvOut = varRecvOut.bstrVal;
					}

					MapInToOutType::iterator it = m_MapRecvInToRecvOut.find(bstrRecvIn);
					if (it != m_MapRecvInToRecvOut.end())
					{
						(*it).second.first	= bstrRecvOut;
						(*it).second.second	= bstrName;
					}
					else
					{
						m_MapRecvInToRecvOut.insert(MapInToOutType::value_type(bstrRecvIn, PairMapInToOutType(bstrRecvOut, bstrName)));
					}
				}
			}

			lIndex++;
		}		
	}
}

void CLookUpTable::Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode)
{
	_ASSERTE(ptrXMLDOMNode != NULL);

	if (ptrXMLDOMNode != NULL)
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMNode;
		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument = ptrXMLDOMNode->ownerDocument;

		ptrXMLDOMElement->setAttribute(g_pszSendSrc, m_bstrSendSrc);
		ptrXMLDOMElement->setAttribute(g_pszRecvSrc, m_bstrRecvSrc);

		typedef std::map<_bstr_t, bool> NamesType;
		NamesType Names;

		MapInToOutType::iterator itSend;
		for (itSend = m_MapSendInToSendOut.begin(); itSend != m_MapSendInToSendOut.end(); itSend++)
		{
			_bstr_t bstrSendIn	= (*itSend).first;
			_bstr_t bstrSendOut = (*itSend).second.first;
			_bstr_t bstrName	= (*itSend).second.second;
			_bstr_t bstrRecvIn	= L"";
			_bstr_t bstrRecvOut = L"";

			// Find matching receive map.
			MapInToOutType::iterator itRecv;
			for (itRecv = m_MapRecvInToRecvOut.begin(); itRecv != m_MapRecvInToRecvOut.end(); itRecv++)
			{
				if (wcsicmp(bstrName, (*itRecv).second.second) == 0)
				{
					bstrRecvIn	= (*itRecv).first;
					bstrRecvOut = (*itRecv).second.first;
					break;
				}
			}
			Names.insert(NamesType::value_type(bstrName, true));

			ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"ITEM");
			ptrXMLDOMNode->appendChild(ptrXMLDOMElement);
			ptrXMLDOMElement->setAttribute(g_pszName, bstrName);
			ptrXMLDOMElement->setAttribute(g_pszSendIn, bstrSendIn);
			ptrXMLDOMElement->setAttribute(g_pszSendOut, bstrSendOut);
			ptrXMLDOMElement->setAttribute(g_pszRecvIn, bstrRecvIn);
			ptrXMLDOMElement->setAttribute(g_pszRecvOut, bstrRecvOut);
		}

		MapInToOutType::iterator itRecv;
		for (itRecv = m_MapRecvInToRecvOut.begin(); itRecv != m_MapRecvInToRecvOut.end(); itRecv++)
		{
			_bstr_t bstrRecvIn	= (*itRecv).first;
			_bstr_t bstrRecvOut = (*itRecv).second.first;
			_bstr_t bstrName	= (*itRecv).second.second;
			_bstr_t bstrSendIn	= L"";
			_bstr_t bstrSendOut = L"";

			NamesType::iterator it = Names.find(bstrName);
			if (it == Names.end())
			{
				// Not already output as part of send map.
				// Find matching send map.
				MapInToOutType::iterator itSend;
				for (itSend = m_MapSendInToSendOut.begin(); itSend != m_MapSendInToSendOut.end(); itSend++)
				{
					if (wcsicmp(bstrName, (*itSend).second.second) == 0)
					{
						bstrSendIn	= (*itSend).first;
						bstrSendOut = (*itSend).second.first;
						break;
					}
				}
				ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"ITEM");
				ptrXMLDOMNode->appendChild(ptrXMLDOMElement);
				ptrXMLDOMElement->setAttribute(g_pszName, bstrName);
				ptrXMLDOMElement->setAttribute(g_pszSendIn, bstrSendIn);
				ptrXMLDOMElement->setAttribute(g_pszSendOut, bstrSendOut);
				ptrXMLDOMElement->setAttribute(g_pszRecvIn, bstrRecvIn);
				ptrXMLDOMElement->setAttribute(g_pszRecvOut, bstrRecvOut);
			}
		}

	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTable::GetValue
//	
//	Description:
//		Gets a conversion value from the look up table.
//	
//	Parameters:
//		LPCWSTR szKey:
//			The value to be converted.
//		const EDirection Direction:
//			The direction of the conversion; one of CLookUpTable::dirSend or 
//			CLookUpTable::dirRecv. This is required as each look up table has
//			one set of conversion values for when sending data, and one set
//			for when receiving data.
//	
//	Return:
//		_bstr_t: 	
//			The converted value.
///////////////////////////////////////////////////////////////////////////////
_bstr_t CLookUpTable::GetValue(LPCWSTR szKey, const EDirection Direction)
{
	MapInToOutType& MapInToOut = GetMapInToOut(Direction);
	MapInToOutType::iterator it = MapInToOut.find(szKey);
	if (it == MapInToOut.end())
	{
		// Key not found, so return empty string;
		return L"";
	}
	else
	{
		// Found key.
		return (*it).second.first;
	}	
}


void CLookUpTable::WriteCSV(FILE* fp, LPCWSTR szProperty)
{
	typedef std::map<_bstr_t, bool> NamesType;
	NamesType Names;

	MapInToOutType::iterator itSend;
	for (itSend = m_MapSendInToSendOut.begin(); itSend != m_MapSendInToSendOut.end(); itSend++)
	{
		_bstr_t bstrSendIn	= (*itSend).first;
		_bstr_t bstrSendOut = (*itSend).second.first;
		_bstr_t bstrName	= (*itSend).second.second;
		_bstr_t bstrRecvIn	= L"";
		_bstr_t bstrRecvOut = L"";

		// Find matching receive map.
		MapInToOutType::iterator itRecv;
		for (itRecv = m_MapRecvInToRecvOut.begin(); itRecv != m_MapRecvInToRecvOut.end(); itRecv++)
		{
			if (wcsicmp(bstrName, (*itRecv).second.second) == 0)
			{
				bstrRecvIn	= (*itRecv).first;
				bstrRecvOut = (*itRecv).second.first;
				break;
			}
		}
		Names.insert(NamesType::value_type(bstrName, true));
		fwprintf(
				fp, 
				L"%s,%s,%s,%s,%s,%s,%s,%s,%s\n", 
				szProperty,
				static_cast<LPCWSTR>(m_bstrName), 
				static_cast<LPCWSTR>(m_bstrSendSrc), 
				static_cast<LPCWSTR>(m_bstrRecvSrc), 
				static_cast<LPCWSTR>(bstrName), 
				static_cast<LPCWSTR>(bstrSendIn), 
				static_cast<LPCWSTR>(bstrSendOut), 
				static_cast<LPCWSTR>(bstrRecvIn), 
				static_cast<LPCWSTR>(bstrRecvOut));
	}

	MapInToOutType::iterator itRecv;
	for (itRecv = m_MapRecvInToRecvOut.begin(); itRecv != m_MapRecvInToRecvOut.end(); itRecv++)
	{
		_bstr_t bstrRecvIn	= (*itRecv).first;
		_bstr_t bstrRecvOut = (*itRecv).second.first;
		_bstr_t bstrName	= (*itRecv).second.second;
		_bstr_t bstrSendIn	= L"";
		_bstr_t bstrSendOut = L"";

		NamesType::iterator it = Names.find(bstrName);
		if (it == Names.end())
		{
			// Not already output as part of send map.
			// Find matching send map.
			MapInToOutType::iterator itSend;
			for (itSend = m_MapSendInToSendOut.begin(); itSend != m_MapSendInToSendOut.end(); itSend++)
			{
				if (wcsicmp(bstrName, (*itSend).second.second) == 0)
				{
					bstrSendIn	= (*itSend).first;
					bstrSendOut = (*itSend).second.first;
					break;
				}
			}
			fwprintf(
				fp, 
				L"%s,%s,%s,%s,%s,%s,%s,%s,%s\n", 
				szProperty,
				static_cast<LPCWSTR>(m_bstrName), 
				static_cast<LPCWSTR>(m_bstrSendSrc), 
				static_cast<LPCWSTR>(m_bstrRecvSrc), 
				static_cast<LPCWSTR>(bstrName), 
				static_cast<LPCWSTR>(bstrSendIn), 
				static_cast<LPCWSTR>(bstrSendOut), 
				static_cast<LPCWSTR>(bstrRecvIn), 
				static_cast<LPCWSTR>(bstrRecvOut));
		}
	}

	fwprintf(fp, L"\n");
}

