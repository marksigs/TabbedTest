///////////////////////////////////////////////////////////////////////////////
//	FILE:			LookUpTables.cpp
//	DESCRIPTION:	
//		Encapsulates a container for look up tables. The look up tables enable 
//		values for a property to be automatically converted from one value to 
//		another when the property is sent to the remote machine, or received 
//		from the remote machine.
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
#include "LookUpTables.h"


static LPCWSTR g_pszLookUpTableKey						= L"//LOOKUPTABLES/LOOKUPTABLE";
static LPCWSTR g_pszPropertyKey							= L"//LOOKUPTABLES/PROPERTY";
static LPCWSTR g_pszName								= L"NAME";
static LPCWSTR g_pszOperation							= L"OPERATION";
static LPCWSTR g_pszSet									= L"SET";
static LPCWSTR g_pszInsert								= L"INSERT";
static LPCWSTR g_pszUpdate								= L"UPDATE";
static LPCWSTR g_pszDelete								= L"DELETE";
static LPCWSTR g_pszSend								= L"SEND";
static LPCWSTR g_pszRecv								= L"RECV";
static LPCWSTR g_pszLookUpTable							= L"LOOKUPTABLE";

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CLookUpTables::CLookUpTables() :
	m_bstrLookUpTablesPath(L"")
{
}


CLookUpTables::~CLookUpTables()
{
	// Free all allocated objects.
	if (m_MapNameToLookUpTable.size() > 0)
	{
		MapNameToLookUpTableType::iterator it;
		for (it = m_MapNameToLookUpTable.begin(); it != m_MapNameToLookUpTable.end(); it++)
		{
			CLookUpTable* pLookUpTable = (*it).second;

			delete pLookUpTable;
		}
		m_MapNameToLookUpTable.clear();
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::Init
//	
//	Description:
//		Initialise look up tables.
//	
//	Parameters:
//		LPCWSTR szLookUpTablesPath:
//			Full path of look up tables XML file. The XML is in the form, e.g.,
//
//			<LOOKUPTABLES>
//
//				<PROPERTY NAME="CustomerDetailImpl.Gender"	LOOKUPTABLE="Gender"/>
//				<PROPERTY NAME="MortgagorImpl.Gender"		LOOKUPTABLE="Gender"/>
//				
//				<LOOKUPTABLE NAME="Gender">
//					<ITEM NAME="Male"			SENDIN="1"			SENDOUT="1"		RECVIN="1"			RECVOUT="1"/>
//					<ITEM NAME="Female"			SENDIN="2"			SENDOUT="0"		RECVIN="0"			RECVOUT="2"/>
//					<ITEM NAME="Undefined"		SENDIN=""			SENDOUT=""		RECVIN="9"			RECVOUT=""/>
//				</LOOKUPTABLE>
//
//			</LOOKUPTABLES>
//	
//			Each PROPERTY node defines the relationship between a property and
//			the look up table used to convert the data for that property.
//
//			Each LOOKUPTABLE nodes defines the how one value should be converted
//			to another.	Note that the conversion values when receiving data are not
//			necessarily the inverse of the conversion values when sending data.
//			Thus each look up table actually consists of two separate mappings,
//			one used for sending and one used for receiving.
//
//			There is a one to many relationship between a LOOKUPTABLE node and
//			PROPERTY nodes, i.e., the same look up table can be used to 
//			convert the data for many different types of property.
//
//			LOOKUPTABLES/PROPERTY:
//				Define a PROPERTY node for each property for which the
//				values should be automatically converted.
//			PROPERTY/@NAME:
//				The unique name of the property, e.g., for Optimus this is in the form
//				<object name>.<property name>. This name is the key used to 
//				set/get the look up table from the map m_MapPropertyToLookUpTable.
//			PROPERTY/@LOOKUPTABLE:
//				The NAME attribute of the look up table to use for converting values
//				of this property.
//
//			LOOKUPTABLES/LOOKUPTABLE:
//				Define a LOOKUPTABLE node for each unique type of conversion.
//			LOOKUPTABLE/@NAME:
//				Uniquely identifies the look up table, and is the key used to 
//				set/get the look up table from the map m_MapNameToLookUpTable.
//			LOOKUPTABLE/ITEM:
//				There is ITEM element for each possible value of the property.
//			LOOKUPTABLE/ITEM/@NAME:
//				The name of the item. This attribute is ignored by ODIConverter.
//				It is useful for documentary purposes.
//			LOOKUPTABLE/ITEM/@SENDIN:
//				The value of the property sent into ODIConverter, e.g., by Omiga.
//				Omit or set to the empty string if you do not want any 
//				conversion on sending.
//			LOOKUPTABLE/ITEM/@SENDOUT:
//				The value of the property sent out of ODIConverter, e.g., to 
//				Optimus. If @SENDIN is not specified, @SENDOUT is ignored.
//			LOOKUPTABLE/ITEM/@RECVIN:
//				The value of the property received into ODIConverter, e.g., 
//				from Optimus. Omit or set to the empty string if you do not want 
//				any conversion on receiving.
//			LOOKUPTABLE/ITEM/@RECVOUT:
//				The value of the property received from ODIConverter, e.g., by 
//				Omiga. If @RECVIN is not specified, @RECVOUT is ignored.
//
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
long CLookUpTables::Init(LPCWSTR szLookUpTablesPath)
{
	long lLookUpTables = 0;

	try
	{
		// initialise from XML file
		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));

		if (_taccess(szLookUpTablesPath, 00) != 0)
		{
			throw CException(E_FILEDOESNOTEXIST, __FILE__, __LINE__, _T("File does not exist: %s"), szLookUpTablesPath);
		}

		ptrXMLDOMDocument->async = false;
		_variant_t varLoaded = ptrXMLDOMDocument->load(szLookUpTablesPath);

		if (varLoaded.boolVal == false)
		{
			throw CException(E_INVALIDMETADATAFILE, __FILE__, __LINE__, _T("Invalid meta data file: %s"), szLookUpTablesPath);
		}

		m_bstrLookUpTablesPath = szLookUpTablesPath;

		lLookUpTables = InitNode(ptrXMLDOMDocument->documentElement, szLookUpTablesPath);
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

#ifdef _DEBUG
//	WriteCSV();
#endif

	return lLookUpTables;
}

long CLookUpTables::InitNode(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath)
{
	long lLookUpTables = 0;
	long lProperties = 0;

	try
	{
		if (szLookUpTableMapPath == NULL)
		{
			szLookUpTableMapPath = L"Unknown source";
		}
		lLookUpTables	= InitLookUpTables(ptrXMLDOMNode, szLookUpTableMapPath);
		lProperties		= InitProperties(ptrXMLDOMNode, szLookUpTableMapPath);
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

	return lLookUpTables;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::InitLookUpTables
//	
//	Description:
//		Initialises the look up table map, m_MapNameToLookUpTable. This maps
//		look up table names, LOOKUPTABLE/@NAME, to CLookUpTable objects. The
//		map owns the objects.
//	
//	Parameters:
//		const IXMLDOMDocumentPtr ptrXMLDOMDocument:
//			The look up table XML document.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
long CLookUpTables::InitLookUpTables(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath)
{
	MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMNode->selectNodes(g_pszLookUpTableKey);
	MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL;

	long lLength = ptrXMLDOMNodeList->length;
	long lIndex = 0;

	while (lIndex < lLength)
	{
		ptrXMLDOMElement = ptrXMLDOMNodeList->item[lIndex];

		if (ptrXMLDOMElement != NULL) 
		{
			_variant_t varName	= ptrXMLDOMElement->getAttribute(g_pszName);
			_variant_t varOp	= ptrXMLDOMElement->getAttribute(g_pszOperation);
			_bstr_t bstrName	= L"";
			_bstr_t bstrOp		= g_pszSet;	// Default operation is to set.

			if (varName.vt != VT_NULL && wcslen(varName.bstrVal) > 0)
			{
				bstrName = varName.bstrVal;
			}
			else
			{
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing LOOKUPTABLE/@%s: %s"), g_pszName, szLookUpTableMapPath);
			}

			if (varOp.vt != VT_NULL && wcslen(varOp.bstrVal) > 0)
			{
				bstrOp = varOp.bstrVal;
			}

			CLookUpTable* pLookUpTable = NULL;
			MapNameToLookUpTableType::iterator it = m_MapNameToLookUpTable.find(bstrName);
			if (it != m_MapNameToLookUpTable.end())
			{
				pLookUpTable = (*it).second;
			}

			EOp Op = opNull;
			if (wcsicmp(bstrOp, g_pszSet) == 0)
			{
				if (pLookUpTable == NULL)
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
				if (pLookUpTable == NULL)
				{
					Op = opInsert;
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - inserting an existing LOOKUPTABLE/@%s=%s: %s"), g_pszName, (LPCWSTR)bstrName, szLookUpTableMapPath);
				}
			}
			else if (wcsicmp(bstrOp, g_pszUpdate) == 0)
			{
				if (pLookUpTable != NULL)
				{
					Op = opUpdate;
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - updating an non-existant LOOKUPTABLE/@%s=%s: %s"), g_pszName, (LPCWSTR)bstrName, szLookUpTableMapPath);
				}
			}
			else if (wcsicmp(bstrOp, g_pszDelete) == 0)
			{
				if (pLookUpTable != NULL)
				{
					Op = opDelete;
				}
				else
				{
					throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - deleting an non-existant LOOKUPTABLE/@%s=%s: %s"), g_pszName, (LPCWSTR)bstrName, szLookUpTableMapPath);
				}
			}

			if (Op == opInsert && pLookUpTable == NULL)
			{
				m_MapNameToLookUpTable.insert(MapNameToLookUpTableType::value_type(
					bstrName,
					new CLookUpTable(bstrName, ptrXMLDOMElement)));
			}
			else if (Op == opUpdate && pLookUpTable != NULL)
			{
				pLookUpTable->Init(ptrXMLDOMElement);
			}
			else if (Op == opDelete && pLookUpTable != NULL)
			{
				// Must remove any properties that use this lookup table.
				DeletePropertiesForLookUpTable(pLookUpTable);
				delete pLookUpTable;
				m_MapNameToLookUpTable.erase(it);
			}
		}

		lIndex++;
	}

	return lIndex;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::DeletePropertiesForLookUpTable
//	
//	Description:
//		For a given lookup table, removes all properties from the property to 
//		lookup table map that point to the lookup table.
//	
//	Parameters:
//		CLookUpTable* pLookUpTable:
//			The lookup table for which properties should be removed.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void CLookUpTables::DeletePropertiesForLookUpTable(CLookUpTable* pLookUpTable)
{
	if (m_MapPropertyToLookUpTable.size() > 0)
	{
		MapNameToLookUpTableType::iterator it;
		for (it = m_MapPropertyToLookUpTable.begin(); it != m_MapPropertyToLookUpTable.end(); )
		{
			CLookUpTable* pThisLookUpTable = (*it).second;
			if (pThisLookUpTable == pLookUpTable)
			{
				it = m_MapPropertyToLookUpTable.erase(it);
			}
			else
			{
				it++;
			}
		}
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::InitProperties
//	
//	Description:
//		Initialises the look up table map, m_MapPropertyToLookUpTable. This maps
//		property names, PROPERTY/@NAME, to CLookUpTable objects. The map does not
//		own the objects. This function must be called AFTER InitLookUpTables, 
//		which creates the look up table objects that will be used in 
//		m_MapPropertyToLookUpTable.
//	
//	Parameters:
//		const IXMLDOMDocumentPtr ptrXMLDOMDocument:
//			The look up table XML document.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
long CLookUpTables::InitProperties(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath)
{
	MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMNode->selectNodes(g_pszPropertyKey);
	MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL;

	long lLength = ptrXMLDOMNodeList->length;
	long lIndex = 0;

	while (lIndex < lLength)
	{
		ptrXMLDOMElement = ptrXMLDOMNodeList->item[lIndex];

		if (ptrXMLDOMElement != NULL) 
		{
			_variant_t varName			= ptrXMLDOMElement->getAttribute(g_pszName);
			_variant_t varLookUpTable	= ptrXMLDOMElement->getAttribute(g_pszLookUpTable);
			_variant_t varOp			= ptrXMLDOMElement->getAttribute(g_pszOperation);
			_bstr_t bstrName			= L"";
			_bstr_t bstrLookUpTable		= L"";
			_bstr_t bstrOp				= g_pszSet;	// Default operation is to set.

			if (varName.vt != VT_NULL && wcslen(varName.bstrVal) > 0)
			{
				bstrName = varName.bstrVal;
			}
			else
			{
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing PROPERTY/@%s: %s"), g_pszName, szLookUpTableMapPath);
			}

			if (varOp.vt != VT_NULL && wcslen(varOp.bstrVal) > 0)
			{
				bstrOp = varOp.bstrVal;
			}

			if (varLookUpTable.vt != VT_NULL && wcslen(varLookUpTable.bstrVal) > 0)
			{
				bstrLookUpTable = varLookUpTable.bstrVal;
			}
			else if (wcsicmp(bstrOp, g_pszDelete) != 0)
			{
				// For all operations except deleting, the new lookup table name must be given.
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data - missing PROPERTY/@%s: %s"), g_pszLookUpTable, szLookUpTableMapPath);
			}

			// Get lookup table (if any) currently associated with this property.
			CLookUpTable* pOldLookUpTable = NULL;
			MapNameToLookUpTableType::iterator it = m_MapPropertyToLookUpTable.find(bstrName);
			if (it != m_MapPropertyToLookUpTable.end())
			{
				pOldLookUpTable = (*it).second;
			}

			// Get new look up table to be associated with this property.
			CLookUpTable* pNewLookUpTable = NULL;
			if (wcsicmp(bstrOp, g_pszDelete) != 0)
			{
				pNewLookUpTable = GetLookUpTableByName(bstrLookUpTable);
			}

			EOp Op = opNull;
			if (wcsicmp(bstrOp, g_pszSet) == 0)
			{
				if (pOldLookUpTable == NULL)
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
				Op = opInsert;
			}
			else if (wcsicmp(bstrOp, g_pszUpdate) == 0)
			{
				Op = opUpdate;
			}
			else if (wcsicmp(bstrOp, g_pszDelete) == 0)
			{
				Op = opDelete;
			}

			if (Op == opInsert && pNewLookUpTable != NULL)
			{
				// Add new lookup table.
				m_MapPropertyToLookUpTable.insert(MapNameToLookUpTableType::value_type(bstrName, pNewLookUpTable));
			}
			else if (Op == opUpdate && pNewLookUpTable != NULL)
			{
				// Replace old lookup table with new one.
				(*it).second = pNewLookUpTable;
			}
			else if (Op == opDelete)
			{
				// Delete existing lookup table.
				m_MapPropertyToLookUpTable.erase(it);
			}
		}

		lIndex++;
	}

	return lIndex;
}

void CLookUpTables::Save(LPCWSTR szLookUpTablesPath)
{
	try
	{
		if (szLookUpTablesPath != NULL)
		{
			m_bstrLookUpTablesPath = szLookUpTablesPath;
		}

		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));
		MSXML::IXMLDOMElementPtr ptrXMLDOMElementParent = ptrXMLDOMDocument->createElement(L"LOOKUPTABLES");
		ptrXMLDOMDocument->appendChild(ptrXMLDOMElementParent);

		if (m_MapPropertyToLookUpTable.size() > 0)
		{
			MapNameToLookUpTableType::iterator it;
			for (it = m_MapPropertyToLookUpTable.begin(); it != m_MapPropertyToLookUpTable.end(); it++)
			{
				_bstr_t bstrName = (*it).first;
				CLookUpTable* pLookUpTable = (*it).second;

				_ASSERTE(pLookUpTable != NULL);

				if (pLookUpTable != NULL)
				{
					MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"PROPERTY");
					ptrXMLDOMElementParent->appendChild(ptrXMLDOMElement);
					ptrXMLDOMElement->setAttribute(g_pszName, bstrName);
					ptrXMLDOMElement->setAttribute(g_pszLookUpTable, pLookUpTable->GetName());
				}
			}
		}

		if (m_MapNameToLookUpTable.size() > 0)
		{
			MapNameToLookUpTableType::iterator it;
			for (it = m_MapNameToLookUpTable.begin(); it != m_MapNameToLookUpTable.end(); it++)
			{
				_bstr_t bstrName = (*it).first;
				CLookUpTable* pLookUpTable = (*it).second;

				_ASSERTE(pLookUpTable != NULL);

				if (pLookUpTable != NULL)
				{
					MSXML::IXMLDOMElementPtr ptrXMLDOMElement = ptrXMLDOMDocument->createElement(L"LOOKUPTABLE");
					ptrXMLDOMElementParent->appendChild(ptrXMLDOMElement);
					ptrXMLDOMElement->setAttribute(g_pszName, bstrName);
					pLookUpTable->Save(ptrXMLDOMElement);
				}
			}
		}

		ptrXMLDOMDocument->save(m_bstrLookUpTablesPath);
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

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::GetLookUpTable
//	
//	Description:
//		Gets the look up table from a map for a key.
//	
//	Parameters:
//		LPCWSTR szName:
//			The key to retrieve the look up table, e.g., the name of a property
//		MapNameToLookUpTableType& MapNameToLookUpTable:
//			The map to use to get the look up table.
//	.
//	Return:
//		CLookUpTable*: 	
//			The look up table for the key.
///////////////////////////////////////////////////////////////////////////////
CLookUpTable* CLookUpTables::GetLookUpTable(LPCWSTR szName, MapNameToLookUpTableType& MapNameToLookUpTable)
{
	CLookUpTable* pLookUpTable = NULL;

	try
	{
		MapNameToLookUpTableType::iterator it = MapNameToLookUpTable.find(szName);
		if (it != MapNameToLookUpTable.end())
		{
			pLookUpTable = (*it).second;
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

	return pLookUpTable;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CLookUpTables::GetLookUpTableValue
//	
//	Description:
//		Get a conversion value from a look up table.
//	
//	Parameters:
//		LPCWSTR szName:
//			The unique name of the look up table, e.g., the name of a property.
//		LPCWSTR szKey:
//			The value to be converted.
//		const CLookUpTable::EDirection Direction:
//			The direction of the conversion; one of CLookUpTable::dirSend or 
//			CLookUpTable::dirRecv. This is required as each look up table has
//			one set of conversion values for when sending data, and one set
//			for when receiving data.
//		MapNameToLookUpTableType& MapNameToLookUpTable:
//			The map to use to get the look up table.
//	
//	Return:
//		_bstr_t: 	
//			The converted value.
///////////////////////////////////////////////////////////////////////////////
_bstr_t CLookUpTables::GetLookUpTableValue(
	LPCWSTR szName, 
	LPCWSTR szKey, 
	const CLookUpTable::EDirection Direction,
	MapNameToLookUpTableType& MapNameToLookUpTable)
{
	_bstr_t bstrValue = L"";

	try
	{
		CLookUpTable* pLookUpTable = NULL;

		if ((pLookUpTable = GetLookUpTable(szName, MapNameToLookUpTable)) != NULL)
		{
			bstrValue = pLookUpTable->GetValue(szKey, Direction);
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

	return bstrValue;
}

void CLookUpTables::WriteCSV()
{
	wchar_t szDrive[MAX_PATH] = L"";
	wchar_t szDir[MAX_PATH] = L"";
	wchar_t szFName[MAX_PATH] = L"";
	wchar_t szExt[MAX_PATH] = L"";
	wchar_t szLookUpTablesCSVPath[MAX_PATH] = L"";

	_wsplitpath(m_bstrLookUpTablesPath, szDrive, szDir, szFName, szExt);
	swprintf(szLookUpTablesCSVPath, L"%s%s%s.csv", szDrive, szDir, szFName, szExt);

	FILE* fp = _wfopen(szLookUpTablesCSVPath, L"wt");

	if (fp != NULL)
	{
		fwprintf(fp, L"Property,Name,Send Source,Receive Source,Item,Send In,Send Out,RecvIn,Recv Out\n\n");

		MapNameToLookUpTableType::iterator it;
		for (it = m_MapPropertyToLookUpTable.begin(); it != m_MapPropertyToLookUpTable.end(); it++)
		{
			_bstr_t bstrProperty		= (*it).first;
			CLookUpTable* pLookUpTable	= (*it).second;

			_ASSERTE(pLookUpTable);

			pLookUpTable->WriteCSV(fp, bstrProperty);
		}

		fclose(fp);
	}
	else
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__, _T("Error opening file: %s"), szLookUpTablesCSVPath);
	}
}
