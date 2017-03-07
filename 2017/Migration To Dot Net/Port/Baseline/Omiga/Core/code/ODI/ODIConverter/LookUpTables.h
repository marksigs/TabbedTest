// LookUpTables.h: interface for the CLookUpTables class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_LOOKUPTABLES_H__F7DC1D6E_23B7_4460_A569_74125D28C1CD__INCLUDED_)
#define AFX_LOOKUPTABLES_H__F7DC1D6E_23B7_4460_A569_74125D28C1CD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>
#include "LookUpTable.h"

class CLookUpTables  
{
public:
	CLookUpTables();
	virtual ~CLookUpTables();
	long Init(LPCWSTR szLookUpTablesPath);
	void Save(LPCWSTR szLookUpTablesPath);
	long InitNode(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath = NULL);

	typedef std::map<_bstr_t, CLookUpTable*, Nocase> MapNameToLookUpTableType;
	MapNameToLookUpTableType& GetMapNameToLookUpTable()
	{
		return m_MapNameToLookUpTable;
	}
	MapNameToLookUpTableType& GetMapPropertyToLookUpTable()
	{
		return m_MapPropertyToLookUpTable;
	}

	CLookUpTable* GetLookUpTableByName(LPCWSTR szName)
	{
		return GetLookUpTable(szName, m_MapNameToLookUpTable);
	}
	CLookUpTable* GetLookUpTableByProperty(LPCWSTR szName)
	{
		return GetLookUpTable(szName, m_MapPropertyToLookUpTable);
	}

	_bstr_t GetLookUpTableByNameValue(
		LPCWSTR szName, 
		LPCWSTR szKey, 
		const CLookUpTable::EDirection Direction)
	{
		return GetLookUpTableValue(szName, szKey, Direction, m_MapNameToLookUpTable);
	}
	_bstr_t GetLookUpTableByPropertyValue(
		LPCWSTR szName, 
		LPCWSTR szKey, 
		const CLookUpTable::EDirection Direction)
	{
		return GetLookUpTableValue(szName, szKey, Direction, m_MapPropertyToLookUpTable);
	}

	void WriteCSV();

protected:
	long InitLookUpTables(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath);
	long InitProperties(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szLookUpTableMapPath);
	CLookUpTable* GetLookUpTable(LPCWSTR szName, MapNameToLookUpTableType& MapNameToLookUpTable);
	_bstr_t GetLookUpTableValue(
		LPCWSTR szName, 
		LPCWSTR szKey, 
		const CLookUpTable::EDirection Direction,
		MapNameToLookUpTableType& MapNameToLookUpTable);
	void DeletePropertiesForLookUpTable(CLookUpTable* pLookUpTable);

private:
	enum EOp { opNull, opInsert, opUpdate, opDelete };
	MapNameToLookUpTableType			m_MapNameToLookUpTable;
	MapNameToLookUpTableType			m_MapPropertyToLookUpTable;
	_bstr_t								m_bstrLookUpTablesPath;
};

#endif // !defined(AFX_LOOKUPTABLES_H__F7DC1D6E_23B7_4460_A569_74125D28C1CD__INCLUDED_)
