///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusMetaData.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      03/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_OPTIMUSMETADATA_H__64A6A6A9_093D_470B_A25F_E36ECDEF4A59__INCLUDED_)
#define AFX_OPTIMUSMETADATA_H__64A6A6A9_093D_470B_A25F_E36ECDEF4A59__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>
#include "mutex.h"

class CMetaDataEnvOptimus;
class CMetaDataEnv;
class CLookUpTable;

///////////////////////////////////////////////////////////////////////////////

class COptimusMetaData  
{
public:
	COptimusMetaData();
	void Init(CMetaDataEnvOptimus* pMetaDataEnvOptimus);
	_bstr_t GetOptimusObjectMapPath() const { return m_bstrOptimusObjectMapPath; }
	void SetOptimusObjectMapPath(LPCWSTR szOptimusObjectMapPath) { m_bstrOptimusObjectMapPath = szOptimusObjectMapPath; }
	_bstr_t GetObjectName(short nObjectId);
	_bstr_t GetPropertyName(short nObjectId, short nPropertyId, CLookUpTable*& pLookUpTable);
	short GetObjectId(LPCWSTR szObjectName);
	long GetObjectPropertyId(short nObjectId, LPCWSTR szPropertyName, CLookUpTable*& pLookUpTable);
	long GetObjectPropertyId(LPCWSTR szObjectName, LPCWSTR szPropertyName, CLookUpTable*& pLookUpTable);
	short GetObjectPropertyIdObjectId(long lObjectPropertyId)
	{
		return LOWORD(lObjectPropertyId);
	}
	short GetObjectPropertyIdPropertyId(long lObjectPropertyId)
	{
		return HIWORD(lObjectPropertyId);
	}
	_bstr_t MakeObjectPropertyName(short nObjectId, LPCWSTR szPropertyName);
	_bstr_t MakeObjectPropertyName(LPCWSTR szObjectName, LPCWSTR szPropertyName);
	_bstr_t MakeObjectIdPropertyName(LPCWSTR szObjectName, LPCWSTR szPropertyName);
	_bstr_t MakeObjectIdPropertyName(short nObjectId, LPCWSTR szPropertyName);

protected:
	void InitMapObjectIdToObjectName();
	void InitMapObjectPropertyIdToPropertyName();
	void InitMapObjectPropertyIdToPropertyName(short nObjectId, MSXML::IXMLDOMElementPtr elemObjectPtr);
	void InitMapObjectNameToObjectId();
	void InitMapPropertyNameToObjectPropertyId();
	void InitMapPropertyNameToObjectPropertyId(short nObjectId, MSXML::IXMLDOMElementPtr elemObjectPtr);

	typedef std::pair<_bstr_t, CLookUpTable*> PairPropertyNameType;
	typedef std::pair<long, CLookUpTable*> PairPropertyIdType;

	typedef std::map<short, _bstr_t> MapObjectIdToObjectNameType;
	typedef std::map<long, PairPropertyNameType> MapObjectPropertyIdToPropertyNameType;				// long is composed of two shorts ((nObjectId << 16) | nPropertyId)
	typedef std::map<_bstr_t, short, Nocase> MapObjectNameToObjectIdType;
	typedef std::map<_bstr_t, PairPropertyIdType, Nocase> MapPropertyNameToObjectPropertyIdType;	// long is composed of two shorts ((nObjectId << 16) | nPropertyId)

	MapObjectIdToObjectNameType& GetMapObjectIdToObjectName()
	{
		return m_MapObjectIdToObjectName;
	}

	MapObjectPropertyIdToPropertyNameType& GetMapObjectPropertyIdToPropertyName()
	{
		return m_MapObjectPropertyIdToPropertyName;
	}

	MapObjectNameToObjectIdType& GetMapObjectNameToObjectId()
	{
		return m_MapObjectNameToObjectId;
	}

	MapPropertyNameToObjectPropertyIdType& GetMapPropertyNameToObjectPropertyId()
	{
		return m_MapPropertyNameToObjectPropertyId;
	}

	long MakeObjectPropertyId(short nObjectId, short nPropertyId) 
	{
		// Long is composed of two shorts; Object Id is low word, Property Id is high word.
		return MAKELONG(nObjectId, nPropertyId);
	}

private:
	CMetaDataEnvOptimus*					m_pMetaDataEnvOptimus;
	_bstr_t									m_bstrOptimusObjectMapPath;
	MapObjectIdToObjectNameType				m_MapObjectIdToObjectName;
	MapObjectPropertyIdToPropertyNameType	m_MapObjectPropertyIdToPropertyName;
	MapObjectNameToObjectIdType				m_MapObjectNameToObjectId;
	MapPropertyNameToObjectPropertyIdType	m_MapPropertyNameToObjectPropertyId;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_OPTIMUSMETADATA_H__64A6A6A9_093D_470B_A25F_E36ECDEF4A59__INCLUDED_)
