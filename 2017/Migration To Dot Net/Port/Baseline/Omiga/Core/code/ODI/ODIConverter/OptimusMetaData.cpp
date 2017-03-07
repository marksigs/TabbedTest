///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusMetaData.cpp
//	DESCRIPTION:	
//		Encapsulates the Optimus meta data in the ObjectMapOSG.xml. The
//		the ObjectMapOSG.xml is automatically generated from Optimus, and
//		maps Optimus object and property ids to object and property names.
//		
//		The object and property ids are used in the serialized objects byte 
//		stream ODI sends and receives to and from Optimus. The object and
//		property names are used in the XML representation of the objects that
//		ODI sends and receives to and from Omiga.
//
//		The COptimusMetaData class provides a fast method, using STL maps,
//		of looking up object and property ids based on the Optimus XML meta data.
//
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      03/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "LookUpTables.h"
#include "MetaData.h"
#include "MetaDataEnv.h"
#include "MetaDataEnvOptimus.h"
#include "OptimusMetaData.h"
#include "ODIConverter.h"


static LPCWSTR g_pszXPathObject				= L"OptimusObjects/Object";
static LPCWSTR g_pszattributeObjectId		= L"id";
static LPCWSTR g_pszattributeObjectName		= L"name";

static LPCWSTR g_pszXPathProperty			= L"./Property";
static LPCWSTR g_pszattributePropertyId		= L"id";
static LPCWSTR g_pszattributePropertyName	= L"name";


///////////////////////////////////////////////////////////////////////////////
// Initialise static variables.

///////////////////////////////////////////////////////////////////////////////
//Additional Mappings not defined in external XML file

struct ObjectDataType
{
	short nObjectId;
	LPCWSTR pszObjectName;
};

static ObjectDataType ObjectData[] = 
{
	2000, L"MSG_ARRAY",
	2003, L"MSG_SIGNONPROFILE",
	NULL, NULL
};


struct PropertyDataType
{
	short nObjectId;
	short nPropertyId;
	LPCWSTR pszPropertyName;
};

static PropertyDataType PropertyData[] = 
{
	// Request
	2001, 2, L"MSG_METHOD",
	2001, 3, L"MSG_SESSION",
	2001, 4, L"MSG_TARGET",

	// SignOnProfile
	2003, 1, L"MSG_HOST",
	2003, 2, L"MSG_ENVIRONMENT",
	2003, 3, L"MSG_USERNAME",
	2003, 4, L"MSG_PASSWORD",

	// SessionImpl
	2004, 1, L"MSG_ACCOUNTINGDATE",
	2004, 2, L"MSG_EFFECTIVEDATE",
	2004, 3, L"MSG_SESSIONID",
	2004, 4, L"MSG_USERKEY",
	2004, 5, L"MSG_OBJECT",
	2004, 6, L"MSG_SUPPORTEDLOCALES",
	2004, 7, L"MSG_PORT",
	2004, 8, L"MSG_PLEXUS",

	NULL, NULL, NULL
};


COptimusMetaData::COptimusMetaData() :
	m_pMetaDataEnvOptimus(NULL)
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::Init
//	
//	Description:
//		Overall initialisation.
//	
//	Parameters:
//		LPCWSTR szOptimusObjectMapPath:
//			Location of Optimus object map XML file.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::Init(CMetaDataEnvOptimus* pMetaDataEnvOptimus)
{
	_ASSERTE(pMetaDataEnvOptimus != NULL);

	m_pMetaDataEnvOptimus = pMetaDataEnvOptimus;
	SetOptimusObjectMapPath(pMetaDataEnvOptimus->GetOptimusObjectMapPath());

	InitMapObjectIdToObjectName();
	InitMapObjectPropertyIdToPropertyName();
	InitMapObjectNameToObjectId();
	InitMapPropertyNameToObjectPropertyId();
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapObjectIdToObjectName
//	
//	Description:
//		Creates the mapping of object ids, e.g., 2001, to object names, 
//		e.g., "Request".
//	
//	Parameters:
//		None
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapObjectIdToObjectName()
{
	MSXML::IXMLDOMDocumentPtr DOMDocumentPtr(__uuidof(MSXML::DOMDocument));

	DOMDocumentPtr->async = false;
	_variant_t varLoaded = DOMDocumentPtr->load(m_bstrOptimusObjectMapPath);
	if (varLoaded.boolVal == false)
	{
		throw CException(E_INVALIDMETADATAFILE, __FILE__, __LINE__, _T("Invalid meta data file: %s"), static_cast<LPCWSTR>(m_bstrOptimusObjectMapPath));
	}

	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = DOMDocumentPtr->selectNodes(g_pszXPathObject);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr = NULL;
	
	// For each object.
	while (lIndex < lLength)
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtObjectId = XMLDOMElementPtr->getAttribute(g_pszattributeObjectId);
		_variant_t vtObjectName = XMLDOMElementPtr->getAttribute(g_pszattributeObjectName);

		GetMapObjectIdToObjectName().insert(MapObjectIdToObjectNameType::value_type(
			static_cast<short>(vtObjectId), 
			_wcsupr(static_cast<_bstr_t>(vtObjectName))));
		lIndex++;
	}

	// Initialise additional definitions from static data defined above.
	// These are only required where there are omissions in the automatically generated
	// Optimus object XML map.
	int nIndex = 0;
	while (ObjectData[nIndex].pszObjectName != NULL)
	{
		GetMapObjectIdToObjectName().insert(MapObjectIdToObjectNameType::value_type(
			ObjectData[nIndex].nObjectId, 
			ObjectData[nIndex].pszObjectName));
		nIndex++;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapObjectPropertyIdToPropertyName
//	
//	Description:
//		Creates mapping of object+property ids to property names. The object and
//		property id, both shorts, are combined in a single long, where the
//		object id is the low word, and the property id is the high word. The
//		long is used as the key for the map entry that references the property 
//		name. The key needs to include the object id as the same name can be 
//		used by properties in different objects.
//	
//	Parameters:
//		None
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapObjectPropertyIdToPropertyName()
{
	MSXML::IXMLDOMDocumentPtr DOMDocumentPtr(__uuidof(MSXML::DOMDocument));
	DOMDocumentPtr->async = false;
	DOMDocumentPtr->load(m_bstrOptimusObjectMapPath);

	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = DOMDocumentPtr->selectNodes(g_pszXPathObject);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr;

	// For each object.
	while (lIndex < lLength)
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtObjectId = XMLDOMElementPtr->getAttribute(g_pszattributeObjectId);
		InitMapObjectPropertyIdToPropertyName(static_cast<short>(vtObjectId), XMLDOMElementPtr);
		lIndex++;
	}

	// Initialise additional definitions from static data defined above.
	// These are only required where there are omissions in the automatically generated
	// Optimus object XML map.
	int nIndex = 0;
	while (PropertyData[nIndex].pszPropertyName != NULL)
	{
		// Get any look up table for this property.
		CLookUpTable* pLookUpTable = 
			m_pMetaDataEnvOptimus->GetLookUpTables().
				GetLookUpTableByProperty(MakeObjectPropertyName(PropertyData[nIndex].nObjectId, PropertyData[nIndex].pszPropertyName));
		
		GetMapObjectPropertyIdToPropertyName().insert(MapObjectPropertyIdToPropertyNameType::value_type(
			MakeObjectPropertyId(PropertyData[nIndex].nObjectId, -PropertyData[nIndex].nPropertyId),
			PairPropertyNameType(PropertyData[nIndex].pszPropertyName, pLookUpTable)));
		nIndex++;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapObjectPropertyIdToPropertyName
//	
//	Description:
//		For a specific object, creates mapping of object+property id for all the
//		properties in the object.
//	
//	Parameters:
//		short nObjectId:
//			The id of the object to which the properties belong.
//		IXMLDOMElementPtr elemObjectPtr:
//			The XML for the object, which maps the ids of the properties in the
//			object to property names. 
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapObjectPropertyIdToPropertyName(short nObjectId, MSXML::IXMLDOMElementPtr elemObjectPtr)
{
	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = elemObjectPtr->selectNodes(g_pszXPathProperty);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr = NULL;

	// For each property in this object.
	while (lIndex < lLength)
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtPropertyId		= XMLDOMElementPtr->getAttribute(g_pszattributePropertyId);
		_variant_t vtPropertyName	= XMLDOMElementPtr->getAttribute(g_pszattributePropertyName);
		_bstr_t bstrPropertyName	= _wcsupr(vtPropertyName.bstrVal);

		// Get any look up table for this property.
		CLookUpTable* pLookUpTable = 
			m_pMetaDataEnvOptimus->GetLookUpTables().GetLookUpTableByProperty(
				MakeObjectPropertyName(nObjectId, bstrPropertyName));
		
		GetMapObjectPropertyIdToPropertyName().insert(MapObjectPropertyIdToPropertyNameType::value_type(
			MakeObjectPropertyId(nObjectId, -static_cast<short>(vtPropertyId)),
			PairPropertyNameType(bstrPropertyName, pLookUpTable)));

		lIndex++;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapObjectNameToObjectId
//	
//	Description:
//		Creates reverse mapping of object names to object ids.
//	
//	Parameters:
//		None
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapObjectNameToObjectId()
{
	MSXML::IXMLDOMDocumentPtr DOMDocumentPtr(__uuidof(MSXML::DOMDocument));
	DOMDocumentPtr->async = false;
	DOMDocumentPtr->load(m_bstrOptimusObjectMapPath);

	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = DOMDocumentPtr->selectNodes(g_pszXPathObject);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr = NULL;

	// For each object.
	while (lIndex < lLength)
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtObjectId = XMLDOMElementPtr->getAttribute(g_pszattributeObjectId);
		_variant_t vtObjectName = XMLDOMElementPtr->getAttribute(g_pszattributeObjectName);

		GetMapObjectNameToObjectId().insert(MapObjectNameToObjectIdType::value_type(
			_wcsupr(static_cast<_bstr_t>(vtObjectName)),
			static_cast<short>(vtObjectId)));
		lIndex++;
	}

	// Initialise additional definitions from static data defined above.
	// These are only required where there are omissions in the automatically generated
	// Optimus object XML map.
	int nIndex = 0;
	while (ObjectData[nIndex].pszObjectName != NULL)
	{
		GetMapObjectNameToObjectId().insert(MapObjectNameToObjectIdType::value_type(
			ObjectData[nIndex].pszObjectName,
			ObjectData[nIndex].nObjectId));
		nIndex++;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapPropertyNameToObjectPropertyId
//	
//	Description:
//		Creates reverse mapping of property names to property ids. The property 
//		name includes the id of the object to which the property belongs, in the
//		form <object id>.<property name>, e.g., "2001.method". The object id 
//		must be included to make the key for the map entry unique, as the same
//		property name can occur in more than one object (so is not sufficiently
//		unique).
//	
//	Parameters:
//		None
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapPropertyNameToObjectPropertyId()
{
	MSXML::IXMLDOMDocumentPtr DOMDocumentPtr(__uuidof(MSXML::DOMDocument));
	DOMDocumentPtr->async = false;
	DOMDocumentPtr->load(m_bstrOptimusObjectMapPath);

	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = DOMDocumentPtr->selectNodes(g_pszXPathObject);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr = NULL;

	// For each object.
	while (lIndex < lLength)
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtObjectId = XMLDOMElementPtr->getAttribute(g_pszattributeObjectId);
		InitMapPropertyNameToObjectPropertyId(static_cast<short>(vtObjectId), XMLDOMElementPtr);
		lIndex++;
	}

	// Initialise additional definitions from static data defined above.
	// These are only required where there are omissions in the automatically generated
	// Optimus object XML map.
	int nIndex = 0;
	while (PropertyData[nIndex].pszPropertyName != NULL)
	{
		// Get any look up table for this property.
		CLookUpTable* pLookUpTable = 
			m_pMetaDataEnvOptimus->GetLookUpTables().GetLookUpTableByProperty(
				MakeObjectPropertyName(
					PropertyData[nIndex].nObjectId, 
					PropertyData[nIndex].pszPropertyName));

		GetMapPropertyNameToObjectPropertyId().insert(MapPropertyNameToObjectPropertyIdType::value_type(
			MakeObjectIdPropertyName(
				PropertyData[nIndex].nObjectId, PropertyData[nIndex].pszPropertyName),
			PairPropertyIdType(
				MakeObjectPropertyId(PropertyData[nIndex].nObjectId, -PropertyData[nIndex].nPropertyId),
				pLookUpTable)));
		nIndex++;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::InitMapPropertyNameToObjectPropertyId	
//	
//	Description:
//		For a specific object, creates the reverse mapping of property names to 
//		property ids.
//	
//	Parameters:
//		short nObjectId:
//			The id of the object to which the properties belong.
//		IXMLDOMElementPtr elemObjectPtr:
//			The XML for the object, which maps the ids of the properties in the
//			object to property names. 
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusMetaData::InitMapPropertyNameToObjectPropertyId(short nObjectId, MSXML::IXMLDOMElementPtr elemObjectPtr)
{
	MSXML::IXMLDOMNodeListPtr XMLDOMNodeListPtr = elemObjectPtr->selectNodes(g_pszXPathProperty);
	long lLength = XMLDOMNodeListPtr->length;
	long lIndex = 0;
	MSXML::IXMLDOMElementPtr XMLDOMElementPtr;

	// For each property in the object.
	while (lIndex < lLength) 
	{
		XMLDOMElementPtr = XMLDOMNodeListPtr->item[lIndex];
		_variant_t vtPropertyId		= XMLDOMElementPtr->getAttribute(g_pszattributePropertyId);
		_variant_t vtPropertyName	= XMLDOMElementPtr->getAttribute(g_pszattributePropertyName);
		_bstr_t bstrPropertyName	= _wcsupr(vtPropertyName.bstrVal);

		// Get any look up table for this property.
		CLookUpTable* pLookUpTable = 
			m_pMetaDataEnvOptimus->GetLookUpTables().GetLookUpTableByProperty(
				MakeObjectPropertyName(nObjectId, bstrPropertyName));

		GetMapPropertyNameToObjectPropertyId().insert(MapPropertyNameToObjectPropertyIdType::value_type(
			MakeObjectIdPropertyName(nObjectId, bstrPropertyName),
			PairPropertyIdType(
				MakeObjectPropertyId(nObjectId, -static_cast<short>(vtPropertyId)),
				pLookUpTable)));

		lIndex++;
	}
}


///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::GetObjectName
//	
//	Description:
//		Gets the name for an Optimus object from its unique id.
//	
//	Parameters:
//		short nObjectId:
//			The unique object id.
//	
//	Return:
//		_bstr_t: 	
//			The name of the object.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::GetObjectName(short nObjectId)
{
	MapObjectIdToObjectNameType& MapObjectIdToObjectName = GetMapObjectIdToObjectName();

	MapObjectIdToObjectNameType::iterator it = MapObjectIdToObjectName.find(nObjectId);
	if (it == MapObjectIdToObjectName.end())
	{
		wchar_t buff[64];
		swprintf(buff, L"OBJECTNAMENOTFOUND: %d", nObjectId);
		return buff;
	}
	else
	{
		return (*it).second;
	}	
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::GetPropertyName
//	
//	Description:
//		Gets the name of an Optimus property given the id of the object the 
//		property belongs to, and the id of the property itself.
//	
//	Parameters:
//		short nObjectId:
//			The unique id of the property's parent object.
//		short nPropertyId:
//			The id of the property. Only unique within the parent object.
//		CLookUpTable*& pLookUpTable:
//			Output: The look up table for the property. Will be NULL if 
//			values of this property should not be automatically converted.
//	
//	Return:
//		_bstr_t: 	
//			The name of the property.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::GetPropertyName(short nObjectId, short nPropertyId, CLookUpTable*& pLookUpTable)
{
	pLookUpTable = NULL;
	long lObjectPropertyId = MakeObjectPropertyId(nObjectId, nPropertyId);
	MapObjectPropertyIdToPropertyNameType& MapObjectPropertyIdToPropertyName = GetMapObjectPropertyIdToPropertyName();

	MapObjectPropertyIdToPropertyNameType::iterator it = MapObjectPropertyIdToPropertyName.find(lObjectPropertyId);
	if (it == MapObjectPropertyIdToPropertyName.end())
	{
		wchar_t buff[64];
		swprintf(buff, L"PROPERTYNAMENOTFOUND: %d, %d, %d", nObjectId, nPropertyId, lObjectPropertyId);
		return buff;
	}
	else
	{
		PairPropertyNameType PairPropertyName = (*it).second;
		// First item in the pair is the property name.
		// Second item in the pair is any look up table used for converting values of this
		// property.
		pLookUpTable = PairPropertyName.second;
		return PairPropertyName.first;
	}	
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::GetObjectId
//	
//	Description:
//		Gets the unique id for an Optimus object, given the name of the object.
//	
//	Parameters:
//		LPCWSTR szObjectName:
//			The name of the object.
//	
//	Return:
//		short: 	
//			The object's id.
///////////////////////////////////////////////////////////////////////////////
short COptimusMetaData::GetObjectId(LPCWSTR szObjectName)
{
	MapObjectNameToObjectIdType& MapObjectNameToObjectId = GetMapObjectNameToObjectId();

	MapObjectNameToObjectIdType::iterator it = MapObjectNameToObjectId.find(szObjectName);
	if (it == MapObjectNameToObjectId.end())
	{
		return 0;
	}
	else
	{
		return (*it).second;
	}	
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::GetObjectPropertyId
//	
//	Description:
//		Gets the id for an Optimus property, given the id of the parent object,
//		and the name of the property. 
//	
//	Parameters:
//		short nObjectId:
//			The unique id of the parent object.
//		LPCWSTR szPropertyName:
//			The name of the property. Only unique within the parent object.
//		CLookUpTable*& pLookUpTable:
//			Output: The look up table for the property. Will be NULL if 
//			values of this property should not be automatically converted.
//	
//	Return:
//		long: 	
//			The id of the property. Only unique within the parent object.
///////////////////////////////////////////////////////////////////////////////
long COptimusMetaData::GetObjectPropertyId(short nObjectId, LPCWSTR szPropertyName, CLookUpTable*& pLookUpTable)
{
	pLookUpTable = NULL;

	MapPropertyNameToObjectPropertyIdType& MapPropertyNameToObjectPropertyId = GetMapPropertyNameToObjectPropertyId();

	MapPropertyNameToObjectPropertyIdType::iterator it = MapPropertyNameToObjectPropertyId.find(MakeObjectIdPropertyName(nObjectId, szPropertyName));
	if (it == MapPropertyNameToObjectPropertyId.end())
	{
		return 0;
	}
	else
	{
		PairPropertyIdType PairPropertyId = (*it).second;
		pLookUpTable = PairPropertyId.second;
		return PairPropertyId.first;
	}	
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::GetObjectPropertyId
//	
//	Description:
//		Gets the id for an Optimus property, given the name of the parent object,
//		and the name of the property. 
//	
//	Parameters:
//		LPCWSTR szObjectName:
//			The unique name of the parent object.
//		LPCWSTR szPropertyName:
//			The name of the property. Only unique within the parent object.
//		CLookUpTable*& pLookUpTable:
//			Output: The look up table for the property. Will be NULL if 
//			values of this property should not be automatically converted.
//	
//	Return:
//		long: 	
//			The id of the property. Only unique within the parent object.
///////////////////////////////////////////////////////////////////////////////
long COptimusMetaData::GetObjectPropertyId(LPCWSTR szObjectName, LPCWSTR szPropertyName, CLookUpTable*& pLookUpTable)
{
	pLookUpTable = NULL;

	MapPropertyNameToObjectPropertyIdType& MapPropertyNameToObjectPropertyId = GetMapPropertyNameToObjectPropertyId();

	MapPropertyNameToObjectPropertyIdType::iterator it = MapPropertyNameToObjectPropertyId.find(MakeObjectIdPropertyName(szObjectName, szPropertyName));
	if (it == MapPropertyNameToObjectPropertyId.end())
	{
		return 0;
	}
	else
	{
		PairPropertyIdType PairPropertyId = (*it).second;
		pLookUpTable = PairPropertyId.second;
		return PairPropertyId.first;
	}	
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::MakeObjectPropertyName
//	
//	Description:
//		Creates a property name, given the id of the parent object and
//		the name of the property. The unique property name is in the form 
//		<object name>.<property name>, e.g., "Request.Method".
//	
//	Parameters:
//		short nObjectId:
//			The id of the property's parent object.
//		LPCWSTR szPropertyName:
//			The name of the property.
//	
//	Return:
//		_bstr_t: 	
//			The property name, including the parent object name.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::MakeObjectPropertyName(short nObjectId, LPCWSTR szPropertyName)
{
	return MakeObjectPropertyName(GetObjectName(nObjectId), szPropertyName);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::MakeObjectPropertyName
//	
//	Description:
//		Creates a property name, given the name of the parent object and
//		the name of the property. The unique property name is in the form 
//		<object name>.<property name>, e.g., "Request.Method".
//	
//	Parameters:
//		LPCWSTR szObjectName:
//			The name of the property's parent object.
//		LPCWSTR szPropertyName:
//			The name of the property.
//	
//	Return:
//		_bstr_t: 	
//			The property name, including the parent object name.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::MakeObjectPropertyName(LPCWSTR szObjectName, LPCWSTR szPropertyName)
{
	_bstr_t bstrObjectPropertyName;
	
	bstrObjectPropertyName = szObjectName;
	bstrObjectPropertyName += L".";
	bstrObjectPropertyName += szPropertyName;

	return _wcsupr(bstrObjectPropertyName);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::MakeObjectIdPropertyName
//	
//	Description:
//		Creates a unique property name, given the name of the parent object and
//		the name of the property. The unique property name is in the form 
//		<object id>.<property name>, e.g., "2001.Method".
//	
//	Parameters:
//		LPCWSTR szObjectName:
//			The name of the property's parent object.
//		LPCWSTR szPropertyName:
//			The name of the property.
//	
//	Return:
//		_bstr_t: 	
//			The unique property name, including the parent object id.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::MakeObjectIdPropertyName(LPCWSTR szObjectName, LPCWSTR szPropertyName)
{
	return MakeObjectIdPropertyName(GetObjectId(szObjectName), szPropertyName);
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusMetaData::MakeObjectIdPropertyName
//	
//	Description:
//		Creates a unique property name, given the id of the parent object and
//		the name of the property. The unique property name is in the form 
//		<object id>.<property name>, e.g., "2001.Method".
//	
//	Parameters:
//		LPCWSTR szObjectName:
//			The name of the property's parent object.
//		LPCWSTR szPropertyName:
//			The name of the property.
//	
//	Return:
//		_bstr_t: 	
//			The unique property name, including the parent object id.
///////////////////////////////////////////////////////////////////////////////
_bstr_t COptimusMetaData::MakeObjectIdPropertyName(short nObjectId, LPCWSTR szPropertyName)
{
	_bstr_t bstrObjectIdPropertyName;
	
	wchar_t szObjectId[16] = L"";
	swprintf(szObjectId, L"%d", nObjectId);
	bstrObjectIdPropertyName = szObjectId;
	bstrObjectIdPropertyName += L".";
	bstrObjectIdPropertyName += szPropertyName;

	return _wcsupr(bstrObjectIdPropertyName);
}

