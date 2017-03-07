///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaData.cpp
//	DESCRIPTION:	
//		Static class which encapsulates all the meta data used by ODI Converter.
//	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <io.h>
#include "Exception.h"
#include "MetaData.h"
#include "MetaDataEnvOptimus.h"
#include "Profiler.h"
#include "ODIConverter.h"
#include "Request.h"


static LPCWSTR g_pszXMLObjectMapFileName			= L"ObjectMapOSG.xml";
static LPCWSTR g_pszEnvs							= L".//ODIENVIRONMENT";
static LPCWSTR g_pszEnv								= L"ODIENVIRONMENT";
static LPCWSTR g_pszName							= L"NAME";
static LPCWSTR g_pszType							= L"TYPE";
static LPCWSTR g_pszDefault							= L"DEFAULT";

namespaceMutex::CCriticalSection CMetaData::s_csMetaData;
namespaceMutex::CCriticalSection CMetaData::s_csODIConverterXMLPath;

bool CMetaData::s_bInitialised = false;
_bstr_t CMetaData::s_bstrODIConverterXMLPath = L"";
_bstr_t CMetaData::s_bstrOptimusObjectMapPath = L"";
CMetaData::MapMetaDataEnvType CMetaData::s_MapMetaDataEnvs;
CMetaData::MapRequestType CMetaData::s_MapRequests;
_bstr_t CMetaData::s_bstrDocumentType = L"";


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMetaData::CMetaData()
{
}

CMetaData::~CMetaData()
{
	if (s_MapMetaDataEnvs.size() > 0)
	{
		MapMetaDataEnvType::iterator it;
		for (it = s_MapMetaDataEnvs.begin(); it != s_MapMetaDataEnvs.end(); it++)
		{
			CMetaDataEnv* pMetaDataEnv = (*it).second;

			delete pMetaDataEnv;
		}
		s_MapMetaDataEnvs.clear();
	}

	if (s_MapRequests.size() > 0)
	{
		MapRequestType::iterator it;
		for (it = s_MapRequests.begin(); it != s_MapRequests.end(); it++)
		{
			CRequest* pRequest = (*it).second;

			delete pRequest;
		}
		s_MapRequests.clear();
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::Init 
//	 
//	Description:  
//		Initialises meta data.   
//	 
//	Parameters:  
//		None 
//	 
//	Return:  
//		bool:    
//			Returns true if successful. 
///////////////////////////////////////////////////////////////////////////////
bool CMetaData::Init()
{
	try
	{
		if (!s_bInitialised)
		{
			namespaceMutex::CSingleLock	lck(&s_csMetaData, TRUE);

			if (!s_bInitialised)
			{
				s_bstrOptimusObjectMapPath = _Module.MakeModulePath(g_pszXMLObjectMapFileName);

				MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));
				ptrXMLDOMDocument->async = false;
				_variant_t varLoaded = ptrXMLDOMDocument->load(GetODIConverterXMLPath());

				if (varLoaded.boolVal == false)
				{
					throw CException(E_INVALIDMETADATAFILE, __FILE__, __LINE__, _T("Invalid meta data file: %s"), static_cast<LPCWSTR>(GetODIConverterXMLPath()));
				}

				MSXML::IXMLDOMDocumentTypePtr ptrXMLDOMDocumentType = ptrXMLDOMDocument->doctype;
				if (ptrXMLDOMDocumentType != NULL)
				{
					s_bstrDocumentType = ptrXMLDOMDocumentType->xml;
				}

				InitMetaDataEnvs(ptrXMLDOMDocument);

				InitRequests();

				s_bInitialised = true;
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
	/*
	// AS 22/01/2007 VS2005 Port.
	// error C2316: 'exception &' : cannot be caught as the destructor and/or copy constructor are inaccessible.
	catch(exception& e)
	{
		throw CException(E_GENERICERROR, e, __FILE__, __LINE__);
	}
	*/
	catch(...)
	{
		throw CException(E_GENERICERROR, __FILE__, __LINE__);
	}

	return s_bInitialised;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::InitMetaDataEnvs 
//	 
//	Description:  
//		Initialises each meta data environment. A single instance of 
//		ODIConverter can support multiple simultaneous meta data environments.
//		Each meta data environment contains the settings which control how
//		ODIConverter connects to a remote host.
//	 
//	Parameters:  
//		const IXMLDOMDocumentPtr ptrXMLDOMDocument:    
//			DOM Document that contains the meta data for each environment.   
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
void CMetaData::InitMetaDataEnvs(const MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument)
{
	try
	{
		MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMDocument->selectNodes(g_pszEnvs);

		long lLength = ptrXMLDOMNodeList->length;
		if (lLength > 0)
		{
			long lIndex = 0;

			while (lIndex < lLength)
			{
				InitMetaDataEnv(ptrXMLDOMNodeList->item[lIndex]);
				lIndex++;
			}
		}
		else
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszEnvs);
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
}

void CMetaData::InitMetaDataEnv(const MSXML::IXMLDOMElementPtr ptrXMLDOMElement)
{
	try
	{
		if (ptrXMLDOMElement != NULL) 
		{
			// Name is the unique name for the meta data environment.
			// Type corresponds to a meta data class derived from CMetaDataEnv.
			_variant_t varName = ptrXMLDOMElement->getAttribute(g_pszName);
			_variant_t varType = ptrXMLDOMElement->getAttribute(g_pszType);

			if (varName.vt != VT_NULL && varType.vt != VT_NULL)
			{
				_bstr_t bstrName = varName.bstrVal;
				_bstr_t bstrType = varType.bstrVal;

				namespaceMutex::CSingleLock lck(&s_csMetaData, TRUE);

				CMetaDataEnv* pMetaDataEnv = NULL;
				MapMetaDataEnvType::iterator it = s_MapMetaDataEnvs.find(bstrName);
				if (it != s_MapMetaDataEnvs.end())
				{
					pMetaDataEnv = (*it).second;
				}

				bool bNew = false;
				if (pMetaDataEnv == NULL)
				{
					// Env not already in map, so create it.
					pMetaDataEnv = CMetaDataEnv::Create(bstrType, bstrName);
					bNew = true;
				}
				else if (wcsicmp(pMetaDataEnv->GetType(), bstrType) != 0)
				{
					// Use existing env in map; type has changed.
					pMetaDataEnv->SetType(bstrType);
				}

				_ASSERTE(pMetaDataEnv != NULL);

				if (pMetaDataEnv != NULL)
				{
					if (bNew)
					{
						s_MapMetaDataEnvs.insert(MapMetaDataEnvType::value_type(
							bstrName,
							pMetaDataEnv));
					}
					// Unlock map as soon as new env has been added to it.
					lck.Unlock();

					try
					{
						if (!pMetaDataEnv->Init(ptrXMLDOMElement))
						{
							throw CException(E_INVALIDMETADATAENV, __FILE__, __LINE__, _T("Invalid meta data environment: %s"), (LPCWSTR)bstrName);
						}
					}
					catch(...)
					{
						// Error initialising environment.
						if (bNew)
						{
							// Environment has just been added to map, and did not already exist,
							// so remove it from the map, and free memory.
							lck.Lock();
							MapMetaDataEnvType::iterator it = s_MapMetaDataEnvs.find(bstrName);
							if (it != s_MapMetaDataEnvs.end())
							{
								s_MapMetaDataEnvs.erase(it);
							}
							lck.Unlock();
							delete pMetaDataEnv;
						}
						throw;
					}
				}
				else
				{
					throw CException(E_GENERICERROR, __FILE__, __LINE__, _T("Insufficient memory"));
				}
			}
			else
			{
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszType);
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
}

void CMetaData::Save(LPCWSTR szODIConverterXMLPath)
{
	try
	{
		namespaceMutex::CSingleLock	lck(&s_csMetaData, TRUE);

		MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));
		ptrXMLDOMDocument->async = false;
		ptrXMLDOMDocument->preserveWhiteSpace = true;
		_variant_t varLoaded = ptrXMLDOMDocument->load(GetODIConverterXMLPath());

		if (varLoaded.boolVal == false)
		{
			throw CException(E_INVALIDMETADATAFILE, __FILE__, __LINE__, _T("Invalid meta data file: %s"), static_cast<LPCWSTR>(GetODIConverterXMLPath()));
		}

		MSXML::IXMLDOMNodeListPtr ptrXMLDOMNodeList = ptrXMLDOMDocument->selectNodes(g_pszEnvs);

		long lLength = ptrXMLDOMNodeList->length;
		if (lLength > 0)
		{
			long lIndex = 0;

			while (lIndex < lLength)
			{
				SaveMetaDataEnv(ptrXMLDOMNodeList->item[lIndex]);
				lIndex++;
			}
		}
		else
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszEnvs);
		}

		if (szODIConverterXMLPath != NULL)
		{
			// Saving XML to a new file name.
			SetODIConverterXMLPath(szODIConverterXMLPath);
		}

		ptrXMLDOMDocument->save(szODIConverterXMLPath);
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

void CMetaData::SaveMetaDataEnv(const MSXML::IXMLDOMElementPtr ptrXMLDOMElement)
{
	try
	{
		if (ptrXMLDOMElement != NULL) 
		{
			// Name is the unique name for the meta data environment.
			// Type corresponds to a meta data class derived from CMetaDataEnv.
			_variant_t varName = ptrXMLDOMElement->getAttribute(g_pszName);

			if (varName.vt != VT_NULL)
			{
				_bstr_t bstrName = varName.bstrVal;

				CMetaDataEnv* pMetaDataEnv = NULL;
				MapMetaDataEnvType::iterator it = s_MapMetaDataEnvs.find(bstrName);
				if (it != s_MapMetaDataEnvs.end())
				{
					pMetaDataEnv = (*it).second;
				}

				if (pMetaDataEnv == NULL)
				{
					throw CException(E_GENERICERROR, __FILE__, __LINE__, _T("Invalid meta data environment: %s"), (LPCWSTR)bstrName);
				}

				pMetaDataEnv->Save(ptrXMLDOMElement);
			}
			else
			{
				throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data: %s"), g_pszType);
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
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::GetODIConverterXMLPath 
//	 
//	Description:  
//		Returns the location of the XML meta data file. This is the path name  
//		of the ODIConverter executable, with a .xml extension.   
//	 
//	Parameters:  
//		None 
//	 
//	Return:  
//		_bstr_t:    
//			The full path to the XML meta data file. 
///////////////////////////////////////////////////////////////////////////////
_bstr_t CMetaData::GetODIConverterXMLPath()
{
	try
	{
		if (wcslen(s_bstrODIConverterXMLPath) == 0)
		{
			namespaceMutex::CSingleLock lck(&s_csODIConverterXMLPath, TRUE);

			if (wcslen(s_bstrODIConverterXMLPath) == 0)
			{
				_TCHAR szODIConverterXMLPath[_MAX_PATH] = _T("");

				_stprintf(szODIConverterXMLPath, _T("%s%s%s.xml"), _Module.m_szDrive, _Module.m_szDir, _Module.m_szFName);

				if (_taccess(szODIConverterXMLPath, 00) != 0)
				{
					throw CException(E_FILEDOESNOTEXIST, __FILE__, __LINE__, _T("File does not exist: %s"), szODIConverterXMLPath);
				}
				s_bstrODIConverterXMLPath = szODIConverterXMLPath;
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

	return s_bstrODIConverterXMLPath;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::SetODIConverterXMLPath 
//	 
//	Description:  
//		Sets the path to the XML meta data file.   
//	 
//	Parameters:  
//		LPCWSTR szODIConverterXMLPath:    
//			The new path to the XML meta data file.   
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
void CMetaData::SetODIConverterXMLPath(LPCWSTR szODIConverterXMLPath)
{
	namespaceMutex::CSingleLock lck(&s_csODIConverterXMLPath, TRUE);
	s_bstrODIConverterXMLPath = szODIConverterXMLPath;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::GetMetaDataEnv 
//	 
//	Description:  
//		Returns a specific meta data environment.   
//	 
//	Parameters:  
//		const IXMLDOMDocumentPtr ptrDOMDocument:    
//			DOM document. The root node should contain an attribute of the form  
//			ODIENVIRONMENT="<ENV>", where <ENV> is the name of the meta data  
//			environment.   
//
//			A single ODI instance can access multiple remote hosts. For each
//			remote host a meta data environment needs to be defined. The meta 
//			data environment contains settings which control the connection to
//			the remote host.
//
//			ODI Converter can talk to different types of remote host, e.g.,
//			Optimus etc. For each type of remote host a new class needs to be
//			derived from CMetaDataEnv.
//
//			For each meta data environment type ODIConverter supports multiple
//			instances, identified by a unique name.
//
//			For each request ODIConverter receives from the client the meta data
//			environment name needs to specified; this then determines the meta
//			data instance, and hence the meta data type, which then determines
//			the remote host which is used, and how ODIConverter talks to it.
//	 
//	Return:  
//		CMetaDataEnv*:    
//			The meta data environment. 
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnv* CMetaData::GetMetaDataEnv(const MSXML::IXMLDOMDocumentPtr ptrDOMDocument, LPCWSTR szType)
{
	CMetaDataEnv* pMetaDataEnv = NULL;

	try
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement;

		LPCWSTR szName = g_pszDefault;
		ptrXMLDOMElement = ptrDOMDocument->GetdocumentElement();
		if (ptrXMLDOMElement != NULL)
		{
			_variant_t varName = ptrXMLDOMElement->getAttribute(g_pszEnv);

			if (varName.vt != VT_NULL)
			{
				szName = varName.bstrVal;
			}
		}

		pMetaDataEnv = GetMetaDataEnv(szName, szType);
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

	return pMetaDataEnv;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::GetMetaDataEnv 
//	 
//	Description:  
//		Returns a specific meta data environment.   
//	 
//	Parameters:  
//		const IXMLDOMNodePtr ptrNode:    
//			DOM node. The node should contain an attribute of the form  
//			ODIENVIRONMENT="<ENV>", where <ENV> is the name of the meta data  
//			environment.  
//		LPCWSTR szType:
//			The type of the meta data environment. This is used to check that the
//			returned meta data environment instance is of the correct type.
//	 
//	Return:  
//		CMetaDataEnv*:  
//			The meta data environment. 
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnv* CMetaData::GetMetaDataEnv(const MSXML::IXMLDOMNodePtr ptrNode, LPCWSTR szType)
{
	CMetaDataEnv* pMetaDataEnv = NULL;

	try
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElement;

		LPCWSTR szName = static_cast<IXMLAssistDOMNodePtr>(ptrNode)->getAttributeText(g_pszEnv);
		if (wcslen(szName) == 0)
		{
			szName = g_pszDefault;
		}

		pMetaDataEnv = GetMetaDataEnv(szName, szType);
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

	return pMetaDataEnv;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::GetMetaDataEnv 
//	 
//	Description:  
//		Returns a specific meta data environment.   
//	 
//	Parameters:  
//		LPCWSTR szName:
//			The unique name of the meta data environment.
//		LPCWSTR szType:    
//			The type of the meta data environment. This is used to check that the
//			returned meta data environment instance is of the correct type.
//	 
//	Return:  
//		CMetaDataEnv*:  
//			The meta data environment. 
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnv* CMetaData::GetMetaDataEnv(LPCWSTR szName, LPCWSTR szType)
{
	CMetaDataEnv* pMetaDataEnv = NULL;

	try
	{
		MapMetaDataEnvType::iterator it = s_MapMetaDataEnvs.find(szName);
		if (it != s_MapMetaDataEnvs.end())
		{
			pMetaDataEnv = (*it).second;
		}

		if (pMetaDataEnv == NULL)
		{
			throw CException(E_INVALIDMETADATAENV, __FILE__, __LINE__, _T("Invalid meta data environment name: %s"), szName);
		}
		else if (szType != NULL && wcsicmp(pMetaDataEnv->GetType(), szType) != 0)
		{
			throw CException(E_INVALIDMETADATAENV, __FILE__, __LINE__, _T("Invalid meta data environment type %s for name %s"), szType, szName);
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

	return pMetaDataEnv;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::InitRequests 
//	 
//	Description:  
//		Initialises requests. Each request is a method that can be called  
//		through the ODIConverter.ODIConverter COM interface, Request method.   
//	 
//	Parameters:  
//		None 
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
void CMetaData::InitRequests()
{
	try
	{
		InitRequest(CRequest::Create(L"OPTIMUSREQUEST"));
		InitRequest(CRequest::Create(L"TRANSACTREQUEST"));
		InitRequest(CRequest::Create(L"GETCODEPAGETABLE"));
		InitRequest(CRequest::Create(L"DECODEOPTIMUSSTREAM"));
		InitRequest(CRequest::Create(L"THREADWAIT"));
		InitRequest(CRequest::Create(L"STRESS"));
		InitRequest(CRequest::Create(L"CONFIGURE"));
		// ADDHERE Initialisation of other request objects.
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
//	Function: CMetaData::InitRequest 
//	 
//	Description:  
//		Initialise a specific request.   
//	 
//	Parameters:  
//		CRequest* pRequest:    
//			The request object.   
//	 
//	Return:  
//		None 
///////////////////////////////////////////////////////////////////////////////
void CMetaData::InitRequest(CRequest* pRequest)
{
	try
	{
		_ASSERTE(pRequest != NULL);

		if (pRequest != NULL)
		{
			s_MapRequests.insert(MapRequestType::value_type(_wcsupr(pRequest->GetName()), pRequest));
		}
		else
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid request"));
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
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaData::GetRequest 
//	 
//	Description:  
//		Returns a request object.   
//	 
//	Parameters:  
//		LPCWSTR szType:    
//			The unique name of the request object.   
//	 
//	Return:  
//		CRequest*:    
//			The request object. 
///////////////////////////////////////////////////////////////////////////////
CRequest* CMetaData::GetRequest(LPCWSTR szType)
{
	CRequest* pRequest = NULL;

	try
	{
		MapRequestType::iterator it = s_MapRequests.find(szType);
		if (it != s_MapRequests.end())
		{
			pRequest = (*it).second;
		}

		if (pRequest == NULL)
		{
			throw CException(E_INVALIDREQUEST, __FILE__, __LINE__, _T("Invalid request type: %s"), szType);
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

	return pRequest;
}

