///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaData.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_METADATA_H__914C4EE3_1701_4BA9_A15F_372DCC2471E4__INCLUDED_)
#define AFX_METADATA_H__914C4EE3_1701_4BA9_A15F_372DCC2471E4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>
#include "mutex.h"

class CMetaDataEnv;
class CRequest;

class CMetaData  
{
public:
	CMetaData();
	virtual ~CMetaData();
	static bool Init();
	static void Save(LPCWSTR szODIConverterXMLPath = NULL);
	static void InitMetaDataEnvs(const MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument);
	static void SetInitialised(bool bInitialised = true) { ::InterlockedExchange(reinterpret_cast<LPLONG>(&s_bInitialised), bInitialised); }
	static bool GetInitialised() { return s_bInitialised; }
	static void SetODIConverterXMLPath(LPCWSTR szODIConverterXMLPath);
	static _bstr_t GetODIConverterXMLPath();
	static _bstr_t GetDefaultOptimusObjectMapPath() { return s_bstrOptimusObjectMapPath; }
	static CMetaDataEnv* GetMetaDataEnv(const MSXML::IXMLDOMDocumentPtr ptrDOMDocument, LPCWSTR szType);
	static CMetaDataEnv* GetMetaDataEnv(const MSXML::IXMLDOMNodePtr ptrNode, LPCWSTR szType);
	static CMetaDataEnv* GetMetaDataEnv(LPCWSTR szName, LPCWSTR szType);
	static CRequest* GetRequest(LPCWSTR szType);

protected:
	typedef std::map<_bstr_t, CMetaDataEnv*, Nocase> MapMetaDataEnvType;
	typedef std::map<_bstr_t, CRequest*, Nocase> MapRequestType;

	static void InitMetaDataEnv(const MSXML::IXMLDOMElementPtr ptrXMLDOMElement);
	static void SaveMetaDataEnv(const MSXML::IXMLDOMElementPtr ptrXMLDOMElement);
	static void InitRequests();
	static void InitRequest(CRequest* pRequest);

private:
    static namespaceMutex::CCriticalSection s_csMetaData;
	static namespaceMutex::CCriticalSection s_csODIConverterXMLPath;

	static bool						s_bInitialised;				// Ensures only initialised once.
	static _bstr_t					s_bstrODIConverterXMLPath;	// Path to configuration meta data.
	static _bstr_t					s_bstrOptimusObjectMapPath;	// Default path to Optimus object map.
	static MapMetaDataEnvType		s_MapMetaDataEnvs;			// Meta data environments.
	static MapRequestType			s_MapRequests;				// Request handlers.
	static _bstr_t					s_bstrDocumentType;
};

#endif // !defined(AFX_METADATA_H__914C4EE3_1701_4BA9_A15F_372DCC2471E4__INCLUDED_)
