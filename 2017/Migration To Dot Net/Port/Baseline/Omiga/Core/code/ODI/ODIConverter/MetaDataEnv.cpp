///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaDataEnv.cpp
//	DESCRIPTION:	Based class for meta data environments.
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		28/09/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Exception.h"
#include "MetaDataEnv.h"
#include "MetaDataEnvOptimus.h"
#include "MetaDataEnvTransact.h"
#include "ODIConverter.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMetaDataEnv::CMetaDataEnv(LPCWSTR szType, LPCWSTR szName) :
	m_bstrType(szType),
	m_bstrName(szName),
	m_bInitialised(false)
{
}

CMetaDataEnv::~CMetaDataEnv()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Function: CMetaDataEnv::Create
//	
//	Description:
//		Static virtual constructor for meta data environments.
//	
//	Parameters:
//		LPCWSTR szType:
//			The type of meta data environment to create.
//		LPCWSTR szName:
//			The unique name for the new meta data environment object.
//	
//	Return:
//		CMetaDataEnv*: 	
//			The newly created meta data environment.
///////////////////////////////////////////////////////////////////////////////
CMetaDataEnv* CMetaDataEnv::Create(LPCWSTR szType, LPCWSTR szName)
{
	CMetaDataEnv* pMetaDataEnv = NULL;

	try
	{
		static const LPCWSTR pszOptimus		= L"OPTIMUS";
		static const LPCWSTR pszTransact	= L"TRANSACT";
		
		if (wcsicmp(szType, pszOptimus) == 0)
		{
			pMetaDataEnv = new CMetaDataEnvOptimus(szType, szName);
		}
		else if (wcsicmp(szType, pszTransact) == 0)
		{
			pMetaDataEnv = new CMetaDataEnvTransact(szType, szName);
		}
		// ADDHERE Contruction of other meta data environment objects.
		else
		{
			throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data environment type: %s"), szType);
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
		throw CException(E_INVALIDMETADATA, __FILE__, __LINE__, _T("Invalid meta data environment type: %s"), szType);
	}

	return pMetaDataEnv;
}

