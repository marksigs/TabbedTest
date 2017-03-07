///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerMTS1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//	LD		10/09/01	SYS2705 - Support for SQL Server added
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		09/02/02	SYS4414	- Add s_vtNull
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <stdio.h>
#include "MessageQueueListenerMTS.h"
#include "MessageQueueListenerMTS1.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerMTS1

#import "..\MessageQueueComponentVC\MessageQueueComponentVC.tlb" no_namespace

_bstr_t CMessageQueueListenerMTS1::s_bstrNULL(static_cast<LPWSTR>(NULL));
_variant_t CMessageQueueListenerMTS1::s_vZero(0L);
_variant_t CMessageQueueListenerMTS1::s_vOne(1L);
_variant_t InitVariantNull()
{
	VARIANT vt;
	VariantInit(&vt);
	vt.vt = VT_NULL;
	return _variant_t(vt, false);
}
_variant_t CMessageQueueListenerMTS1::s_vNull = InitVariantNull();

const long s_lMSMQPeekTimeOut = 1000L; // in ms
const long s_lMSMQReceiveTimeOut = INFINITE; // in ms

///////////////////////////////////////////////////////////////////////////////

HRESULT CMessageQueueListenerMTS1::FinalConstruct()
{
	try
	{
		m_MessageQueueListenerLOG1Ptr.CreateInstance(__uuidof(MESSAGEQUEUELISTENERLOGLib::MessageQueueListenerLOG1));
	}
	catch(...)
	{
	}
	return CComObjectRootEx<CComSingleThreadModel>::FinalConstruct();
}

void CMessageQueueListenerMTS1::FinalRelease()
{
	try
	{
		m_MessageQueueListenerLOG1Ptr = NULL;
	}
	catch(...)
	{
	}
	CComObjectRootEx<CComSingleThreadModel>::FinalRelease();
}

HRESULT CMessageQueueListenerMTS1::Activate()
{
	HRESULT hr = GetObjectContext(&m_spObjectContext);
	if (SUCCEEDED(hr))
		return S_OK;
	return hr;
} 

BOOL CMessageQueueListenerMTS1::CanBePooled()
{
	return FALSE;
} 

void CMessageQueueListenerMTS1::Deactivate()
{
	m_spObjectContext.Release();
} 


STDMETHODIMP CMessageQueueListenerMTS1::get_IsInsideMTS(DWORD dwOriginatingThreadId, BOOL *pVal)
{
	if (m_spObjectContext)
	{
		*pVal = TRUE;
		m_spObjectContext->SetComplete();
	}
	else
	{
		*pVal = FALSE;
	}
	return S_OK;
}

STDMETHODIMP CMessageQueueListenerMTS1::get_IsInMTSTransaction(DWORD dwOriginatingThreadId, BOOL *pVal)
{
	if (m_spObjectContext && m_spObjectContext->IsInTransaction())
	{
		*pVal = TRUE;
		m_spObjectContext->SetComplete();
	}
	else
	{
		*pVal = FALSE;
	}
	return S_OK;
}

///////////////////////////////////////////////////////////////////////////////

void CMessageQueueListenerMTS1::AddToErrorMessage(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		// and append it to the list
		_tcscat(m_szErrorMessage, _T("("));
		_tcscat(m_szErrorMessage, pszFunctionName);
		_tcscat(m_szErrorMessage, _T(") "));
		_tcscat(m_szErrorMessage, chMsg);
		_tcscat(m_szErrorMessage, _T("    "));
	}
	catch(...)
	{
	}
}

void CMessageQueueListenerMTS1::AddToErrorMessage(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...)
{
	try
	{
		// expand the current message
		TCHAR    chMsg[MAXLOGMESSAGESIZE];
		va_list pArg;

		va_start(pArg, pszFormat);
		VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
		va_end(pArg);

		// add the comm error message
		_tcscat(m_szErrorMessage, _T("("));
		_tcscat(m_szErrorMessage, pszFunctionName);
		_tcscat(m_szErrorMessage, _T(") "));
		_tcscat(m_szErrorMessage, chMsg);
		VERIFY(_sntprintf(chMsg, MAXLOGMESSAGESIZE - 1, _T(" (HResult = %d, Description = '%s')"), comerr.Error(), (LPCTSTR)comerr.Description()) > -1);
		_tcscat(m_szErrorMessage, chMsg);
		_tcscat(m_szErrorMessage, _T("    "));
	}
	catch(...)
	{
	}
}

