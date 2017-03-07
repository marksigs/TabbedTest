///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestTransact.h
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

#if !defined(AFX_REQUESTTRANSACT_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_)
#define AFX_REQUESTTRANSACT_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "mutex.h"
#include "Request.h"

class CMetaDataEnvTransact;

class CRequestTransact : public CRequest
{
public:
	CRequestTransact(LPCWSTR szType);
	virtual ~CRequestTransact();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);

protected:
	MSXML::IXMLDOMNodePtr GetResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvTransact* pMetaDataEnv) const;
	MSXML::IXMLDOMNodePtr GetTransactResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvTransact* pMetaDataEnv) const;
	MSXML::IXMLDOMNodePtr GetTestResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvTransact* pMetaDataEnv) const;
	int GetTransactPort(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvTransact* pMetaDataEnv) const;
	void WriteXML(MSXML::IXMLDOMNodePtr ptrNode, LPCTSTR pszType, int nMaxWriteXML, int nIncrement) const;

    static namespaceMutex::CCriticalSection s_csWriteBufferLastRequest;
    static namespaceMutex::CCriticalSection s_csWriteBufferLastResponse;
    static namespaceMutex::CCriticalSection s_csWriteXML;
};

#endif // !defined(AFX_REQUESTTRANSACT_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_)
