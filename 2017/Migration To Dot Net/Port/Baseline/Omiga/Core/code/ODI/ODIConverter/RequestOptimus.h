///////////////////////////////////////////////////////////////////////////////
//	FILE:			RequestOptimus.h
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

#if !defined(AFX_REQUESTOPTIMUS_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_)
#define AFX_REQUESTOPTIMUS_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "mutex.h"
#include "Request.h"

class CMetaDataEnvOptimus;
class CCodePage;

class CRequestOptimus : public CRequest
{
public:
	CRequestOptimus(LPCWSTR szType);
	virtual ~CRequestOptimus();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);

protected:
	MSXML::IXMLDOMNodePtr GetResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvOptimus* pMetaDataEnv, CCodePage* pCodePage) const;
	MSXML::IXMLDOMNodePtr GetOptimusResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvOptimus* pMetaDataEnv, CCodePage* pCodePage) const;
	MSXML::IXMLDOMNodePtr GetTestResponseNode(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvOptimus* pMetaDataEnv, CCodePage* pCodePage) const;
	int GetOptimusPort(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvOptimus* pMetaDataEnv) const;
	bool UpdateStartNewSessionRequest(const MSXML::IXMLDOMNodePtr ptrRequestNode, CMetaDataEnvOptimus* pMetaDataEnv) const;
	bool UpdateAttribute(
		const MSXML::IXMLDOMNodePtr ptrParentNode,
		LPCWSTR szNodeKey,
		LPCWSTR szAttributeKey,
		LPCWSTR szAttributeValue) const;
	void WriteXML(MSXML::IXMLDOMNodePtr ptrNode, LPCTSTR pszType, int nMaxWriteXML, int nIncrement) const;

    static namespaceMutex::CCriticalSection s_csWriteBufferLastRequest;
    static namespaceMutex::CCriticalSection s_csWriteBufferLastResponse;
    static namespaceMutex::CCriticalSection s_csWriteXML;
};

#endif // !defined(AFX_REQUESTOPTIMUS_H__E0ADE6F3_CEC2_4EE0_B775_C0E16A59D614__INCLUDED_)
