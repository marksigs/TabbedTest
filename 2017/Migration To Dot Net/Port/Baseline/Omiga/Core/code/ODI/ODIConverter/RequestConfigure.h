// Request.h: interface for the CRequestConfigure class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUESTCONFIGURE_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_)
#define AFX_REQUESTCONFIGURE_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Request.h"

class CRequestConfigure : public CRequest  
{
public:
	CRequestConfigure(LPCWSTR szType);
	virtual ~CRequestConfigure();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);
};

#endif // !defined(AFX_REQUEST_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_)
