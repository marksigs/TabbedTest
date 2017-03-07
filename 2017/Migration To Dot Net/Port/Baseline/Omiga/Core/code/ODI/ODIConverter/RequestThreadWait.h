// RequestThreadWait.h: interface for the CRequestThreadWait class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUESTTHREADWAIT_H__B27BEA4B_9D20_48DB_9D92_948F0BD483B7__INCLUDED_)
#define AFX_REQUESTTHREADWAIT_H__B27BEA4B_9D20_48DB_9D92_948F0BD483B7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Request.h"

class CRequestThreadWait : public CRequest  
{
public:
	CRequestThreadWait(LPCWSTR szType);
	virtual ~CRequestThreadWait();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);
};

#endif // !defined(AFX_REQUESTTHREADWAIT_H__B27BEA4B_9D20_48DB_9D92_948F0BD483B7__INCLUDED_)
