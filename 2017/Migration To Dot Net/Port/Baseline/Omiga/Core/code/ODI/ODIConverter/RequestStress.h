// RequestStress.h: interface for the CRequestStress class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUESTSTRESS_H__66CD53A6_64EC_4A72_B371_AE0872B0D95A__INCLUDED_)
#define AFX_REQUESTSTRESS_H__66CD53A6_64EC_4A72_B371_AE0872B0D95A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Request.h"

class CRequestStress : public CRequest  
{
public:
	CRequestStress(LPCWSTR szType);
	virtual ~CRequestStress();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);

protected:
	static DWORD WINAPI RequestThread(LPVOID lpParameter);

private:
	_bstr_t		m_bstrRequest;
	int			m_nMaxThreads;
	int			m_nSleep;
	static BOOL s_bStressing;
};

#endif // !defined(AFX_REQUESTSTRESS_H__66CD53A6_64EC_4A72_B371_AE0872B0D95A__INCLUDED_)
