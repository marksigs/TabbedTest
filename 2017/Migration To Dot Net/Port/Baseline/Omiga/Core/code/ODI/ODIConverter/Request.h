// Request.h: interface for the CRequest class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUEST_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_)
#define AFX_REQUEST_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CRequest  
{
public:
	CRequest(LPCWSTR szType);
	virtual ~CRequest();
	static CRequest* Create(LPCWSTR szType);
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode) = 0;
	_bstr_t GetName() const { return m_bstrName; }
	void SetName(LPCWSTR szName) { _ASSERTE(szName != NULL); _ASSERTE(wcslen(szName) > 0); m_bstrName = szName; }
	static void WriteBuffer(LPCTSTR lpszFName, LPBYTE lpBuffer, int nBufferSize);
	static void WriteXML(LPCTSTR lpszFName, MSXML::IXMLDOMNodePtr ptrNode);
	static void LogProfile(const MSXML::IXMLDOMNodePtr ptrNode);

private:
	_bstr_t		m_bstrName;
};

#define IF_LOG(expr) if (pMetaDataEnv && pMetaDataEnv->GetLogDebug()) expr;

#endif // !defined(AFX_REQUEST_H__9EF44E7F_D8B0_4A63_BC48_5898C043DEC5__INCLUDED_)
