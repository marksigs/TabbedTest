// CodePages.h: interface for the CCodePages class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CODEPAGES_H__B05E330D_185B_40EF_A5D9_5AE084F250DE__INCLUDED_)
#define AFX_CODEPAGES_H__B05E330D_185B_40EF_A5D9_5AE084F250DE__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>

class CCodePage;

class CCodePages  
{
public:
	CCodePages();
	virtual ~CCodePages();
	long Init(LPCWSTR szCodePagesPath);
	void Save(LPCWSTR szCodePagesPath);
	long InitNode(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode, LPCWSTR szCodePagesPath = NULL);
	bool GetEnableEuro() const { return m_bEnableEuro; }

	typedef std::map<_bstr_t, CCodePage*, Nocase> MapLocaleToCodePageType;

	MapLocaleToCodePageType& GetMapLocaleToCodePage()
	{
		return m_MapLocaleToCodePage;
	}
	_bstr_t GetLocaleKey(LPCWSTR szCountry, LPCWSTR szLanguage);
	_bstr_t GetLocaleKey(LPCWSTR szCountry, LPCWSTR szLanguage, bool bEuro);
	CCodePage* GetCodePage(LPCWSTR szCountry = L"", LPCWSTR szLanguage = L"");

private:
	enum EOp { opNull, opInsert, opUpdate, opDelete };
	MapLocaleToCodePageType				m_MapLocaleToCodePage;
	bool								m_bEnableEuro;
	bool								m_bFoundDefault;
	_bstr_t								m_bstrCodePagesPath;
};

#endif // !defined(AFX_CODEPAGES_H__B05E330D_185B_40EF_A5D9_5AE084F250DE__INCLUDED_)
