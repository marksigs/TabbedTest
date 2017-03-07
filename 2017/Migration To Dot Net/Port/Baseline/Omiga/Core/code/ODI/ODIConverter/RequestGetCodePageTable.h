// RequestGetCodePageTable.h: interface for the CRequestGetCodePageTable class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUESTGETCODEPAGETABLE_H__3F91BFD2_B040_4A0D_B641_E1A1F08D0E6E__INCLUDED_)
#define AFX_REQUESTGETCODEPAGETABLE_H__3F91BFD2_B040_4A0D_B641_E1A1F08D0E6E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Request.h"

class CRequestGetCodePageTable : public CRequest  
{
public:
	CRequestGetCodePageTable(LPCWSTR szType);
	virtual ~CRequestGetCodePageTable();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);
};

#endif // !defined(AFX_REQUESTGETCODEPAGETABLE_H__3F91BFD2_B040_4A0D_B641_E1A1F08D0E6E__INCLUDED_)
