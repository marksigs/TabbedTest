// RequestDecodeOptimusStream.h: interface for the CRequestDecodeOptimusStream class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_REQUESTDECODEOPTIMUSSTREAM_H__E49EC0A1_2412_4F15_909F_EFAB366498D6__INCLUDED_)
#define AFX_REQUESTDECODEOPTIMUSSTREAM_H__E49EC0A1_2412_4F15_909F_EFAB366498D6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Request.h"

class CRequestDecodeOptimusStream : public CRequest  
{
public:
	CRequestDecodeOptimusStream(LPCWSTR szType);
	virtual ~CRequestDecodeOptimusStream();
	virtual MSXML::IXMLDOMNodePtr Execute(const MSXML::IXMLDOMNodePtr ptrRequestNode);
};

#endif // !defined(AFX_REQUESTDECODEOPTIMUSSTREAM_H__E49EC0A1_2412_4F15_909F_EFAB366498D6__INCLUDED_)
