///////////////////////////////////////////////////////////////////////////////
//	FILE:			WSABuf.h
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

#if !defined(AFX_WSABUF_H__05D6767C_1A4E_4116_8BF0_D5757E314537__INCLUDED_)
#define AFX_WSABUF_H__05D6767C_1A4E_4116_8BF0_D5757E314537__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class WSABuf : public WSABUF  
{
public:
	WSABuf();
	virtual ~WSABuf();

};

#endif // !defined(AFX_WSABUF_H__05D6767C_1A4E_4116_8BF0_D5757E314537__INCLUDED_)
