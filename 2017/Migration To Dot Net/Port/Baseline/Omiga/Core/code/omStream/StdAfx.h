/* <VERSION  CORELABEL="" LABEL="020.02.06.19.00" DATE="19/06/2002 18:51:20" VERSION="384" PATH="$/CodeCore/Code/FileVersioningSystem/omStream/StdAfx.h"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			StdAfx.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      13/02/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_STDAFX_H__E6BAA594_FE9F_11D4_82BE_005004E8D1A7__INCLUDED_)
#define AFX_STDAFX_H__E6BAA594_FE9F_11D4_82BE_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__E6BAA594_FE9F_11D4_82BE_005004E8D1A7__INCLUDED)
