///////////////////////////////////////////////////////////////////////////////
//	FILE:			STDAFX.H
//	DESCRIPTION: 	
//
//	SYSTEM:	    	DMS
//	COPYRIGHT:		(c) 2002, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	LD		23/04/02	First version
////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_STDAFX_H__A6FEB296_CD0D_46F9_BA70_5A16741CBFA3__INCLUDED_)
#define AFX_STDAFX_H__A6FEB296_CD0D_46F9_BA70_5A16741CBFA3__INCLUDED_

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

#endif // !defined(AFX_STDAFX_H__A6FEB296_CD0D_46F9_BA70_5A16741CBFA3__INCLUDED)
