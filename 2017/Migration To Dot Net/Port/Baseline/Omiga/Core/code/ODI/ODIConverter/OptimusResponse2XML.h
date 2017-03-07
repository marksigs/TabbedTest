///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusResponse2XML.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      02/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_OPTIMUSRESPONSE2XML_H__214B4F51_A5C4_422F_BBF2_64488E2927DD__INCLUDED_)
#define AFX_OPTIMUSRESPONSE2XML_H__214B4F51_A5C4_422F_BBF2_64488E2927DD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CMetaDataEnvOptimus;
class CCodePage;
class CProfiler;

class COptimusResponse2XML  
{
public:
	static MSXML::IXMLDOMNodePtr Convert(
		BYTE* pbResponse, 
		size_t nSize, 
		CMetaDataEnvOptimus* pMetaDataEnv,
		CCodePage* pCodePage);

private:
	static short GetShort(BYTE*& pbCurrent)
	{
		short nShort = MAKEWORD(pbCurrent[1], pbCurrent[0]);
		pbCurrent += sizeof(short);
		return nShort;
	}
	static long GetLong(BYTE*& pbCurrent)
	{
		long llong = MAKELONG(MAKEWORD(pbCurrent[3], pbCurrent[2]), MAKEWORD(pbCurrent[1], pbCurrent[0]));
		pbCurrent += sizeof(long);
		return llong;
	}
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_OPTIMUSRESPONSE2XML_H__214B4F51_A5C4_422F_BBF2_64488E2927DD__INCLUDED_)
