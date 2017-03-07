///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusXML2Request.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      08/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_OPTIMUSXML2REQUEST_H__795D86DF_D7FA_4CE3_A684_D45FF861A2FD__INCLUDED_)
#define AFX_OPTIMUSXML2REQUEST_H__795D86DF_D7FA_4CE3_A684_D45FF861A2FD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CMetaDataEnvOptimus;
class CCodePage;

class COptimusXML2Request  
{
public:
	static size_t Convert(
		const MSXML::IXMLDOMNodePtr ptrRequestNode, 
		BYTE* pbRequest, 
		const size_t nBufferSize, 
		CMetaDataEnvOptimus* pMetaDataEnv,
		CCodePage* pCodePage);

private:
	static size_t ConvertDOMElement(
		const MSXML::IXMLDOMElementPtr ptrXMLDOMElement, 
		BYTE*& pbRequest, 
		const size_t nBufferSize, 
		long lParentNumber, 
		long& lNumbering, 
		short nParentId,
		CMetaDataEnvOptimus* pMetaDataEnv,
		CCodePage* pCodePage);

	static void PutShort(BYTE*& pbCurrent, short nShort)
	{
		pbCurrent[0] = HIBYTE(nShort);
		pbCurrent[1] = LOBYTE(nShort);
		pbCurrent += sizeof(short);
	}
	static void PutLong(BYTE*& pbCurrent, long lLong)
	{
		pbCurrent[0] = HIBYTE(HIWORD(lLong));
		pbCurrent[1] = LOBYTE(HIWORD(lLong));
		pbCurrent[2] = HIBYTE(LOWORD(lLong));
		pbCurrent[3] = LOBYTE(LOWORD(lLong));
		pbCurrent += sizeof(long);
	}
};

#endif // !defined(AFX_OPTIMUSXML2REQUEST_H__795D86DF_D7FA_4CE3_A684_D45FF861A2FD__INCLUDED_)
