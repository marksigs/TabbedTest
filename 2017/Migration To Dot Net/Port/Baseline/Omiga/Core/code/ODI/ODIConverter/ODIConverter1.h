///////////////////////////////////////////////////////////////////////////////
//	FILE:			ODIConverter1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      10/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#ifndef __ODICONVERTER1_H_
#define __ODICONVERTER1_H_

#include "resource.h"       // main symbols

class CException;
class exception;

/////////////////////////////////////////////////////////////////////////////
// CODIConverter1
class ATL_NO_VTABLE CODIConverter1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CODIConverter1, &CLSID_ODIConverter>,
	public ISupportErrorInfo,
	public IDispatchImpl<IODIConverter, &IID_IODIConverter, &LIBID_ODICONVERTER>
{
public:
DECLARE_REGISTRY_RESOURCEID(IDR_ODICONVERTER1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CODIConverter1)
	COM_INTERFACE_ENTRY(IODIConverter)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IODIConverter
public:
	STDMETHOD(Request)(/*[in]*/ BSTR bstrRequest, /*[out, retval]*/ BSTR* pbstrResponse);
	static _bstr_t Request(LPCWSTR szRequest);
	static HRESULT LogError(CException& e);

protected:
	static MSXML::IXMLDOMNodePtr DoRequest(MSXML::IXMLDOMNodePtr ptrRequestNode);

public:
// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
DECLARE_CLASSFACTORY_AUTO_THREAD() 
//DECLARE_NOT_AGGREGATABLE_DBG(CODIConverter1) 
DECLARE_NOT_AGGREGATABLE(CODIConverter1) 
};

#endif //__ODICONVERTER1_H_
