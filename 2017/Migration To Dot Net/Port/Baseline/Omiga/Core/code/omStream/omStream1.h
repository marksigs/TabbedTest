/* <VERSION  CORELABEL="" LABEL="020.02.06.19.00" DATE="19/06/2002 18:51:20" VERSION="384" PATH="$/CodeCore/Code/FileVersioningSystem/omStream/omStream1.h"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			omStream1.h
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

#ifndef __OMSTREAM1_H_
#define __OMSTREAM1_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// ComStream1
class ATL_NO_VTABLE ComStream1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<ComStream1, &CLSID_omStream1>,
	public ISupportErrorInfo,
	public IDispatchImpl<IomStream1, &IID_IomStream1, &LIBID_OMSTREAMLib>
{
public:
	ComStream1()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_OMSTREAM1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(ComStream1)
	COM_INTERFACE_ENTRY(IomStream1)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IomStream1
public:
	STDMETHOD(Base64ToObject)(/*[in]*/ IUnknown* pUnknownXMLNode, /*[out, retval]*/ IUnknown** ppUnknownObject);
	STDMETHOD(ObjectToBase64)(/*[in]*/ IUnknown* pUnknownObject, /*[in]*/ IUnknown* pUnknownXMLNode);
};

#endif //__OMSTREAM1_H_
