// PCAttributesBO.h : Declaration of the CPCAttributesBO

#ifndef __PCATTRIBUTESBO_H_
#define __PCATTRIBUTESBO_H_

#include "resource.h"       // main symbols
#include <atlctl.h>

/////////////////////////////////////////////////////////////////////////////
// CPCAttributesBO
class ATL_NO_VTABLE CPCAttributesBO : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CPCAttributesBO, &CLSID_PCAttributesBO>,
	public ISupportErrorInfo,
	public IDispatchImpl<IPCAttributesBO, &IID_IPCAttributesBO, &LIBID_OMPCLib>,
	public IObjectSafetyImpl<CPCAttributesBO, INTERFACESAFE_FOR_UNTRUSTED_CALLER>
{
public:
	CPCAttributesBO()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_PCATTRIBUTESBO)
DECLARE_NOT_AGGREGATABLE(CPCAttributesBO)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CPCAttributesBO)
	COM_INTERFACE_ENTRY(IPCAttributesBO)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IObjectSafety)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPCAttributesBO
public:
	STDMETHOD(GetMACAddress)(/*[out, retval]*/ BSTR* pbstrMACAddress);
	STDMETHOD(NameOfPC)(/*[out, retval]*/ BSTR * pbstrPCName);
	STDMETHOD(FindLocalPrinterList)(BSTR bstrXMLRequest, BSTR * pbstrPrinterList);
	STDMETHOD(Sleep)(int dwMilliseconds);
};

#endif //__PCATTRIBUTESBO_H_
