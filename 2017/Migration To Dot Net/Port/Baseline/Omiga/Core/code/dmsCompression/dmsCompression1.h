// dmsCompression1.h : Declaration of the CdmsCompression1

#ifndef __DMSCOMPRESSION1_H_
#define __DMSCOMPRESSION1_H_

#include "resource.h"       // main symbols
#include <comdef.h>

/////////////////////////////////////////////////////////////////////////////
// CdmsCompression1
class ATL_NO_VTABLE CdmsCompression1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CdmsCompression1, &CLSID_dmsCompression1>,
	public ISupportErrorInfo,
	public IDispatchImpl<IdmsCompression1, &IID_IdmsCompression1, &LIBID_DMSCOMPRESSIONLib>
{
public:
	CdmsCompression1();

DECLARE_REGISTRY_RESOURCEID(IDR_DMSCOMPRESSION1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CdmsCompression1)
	COM_INTERFACE_ENTRY(IdmsCompression1)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IdmsCompression1
public:
	STDMETHOD(SafeArrayDecompressToXMLNODEBINHEX)(/*[in]*/ VARIANT varIn, /*[in]*/ IUnknown* pUnknownXMLNode);
	STDMETHOD(XMLNODEBINHEXCompressToSafeArray)(/*[in]*/ IUnknown* pUnknownXMLNode, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(SafeArrayDecompressToXMLNODEBINBASE64)(/*[in]*/ VARIANT varIn, /*[in]*/ IUnknown* pUnknownXMLNode);
	STDMETHOD(XMLNODEBINBASE64CompressToSafeArray)(/*[in]*/ IUnknown* pUnknownXMLNode, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(SafeArrayDecompressToXMLNODE)(/*[in]*/ VARIANT varIn, /*[in]*/ IUnknown* pUnknownXMLNode);
	STDMETHOD(XMLNODECompressToSafeArray)(/*[in]*/ IUnknown* pUnknownXMLNode, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(SafeArrayDecompressToBSTR)(/*[in]*/ VARIANT varIn, /*[out,retval]*/ BSTR* pbstrOut);
	STDMETHOD(BSTRCompressToSafeArray)(/*[in]*/ BSTR bstrIn, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(SafeArrayCompressToSafeArray)(/*[in]*/ VARIANT varIn, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(SafeArrayDecompressToSafeArray)(/*[in]*/ VARIANT varIn, /*[out,retval]*/ VARIANT* pvarOut);
	STDMETHOD(FileCompressToFile)(/*[in]*/ BSTR bstrSrcFile, /*[in]*/ BSTR bstrTgtFile);
	STDMETHOD(FileDecompressToFile)(/*[in]*/ BSTR bstrSrcFile, /*[in]*/ BSTR bstrTgtFile);
	STDMETHOD(get_CompressionMethod)(/*[out, retval]*/ BSTR* pbstrCompressionMethod);
	STDMETHOD(put_CompressionMethod)(/*[in]*/ BSTR bstrCompressionMethod);

protected:
	_bstr_t		m_bstrCompressionMethod;
};

#endif //__DMSCOMPRESSION1_H_
