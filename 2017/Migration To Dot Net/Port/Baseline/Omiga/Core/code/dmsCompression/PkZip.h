// PkZip.h : Declaration of the CPkZip

#ifndef __PKZIP_H_
#define __PKZIP_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CPkZip
class ATL_NO_VTABLE CPkZip : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CPkZip, &CLSID_PkZip>,
	public ISupportErrorInfo,
	public IDispatchImpl<IPkZip, &IID_IPkZip, &LIBID_DMSCOMPRESSIONLib>
{
public:
	CPkZip()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_PKZIP)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CPkZip)
	COM_INTERFACE_ENTRY(IPkZip)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IPkZip
public:
	STDMETHOD(ZipFile)(BSTR bstrZipFile, BSTR bstrFile, BOOL noPaths);
	STDMETHOD(UnzipFile)(BSTR bstrZipFile, BSTR bstrFile, BSTR bstrUnzipDir);
	STDMETHOD(UnzipFileToSafeArray)(BSTR bstrZipFile, BSTR bstrFile, VARIANT* pvarUnzippedFiles, LONG* pdwUnzippedFilesCount);
};

#endif //__PKZIP_H_
