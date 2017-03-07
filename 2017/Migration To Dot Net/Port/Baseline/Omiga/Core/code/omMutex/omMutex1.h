/* <VERSION  CORELABEL="" LABEL="R15" DATE="10/10/2003 15:52:58" VERSION="254" PATH="$/CodeCoreCust/4UATCust/Code/Synchronisation/omMutex/omMutex1.h"/> */
///////////////////////////////////////////////////////////////////////////////
//	FILE:			omMutex1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      13/03/03    Initial version
///////////////////////////////////////////////////////////////////////////////

#ifndef __OMMUTEX1_H_
#define __OMMUTEX1_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// ComMutex1
class ATL_NO_VTABLE ComMutex1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<ComMutex1, &CLSID_omMutex1>,
	public ISupportErrorInfo,
	public IDispatchImpl<IomMutex1, &IID_IomMutex1, &LIBID_OMMUTEXLib>
{
public:
	ComMutex1() :
	  m_pMutex(NULL),
	  m_pSingleLock(NULL)
	{
	}

	HRESULT FinalConstruct();
	void FinalRelease();

DECLARE_REGISTRY_RESOURCEID(IDR_OMMUTEX1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(ComMutex1)
	COM_INTERFACE_ENTRY(IomMutex1)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IomMutex1
public:
	STDMETHOD(ReleaseMutex)();
	STDMETHOD(AcquireMutex)(/*[in]*/ BSTR bstrMutexName, /*[in]*/ VARIANT_BOOL bHighPrioirty);
	STDMETHOD(AcquireMutexWithTimeout)(/*[in]*/ BSTR bstrMutexName, /*[in]*/ VARIANT_BOOL bHighPrioirty, /*[in]*/ LONG dwMilliseconds);

private:
	namespaceMutex::CMutex* m_pMutex;
	namespaceMutex::CSingleLock* m_pSingleLock;
};

#endif //__OMMUTEX1_H_
