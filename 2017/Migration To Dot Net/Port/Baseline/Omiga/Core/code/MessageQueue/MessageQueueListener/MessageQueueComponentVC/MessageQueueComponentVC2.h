///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC2.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      15/11/00    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUECOMPONENTVC2_H_
#define __MESSAGEQUEUECOMPONENTVC2_H_

#include "resource.h"       // main symbols
#include <mtx.h>
#include "log.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC2
class ATL_NO_VTABLE CMessageQueueComponentVC2 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMessageQueueComponentVC2, &CLSID_MessageQueueComponentVC2>,
	public IObjectControl,
	public IDispatchImpl<IMessageQueueComponentVC2, &IID_IMessageQueueComponentVC2, &LIBID_MESSAGEQUEUECOMPONENTVCLib>,
	public CLog
{
public:
	CMessageQueueComponentVC2()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUECOMPONENTVC2)

DECLARE_PROTECT_FINAL_CONSTRUCT()

DECLARE_NOT_AGGREGATABLE(CMessageQueueComponentVC2)

BEGIN_COM_MAP(CMessageQueueComponentVC2)
	COM_INTERFACE_ENTRY(IMessageQueueComponentVC2)
	COM_INTERFACE_ENTRY(IObjectControl)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// IObjectControl
public:
	STDMETHOD(Activate)();
	STDMETHOD_(BOOL, CanBePooled)();
	STDMETHOD_(void, Deactivate)();

	CComPtr<IObjectContext> m_spObjectContext;

// IMessageQueueComponentVC2
public:
	STDMETHOD(OnMessage)(/*[in]*/ BSTR in_xmlConfig, /*[in]*/ BSTR in_xmlData, /*[out,retval]*/ long* plMESSQ_RESP);
};

#endif //__MESSAGEQUEUECOMPONENTVC2_H_
