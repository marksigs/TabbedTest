///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUECOMPONENTVC1_H_
#define __MESSAGEQUEUECOMPONENTVC1_H_

#include "resource.h"       // main symbols
#include <mtx.h>
#include "log.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC1
class ATL_NO_VTABLE CMessageQueueComponentVC1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMessageQueueComponentVC1, &CLSID_MessageQueueComponentVC1>,
	public IObjectControl,
	public IDispatchImpl<IMessageQueueComponentVC1, &IID_IMessageQueueComponentVC1, &LIBID_MESSAGEQUEUECOMPONENTVCLib>,
	public CLog
{
public:
	CMessageQueueComponentVC1()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUECOMPONENTVC1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

DECLARE_NOT_AGGREGATABLE(CMessageQueueComponentVC1)

BEGIN_COM_MAP(CMessageQueueComponentVC1)
	COM_INTERFACE_ENTRY(IMessageQueueComponentVC1)
	COM_INTERFACE_ENTRY(IObjectControl)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// IObjectControl
public:
	STDMETHOD(Activate)();
	STDMETHOD_(BOOL, CanBePooled)();
	STDMETHOD_(void, Deactivate)();

	CComPtr<IObjectContext> m_spObjectContext;

// IMessageQueueComponentVC1
public:
	STDMETHOD(OnMessage)(/*[in]*/ BSTR in_xml, /*[out,retval]*/ long* plMESSQ_RESP);
};

#endif //__MESSAGEQUEUECOMPONENTVC1_H_
