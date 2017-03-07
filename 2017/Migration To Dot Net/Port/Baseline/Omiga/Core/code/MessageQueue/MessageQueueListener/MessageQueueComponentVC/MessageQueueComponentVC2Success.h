///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueComponentVC2Success.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      10/04/01    Initial version
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUECOMPONENTVC2SUCCESS_H_
#define __MESSAGEQUEUECOMPONENTVC2SUCCESS_H_

#include "resource.h"       // main symbols
#include <mtx.h>
#include "log.h"

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueComponentVC2Success
class ATL_NO_VTABLE CMessageQueueComponentVC2Success : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMessageQueueComponentVC2Success, &CLSID_MessageQueueComponentVC2Success>,
	public IObjectControl,
	public IDispatchImpl<IMessageQueueComponentVC2, &IID_IMessageQueueComponentVC2, &LIBID_MESSAGEQUEUECOMPONENTVCLib>,
	public CLog
{
public:
	CMessageQueueComponentVC2Success()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUECOMPONENTVC2SUCCESS)

DECLARE_PROTECT_FINAL_CONSTRUCT()

DECLARE_NOT_AGGREGATABLE(CMessageQueueComponentVC2Success)

BEGIN_COM_MAP(CMessageQueueComponentVC2Success)
	COM_INTERFACE_ENTRY(IMessageQueueComponentVC2)
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

///////////////////////////////////////////////////////////////////////////////

#endif //__MESSAGEQUEUECOMPONENTVC2SUCCESS_H_
