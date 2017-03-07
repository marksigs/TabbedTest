///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerLog.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD		03/12/01	SYS3303 - Add support for logging to file
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUELISTENERLOG1_H_
#define __MESSAGEQUEUELISTENERLOG1_H_

#include "resource.h"       // main symbols
#include <comdef.h>
#include <fstream>
#include "ThreadPoolManagerFunctionQueue.h"

///////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerLOG1
class ATL_NO_VTABLE CMessageQueueListenerLOG1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMessageQueueListenerLOG1, &CLSID_MessageQueueListenerLOG1>,
	public ISupportErrorInfo,
	public IDispatchImpl<IMessageQueueListenerLOG1, &IID_IMessageQueueListenerLOG1, &LIBID_MESSAGEQUEUELISTENERLOGLib>
{
public:
	CMessageQueueListenerLOG1();
	virtual ~CMessageQueueListenerLOG1();

	HRESULT FinalConstruct();
	void FinalRelease();

// MSDN Q244495 HOWTO: Implement Thread-Pooled, Apartment Model COM Server in ATL
DECLARE_CLASSFACTORY_AUTO_THREAD() 
DECLARE_NOT_AGGREGATABLE(CMessageQueueListenerLOG1) 

	
DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUELISTENERLOG1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CMessageQueueListenerLOG1)
	COM_INTERFACE_ENTRY(IMessageQueueListenerLOG1)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IMessageQueueListenerLOG1
public:
	STDMETHOD(OnLog)(/*[in]*/ long lLOGAREA, /*[in]*/ BSTR bstr);


// Thread Pool Support (shared by all instances)
public:
	class CThreadData : public CObject
	{
	public:
		CThreadData(_bstr_t bstr) {m_bstr = bstr;}
		_bstr_t m_bstr;
	};
	
	static bool StartUp();
	static void CloseDown();
	static CThreadPoolManagerFunctionQueue* s_pThreadPoolManagerFunctionQueue;
	static int ExecutionRoutine(void*);
	static _bstr_t GetLogFileName();
	static std::ofstream s_LogFile;
};

/////////////////////////////////////////////////////////////////////////////

#endif //__MESSAGEQUEUELISTENERLOG1_H_
