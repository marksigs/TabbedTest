///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerOMMQ1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//	LD		06/04/01	SYS2248 - Add performance counters
//	LD		10/09/01	SYS2705 - Support for SQL Server added
//	LD		12/09/01    SYS2707 - Serialize access to public methods
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//  LD		20/06/02	SYS4933 - Create queues on a new thread
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//  RF		24/02/04	BMIDS727 - Different queues can use different tables
///////////////////////////////////////////////////////////////////////////////

#ifndef __THREADPOOLMANAGEROMMQ1_H_
#define __THREADPOOLMANAGEROMMQ1_H_

#include "resource.h"       // main symbols

#include "MessageQueueListener.h"
#include "ScheduleManager.h"
#include "ThreadPoolManagerMQ.h"

#pragma warning(disable: 4146) // unary minus operator applied to unsigned type, result still unsigned
#import "msado15.dll" no_namespace rename("EOF", "ADO_EOF")
#pragma warning(default: 4146)

class CPrfInstanceQueueInfo;

/////////////////////////////////////////////////////////////////////////////
// CThreadPoolManagerOMMQ1
class ATL_NO_VTABLE CThreadPoolManagerOMMQ1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CThreadPoolManagerOMMQ1, &CLSID_ThreadPoolManagerOMMQ1>,
	public IDispatchImpl<IInternalThreadPoolManagerCommon1, &IID_IInternalThreadPoolManagerCommon1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public IDispatchImpl<IInternalThreadPoolManagerOMMQ1, &IID_IInternalThreadPoolManagerOMMQ1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public CThreadPoolManagerMQ
{
public:
	CThreadPoolManagerOMMQ1();
	virtual ~CThreadPoolManagerOMMQ1();

	HRESULT FinalConstruct();
	void FinalRelease();

DECLARE_REGISTRY_RESOURCEID(IDR_THREADPOOLMANAGEROMMQ1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CThreadPoolManagerOMMQ1)
	COM_INTERFACE_ENTRY(IInternalThreadPoolManagerCommon1)
	COM_INTERFACE_ENTRY(IInternalThreadPoolManagerOMMQ1)
	COM_INTERFACE_ENTRY2(IDispatch, IInternalThreadPoolManagerOMMQ1)
END_COM_MAP()

// IThreadPoolManagerCommon1
public:
	STDMETHOD(get_Started)(/*[out, retval]*/ BOOL *pVal);
	STDMETHOD(put_Started)(/*[in]*/ BOOL newVal);
	STDMETHOD(get_NumberOfThreads)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_NumberOfThreads)(/*[in]*/ long newVal);
	STDMETHOD(get_QueueName)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_QueueName)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_QueueStalled)(/*[out, retval]*/ BOOL *pVal);
	STDMETHOD(put_QueueStalled)(/*[in]*/ BOOL newVal);
	STDMETHOD(get_ComponentsStalled)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(put_AddStalledComponents)(/*[in]*/ VARIANT newVal);
	STDMETHOD(get_QueueType)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_dwNextWaitIntervalms)(/*[in]*/ DWORD dwPollInterval);
	STDMETHOD(put_ConnectionString)(/*[in]*/ BSTR bstrConnString);
	STDMETHOD(put_RestartComponents)(/*[in]*/ VARIANT newVal);
	STDMETHOD(get_TableSuffix)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_TableSuffix)(/*[in]*/ BSTR newVal);

	// IThreadPoolManagerOMMQ1
public:
	STDMETHOD(get_ConnectionString)(/*[out, retval]*/ BSTR *pVal);
//	STDMETHOD(put_ConnectionString)(/*[in]*/ BSTR newVal);

// CThreadPoolManager
	virtual bool StartUp();
	virtual void CloseDown();

private:	
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread); // returns NULL if there is no more work to be done 

// Poller
private:
	class CPoller : public CScheduleManager
	{
	public:
		CPoller(CThreadPoolManagerOMMQ1& ThreadPoolManagerOMMQ1) : 
			m_dwNextWaitIntervalms(1000),
			m_ThreadPoolManagerOMMQ1(ThreadPoolManagerOMMQ1)
		{
			SetRepeatCount(1); // single shot scheduling (must keep re-enabling once fired)
		}

		void put_dwNextWaitIntervalms(DWORD dwWaitIntervalms) {m_dwNextWaitIntervalms = dwWaitIntervalms;}

		virtual bool StartUp();
		virtual void CloseDown();

	private:
		virtual DWORD GetdwNextWaitIntervalms() {return m_dwNextWaitIntervalms;}
		virtual void OnEventSchedule();
		
		DWORD m_dwNextWaitIntervalms;
		CThreadPoolManagerOMMQ1& m_ThreadPoolManagerOMMQ1;
	} m_Poller;
	friend class CPoller;

private:
	// called from either the poller at fixed intervals or when a thread is asking for more work
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread, bool bDispatchToAvailableThreads);

private:
	class CThreadData : public CObject
	{
	public:
		CThreadData(
			CThreadPoolManagerOMMQ1& rThreadPoolManagerOMMQ1, 
			_bstr_t bstrConfig, 
			long lProvider, 
			_bstr_t bstrConnectionString, 
			_bstr_t bstrQueueName, 
			_bstr_t bstrTableSuffix,	// RF 30/03/04
			_variant_t vMessageId, 
			_bstr_t bstrProgID, 
			CPrfInstanceQueueInfo* pPrfInstanceQueueInfo) :
			m_rThreadPoolManagerOMMQ1(rThreadPoolManagerOMMQ1),	
			m_bstrConfig(bstrConfig),
			m_lProvider(lProvider),
			m_bstrConnectionString(bstrConnectionString),
			m_bstrQueueName(bstrQueueName), 
			m_bstrTableSuffix(bstrTableSuffix),
			m_vMessageId(vMessageId),
			m_bstrProgID(bstrProgID),
			m_pPrfInstanceQueueInfo(pPrfInstanceQueueInfo)
		{
		}
		~CThreadData() {}
	public:
		CThreadPoolManagerOMMQ1& m_rThreadPoolManagerOMMQ1;
		_bstr_t m_bstrConfig;
		long m_lProvider;
		_bstr_t m_bstrConnectionString;
		_bstr_t m_bstrQueueName;
		_bstr_t m_bstrTableSuffix;
		_variant_t m_vMessageId;
		_bstr_t m_bstrProgID;
		CPrfInstanceQueueInfo* m_pPrfInstanceQueueInfo;
	};
	static void ReceiveExecute(CThreadData* pThreadData);
	
private:
	_bstr_t m_bstrConnectionString;
	_bstr_t m_bstrLockedBy;

	namespaceMutex::CCriticalSection m_csPublicMethod;

	// support for create ADO objects on a new thread
	static DWORD WINAPI StartUponThread(void* pThreadPoolManagerOMMQ1);
	DWORD WINAPI StartUponThread();
	DWORD m_dwThreadId;// thread id on which ADO objects are created
	HANDLE m_hThread; // thread handle on which ADO objects are created
	HANDLE m_hEventStartUp; // event used to signal the ADO objects have been created or an error has occurred
	HANDLE m_hEventCloseDown; // event used to signal the ADO objects have been destroyed or an error has occurred

	_bstr_t m_bstrQueueName;		
	bstr_t m_bstrTableSuffix;

    namespaceMutex::CSharedLockSmallOperations m_sharedlockResetQueue;
	BOOL m_bQueueStalled;
	BOOL m_bQueueStarted;

	void CloseCommand(_CommandPtr& CommandPtr);
	void CloseRecordset(_RecordsetPtr& RecordsetPtr);
	void CloseConnection(_ConnectionPtr& ConnectionPtr);
	void LogEventErrorProvider(_ConnectionPtr& ConnectionPtr);
	void UnlockMessage(_variant_t vMessageId);

	bool UnlockMessages();
	bool ResetQueue();

	// Logging support
	void LogEventInfo(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventWarning(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventWarning(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...);
	void LogEventError(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventError(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...);

	// Performance counter support
	CPrfInstanceQueueInfo* m_pPrfInstanceQueueInfo;

	_bstr_t MessageIdToBstrt(_variant_t vId);

private:
	long m_lProvider; // defined in MessageQueueListenerMTS.idl
	volatile bool m_bDBConnectivityGood;
	static UINT WM_OMMQ1_OPENKEEPALIVE;
	volatile bool m_bProcessingOpenKeepAlive;
	void OnBadConnection();

	static _bstr_t s_bstrNULL;
	static _variant_t s_vZero;
	static _variant_t s_vOne;
};

///////////////////////////////////////////////////////////////////////////////

#endif //__THREADPOOLMANAGEROMMQ1_H_
