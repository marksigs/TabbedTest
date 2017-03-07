///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolManagerMSMQ1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//  LD		19/03/01	SYS2103 - Make MSMQ1 queue tolerant to database shutdown/startup.
//	LD		09/04/01	SYS2249 - Corrections for two listeners on the same queue
//	LD		12/09/01    SYS2707 - Serialize access to public methods
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		07/12/01	SYS3421 - Create MSMQ objects on a new thread
//	LD		21/03/02	SYS4414 - Increase peek time out and enhanced logging.
//								- m_lTotalMessagesSkippedBecauseStalled introduced to manage messages to scan
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//	LD		19/03/03	Only access MSMQ objects from main creator thread
//  RF		24/02/04	Allow for multiple queue tables
///////////////////////////////////////////////////////////////////////////////

#ifndef __THREADPOOLMANAGERMSMQ1_H_
#define __THREADPOOLMANAGERMSMQ1_H_

#include "resource.h"       // main symbols

#include "MessageQueueListener.h"
#include "ThreadPoolManagerMQ.h"
#import "mqoa.dll" no_namespace

class CPrfInstanceQueueInfo;

/////////////////////////////////////////////////////////////////////////////

class ATL_NO_VTABLE CThreadPoolManagerMSMQ1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CThreadPoolManagerMSMQ1, &CLSID_ThreadPoolManagerMSMQ1>,
	public IDispatchImpl<IInternalThreadPoolManagerCommon1, &IID_IInternalThreadPoolManagerCommon1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public IDispatchImpl<IInternalThreadPoolManagerMSMQ1, &IID_IInternalThreadPoolManagerMSMQ1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public IDispatchImpl<IInternalMSMQEventSink, &IID_IInternalMSMQEventSink, &LIBID_MESSAGEQUEUELISTENERLib>,
	public CThreadPoolManagerMQ
{
public:
	CThreadPoolManagerMSMQ1();
	virtual ~CThreadPoolManagerMSMQ1();

	HRESULT FinalConstruct();
	void FinalRelease();

DECLARE_REGISTRY_RESOURCEID(IDR_THREADPOOLMANAGERMSMQ1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CThreadPoolManagerMSMQ1)
	COM_INTERFACE_ENTRY_IID(__uuidof(_DMSMQEventEvents), IInternalMSMQEventSink)
	COM_INTERFACE_ENTRY(IInternalThreadPoolManagerCommon1)
	COM_INTERFACE_ENTRY(IInternalThreadPoolManagerMSMQ1)
	COM_INTERFACE_ENTRY2(IDispatch, IInternalMSMQEventSink)
END_COM_MAP()

// IMSMQEventSink
public:
    // Notification of MSMQEvent.ArrivedError
	STDMETHOD(ArrivedError)(IDispatch* pQueue, long lErrorCode, long lCursor);
    // Notification of MSMQEvent.Arrived
	STDMETHOD(Arrived)(IDispatch* pQueue, long lCursor);

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

	// IThreadPoolManagerMSMQ1
public:

// CThreadPoolManager
public:
	virtual bool CreatorThreadStartUp();
	virtual void CreatorThreadCloseDown();

// Internal functions
protected:
	bool CreatorThreadStartUpThreadPool();
	bool CreatorThreadStartUpMSMQ(bool bReportErrors = true);
	void CreatorThreadCloseDownThreadPool();
	void CreatorThreadCloseDownMSMQ();
	void CreatorThreadEnableEvents(long lCursor);

private:
	namespaceMutex::CCriticalSection m_csPublicMethod;
	
	// support for create MSMQ objects on a new thread
	static DWORD WINAPI StartUponThread(void* pThreadPoolManagerMSMQ1);
	DWORD WINAPI StartUponThread();
	DWORD m_dwCreatorThreadId;// thread id on which MSMQ objects are created
	HANDLE m_hCreatorThread; // thread handle on which MSMQ objects are created
	HANDLE m_hEventCreatedOrError; // event used to signal the MSMQ objects have been created or an error has occurred
	HANDLE m_hEventCloseDown; // event used to signal the MSMQ objects have been destroyed

	// the peeking objects - use to schedule message ids to threads
	IMSMQQueuePtr m_MSMQQueuePtr;
	IMSMQEventPtr m_MSMQEventPtr;
	IDispatchPtr  m_DispatchPtr;
	DWORD m_dwCookie;

	enum TePeek
	{
		ePeekCurrentToDo,
		ePeekCurrentInProgress,
		ePeekCurrentDone
	} m_ePeek;
	
	BOOL m_bNotificationEnabled;
	BOOL m_bQueueStalled;
	volatile BOOL m_bQueueStarted;

	bstr_t m_bstrQueueName; // as specified in the configuration file
	bstr_t m_bstrTableSuffix;
	bstr_t m_bstrFormatName;

	bstr_t m_bstrMoveQueueName; // queue to move messages to
	bstr_t m_bstrMoveFormatName;

	static UINT WM_MSMQ1_THREADAVAILABLEFORREUSE;
	static UINT WM_MSMQ1_CLOSEDOWN;
	static UINT WM_MSMQ1_RESET;
	static UINT WM_MSMQ1_RESTART;

	// for use by the thread which creates the MSMQ objects
	virtual void CreatorThreadDispatchMessageToWorkerThread();
	virtual CThreadPoolMessage* CreatorThreadRemoveMessageFromQueue(); // returns NULL if there is no more work to be done 
	volatile BOOL m_bThreadAvailablePosted;

	// for use by threads in the thread-pool
	virtual CThreadPoolMessage* RemoveMessageFromQueue(CThreadPoolThread* pThreadPoolThread); // returns NULL if there is no more work to be done 
	class CThreadData : public CObject
	{
	public:
		CThreadData(CThreadPoolManagerMSMQ1& rThreadPoolManagerMSMQ1, _bstr_t bstrConfig, _bstr_t bstrQueueName, _bstr_t bstrFormatName, _bstr_t bstrMoveQueueName, _bstr_t bstrMoveFormatName, _variant_t vIdMessage, _bstr_t bstrProgID, long lMessagesToScan, CPrfInstanceQueueInfo* pPrfInstanceQueueInfo) :
			m_rThreadPoolManagerMSMQ1(rThreadPoolManagerMSMQ1),
			m_bstrConfig(bstrConfig), 
			m_bstrQueueName(bstrQueueName), 
			m_bstrFormatName(bstrFormatName), 
			m_bstrMoveQueueName(bstrMoveQueueName), 
			m_bstrMoveFormatName(bstrMoveFormatName), 
			m_vIdMessage(vIdMessage),
			m_bstrProgID(bstrProgID),
			m_lMessagesToScan(lMessagesToScan),
			m_pPrfInstanceQueueInfo(pPrfInstanceQueueInfo) {}
		~CThreadData() {}
	public:
		CThreadPoolManagerMSMQ1& m_rThreadPoolManagerMSMQ1;
		_bstr_t m_bstrConfig;
		_bstr_t m_bstrQueueName;
		_bstr_t m_bstrFormatName;
		_bstr_t m_bstrMoveQueueName;
		_bstr_t m_bstrMoveFormatName;
		_bstr_t m_bstrTableSuffix;
		_variant_t m_vIdMessage;
		_bstr_t m_bstrProgID;
		long m_lMessagesToScan;
		CPrfInstanceQueueInfo* m_pPrfInstanceQueueInfo;

	};
	static void ReceiveExecute(CThreadData* pThreadData);

	// Logging support
	void LogEventInfo(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventWarning(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventError(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void LogEventError(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...);

	// MSMQ Id support
	_bstr_t MSMQIdToBstrt(_variant_t vId);

	// Performance counter support
	CPrfInstanceQueueInfo* m_pPrfInstanceQueueInfo;

	long m_lTotalMessagesSkippedBecauseStalled;
	long m_lTotalMessagesRolledback;
};

#endif //__THREADPOOLMANAGERMSMQ1_H_
