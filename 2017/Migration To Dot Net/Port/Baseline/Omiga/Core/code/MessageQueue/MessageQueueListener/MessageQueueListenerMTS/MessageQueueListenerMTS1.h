///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerMTS1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
//	LD		09/04/01	SYS2249 - Corrections for two listeners on the same queue
//	LD		10/09/01	SYS2705 - Support for SQL Server added
//	LD		03/10/01	SYS2770 - Improve error messaging
//  LD		03/12/01	SYS3303 - Add support for logging to file
//	LD		09/02/02	SYS4414	- Add s_vtNull
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
//  LD		19/03/03	Updates to support queuenames specified as a format name
//  RF		24/02/04	BMIDS727 - Allow multiple queue tables
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUELISTENERMTS1_H_
#define __MESSAGEQUEUELISTENERMTS1_H_

#include "resource.h"       // main symbols
#include <mtx.h>
#include "log.h"

#import "mqoa.dll" no_namespace

#pragma warning(disable: 4146) // unary minus operator applied to unsigned type, result still unsigned
#import "msado15.dll" no_namespace rename("EOF", "ADO_EOF")
#pragma warning(default: 4146)

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerMTS1
class ATL_NO_VTABLE CMessageQueueListenerMTS1 : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMessageQueueListenerMTS1, &CLSID_MessageQueueListenerMTS1>,
	public IObjectControl,
	public IDispatchImpl<IInternalMessageQueueListenerMTS1, &IID_IInternalMessageQueueListenerMTS1, &LIBID_MESSAGEQUEUELISTENERMTSLib>,
	public CLog
{
public:
	CMessageQueueListenerMTS1()
	{
	}

	HRESULT FinalConstruct();
	void FinalRelease();


DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUELISTENERMTS1)

DECLARE_PROTECT_FINAL_CONSTRUCT()

DECLARE_NOT_AGGREGATABLE(CMessageQueueListenerMTS1)

BEGIN_COM_MAP(CMessageQueueListenerMTS1)
	COM_INTERFACE_ENTRY(IInternalMessageQueueListenerMTS1)
	COM_INTERFACE_ENTRY(IObjectControl)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// IObjectControl
public:
	STDMETHOD(Activate)();
	STDMETHOD_(BOOL, CanBePooled)();
	STDMETHOD_(void, Deactivate)();

	CComPtr<IObjectContext> m_spObjectContext;

	// IMessageQueueListenerMTS1
public:
	STDMETHOD(get_IsInMTSTransaction)(/*[in]*/ DWORD dwOriginatingThreadId, /*[out, retval]*/ BOOL *pVal);
	STDMETHOD(get_IsInsideMTS)(/*[in]*/ DWORD dwOriginatingThreadId, /*[out, retval]*/ BOOL *pVal);
	STDMETHOD(MSMQReceiveExecute)(/*[in]*/ DWORD dwOriginatingThreadId, /*[in]*/ BSTR bstrConfig, /*[in]*/ BSTR bstrQueueName, /*[in]*/ BSTR bstrFormatName, /*[in]*/ BSTR bstrMoveQueueName, /*[in]*/ BSTR bstrMoveFormatName, /*[in]*/ VARIANT vMsgId, /*[in]*/ long lMessagesToScan, /*[out]*/ BSTR* pbstrErrorMessage, /*[out, retval]*/ long* plMESSQ_RESP);
	STDMETHOD(MSMQMoveMessage)(/*[in]*/ DWORD dwOriginatingThreadId, /*[in]*/ BSTR bstrQueueName, /*[in]*/ BSTR bstrFormatName, /*[in]*/ BSTR bstrMoveQueueName, /*[in]*/ BSTR bstrMoveFormatName, /*[in]*/ VARIANT vMsgId, /*[in]*/ long lMessagesToScan, /*[out]*/ BSTR* pbstrErrorMessage);
	//STDMETHOD(OMMQReceiveExecute)(/*[in]*/ DWORD dwOriginatingThreadId, /*[in]*/ BSTR bstrConfig, /*[in]*/ long lProvider, /*[in]*/ BSTR bstrConnectionString, /*[in]*/ BSTR bstrQueueName, /*[in]*/ VARIANT vMsgId,  /*[out]*/ BSTR* pbstrErrorMessage, /*[out, retval]*/ long* plMESSQ_RESP);
	STDMETHOD(OMMQReceiveExecute)(
		/*[in]*/ DWORD dwOriginatingThreadId, 
		/*[in]*/ BSTR bstrConfig, 
		/*[in]*/ long lProvider, 
		/*[in]*/ BSTR bstrConnectionString, 
		/*[in]*/ BSTR bstrQueueName, 
		/*[in]*/ VARIANT vMsgId,  
		/*[in]*/ BSTR bstrTableSuffix,
		/*[out]*/ BSTR* pbstrErrorMessage, 
		/*[out, retval]*/ long* plMESSQ_RESP);
	STDMETHOD(OMMQMoveMessage)(
		/*[in]*/ DWORD dwOriginatingThreadId, 
		/*[in]*/ long lProvider, 
		/*[in]*/ BSTR bstrConnectionString, 
		/*[in]*/ BSTR bstrQueueName, 
        /*[in]*/ BSTR bstrTableSuffix,
		/*[in]*/ VARIANT vMsgId,  
		/*[out]*/ BSTR* pbstrErrorMessage);

protected:
	// MSMQ routines
	IMSMQMessagePtr HelperMSMQReceiveMessage(IMSMQQueuePtr MSMQQueuePtr, VARIANT vMsgId, long lMessagesToScan);
	bool HelperMSMQFindMessage(IMSMQQueuePtr MSMQQueuePtr, VARIANT vMsgId, long lMessagesToScan);
	void HelperMSMQSendMessage(BSTR bstrQueueName, BSTR bstrFormatName, BSTR bstrLabel, VARIANT vBody);
	void HelperMSMQSendMessage(IMSMQQueuePtr MSMQQueuePtr, BSTR bstrLabel, VARIANT vBody);
	void HelperMSMQSendMessage(IMSMQQueuePtr MSMQQueuePtr, IMSMQMessagePtr MSMQMessagePtr);

protected:
	// OMMQ routines
	_bstr_t HelperOMMQGetMoveQueueName(BSTR bstrQueueName) {return bstrQueueName + _bstr_t(L"X");}
	//_RecordsetPtr HelperOMMQReceiveMessage(long lProvider, _ConnectionPtr ConnectionPtr, VARIANT vMsgId);
	_RecordsetPtr HelperOMMQReceiveMessage(
		long lProvider, 
		_ConnectionPtr ConnectionPtr, 
		BSTR bstrTableSuffix,
		VARIANT vMsgId);
	//void HelperOMMQSendMessage(long lProvider, BSTR bstrConnectionString, VARIANT vMsgId, BSTR bstrQueueName, BSTR bstrProgID, const _bstr_t& bstrXML);
	void HelperOMMQSendMessage(long lProvider, BSTR bstrConnectionString, VARIANT vMsgId, BSTR bstrQueueName, BSTR bstrTableSuffix, BSTR bstrProgID, const _bstr_t& bstrXML);
	//void HelperOMMQSendMessage(long lProvider, _ConnectionPtr ConnectionPtr, VARIANT vMsgId, BSTR bstrQueueName, BSTR bstrProgID, const _bstr_t& bstrXML);
	void HelperOMMQSendMessage(long lProvider, _ConnectionPtr ConnectionPtr, VARIANT vMsgId, BSTR bstrQueueName, BSTR bstrTableSuffix, BSTR bstrProgID, const _bstr_t& bstrXML);
	void CloseCommand(_CommandPtr& CommandPtr);
	void CloseRecordset(_RecordsetPtr& RecordsetPtr);
	void CloseConnection(_ConnectionPtr& ConnectionPtr);
	void LogEventErrorProvider(_ConnectionPtr& ConnectionPtr);

protected:
	// error support
	void AddToErrorMessage(LPCTSTR pszFunctionName, LPCTSTR pszFormat, ...);
	void AddToErrorMessage(LPCTSTR pszFunctionName, _com_error& comerr, LPCTSTR pszFormat, ...);
	void ResetErrorMessage() {m_szErrorMessage[0] = '\0';}
	BSTR AllocbstrErrorMessage() {if (m_szErrorMessage[0] == '\0') {return NULL;} else {USES_CONVERSION; return ::SysAllocString((LPCOLESTR)T2W(&m_szErrorMessage[0]));}}

private:
	TCHAR m_szErrorMessage[MAXLOGMESSAGESIZE];
	static _bstr_t s_bstrNULL;
	static _variant_t s_vZero;
	static _variant_t s_vOne;
	static _variant_t s_vNull;
};

#endif //__MESSAGEQUEUELISTENERMTS1_H_
