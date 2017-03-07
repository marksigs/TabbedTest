///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListener1.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AD      30/08/00    Initial version
//	LD		03/05/01	SYS2296 - Upgrade to version 3 of the xml parser
//  LD		22/08/02	SYS5464	- Wait for current messages being processed to stop
//  ML		17/03/03	Use password repository (OmigaPlus)
//  PSC     06/08/03    MSMS214 - Change configure lock to CriticalSection
//  RF      04/02/04	BMIDS727 - Different queues can use different tables;
//						           use MSXML4; make XML parsing more robust
///////////////////////////////////////////////////////////////////////////////

#ifndef __MESSAGEQUEUELISTENER1_H_
#define __MESSAGEQUEUELISTENER1_H_

#include "resource.h"       // main symbols
#include "ThreadPoolManagerFunctionQueue.h"
#include "ThreadPoolManagerOMMQ1.h"
#include "ThreadPoolManagerMSMQ1.h"
#include "ScheduleManagerDaily1.h"
#include <map>
#include "mutex.h"
#include "Log.h"

// RF 04/02/04
//#import "msxml3.dll"  rename_namespace("XmlNS")
#import "msxml4.dll" rename_namespace("XmlNS")

#define DEFAULT_REPOSITORY_PATH "C:\\Program Files\\Marlborough Stirling\\Omiga.repository"

using namespace std; 

_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerMSMQ1, __uuidof(IInternalThreadPoolManagerMSMQ1));
_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerOMMQ1, __uuidof(IInternalThreadPoolManagerOMMQ1));
_COM_SMARTPTR_TYPEDEF(IInternalThreadPoolManagerCommon1, __uuidof(IInternalThreadPoolManagerCommon1));

typedef map<_bstr_t, IInternalThreadPoolManagerCommon1Ptr> MapQueueNameToCommon1;

/////////////////////////////////////////////////////////////////////////////
// CMessageQueueListener1
class ATL_NO_VTABLE CMessageQueueListener1 : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CMessageQueueListener1, &CLSID_MessageQueueListener1>,
	public IDispatchImpl<IInternalMessageQueueListener1, &IID_IInternalMessageQueueListener1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public IDispatchImpl<IMessageQueueListener1, &IID_IMessageQueueListener1, &LIBID_MESSAGEQUEUELISTENERLib>,
	public CLog
{
public:
	CMessageQueueListener1();
	virtual ~CMessageQueueListener1();

	HRESULT FinalConstruct();
	void FinalRelease();

DECLARE_REGISTRY_RESOURCEID(IDR_MESSAGEQUEUELISTENER1)

DECLARE_PROTECT_FINAL_CONSTRUCT()
DECLARE_CLASSFACTORY_SINGLETON(CMessageQueueListener1)

BEGIN_COM_MAP(CMessageQueueListener1)
	COM_INTERFACE_ENTRY(IMessageQueueListener1)
	COM_INTERFACE_ENTRY(IInternalMessageQueueListener1)
//	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY2(IDispatch, IMessageQueueListener1)
END_COM_MAP()

// IMessageQueueListener1
public:
	STDMETHOD(Configure)(/*[in]*/ BSTR in_xml, /*[out,retval]*/ BSTR *out_xml);

// IInternalMessageQueueListener1
public:
	STDMETHOD(Start)();
	STDMETHOD(Stop)();
	STDMETHOD(LoadConfigurationDetails)();


protected:
	enum QueueTypeEnum
	{
		eQueueTypeMSMQ1 = 0,
		eQueueTypeOMMQ1 = 1,
		eQueueTypeUnknown
	};

	bstr_t Create(XmlNS::IXMLDOMDocument *ptrDomIn);
	void CreateQueue(const XmlNS::IXMLDOMNodePtr ptrQueueNode,	XmlNS::IXMLDOMDocumentPtr ptrDomOut);
	void CreateUnpackThreadInfo(
		XmlNS::IXMLDOMNode *ptrNode, 
		QueueTypeEnum eQueueType, 
		BSTR bstrQueueName, 
		BSTR bstrConnString, 
		DWORD dwPollInterval,
		BSTR bstrTableSuffix,
		XmlNS::IXMLDOMDocument *ptrDomOut);
	void CreateErrorInfo(XmlNS::IXMLDOMDocument *ptrDomIn, BSTR strType, BSTR strNumber, BSTR strSource, BSTR strDescription,_com_error& comerr);
	bstr_t Update(XmlNS::IXMLDOMDocument *ptrDomIn);
	void StartQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut);
	void StopQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut);
	void ReStartQueue(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut);
	void ReStartComponents(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, VARIANT pVar);
	void AddStalledComponents(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, VARIANT pVar);
	void ChangeQueueThreads(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, int nThreads);
	IInternalThreadPoolManagerCommon1Ptr GetThreadPoolManagerCommon1Ptr(QueueTypeEnum eQueueType, BSTR bstrQueueName);
	QueueTypeEnum GeteQueueType(BSTR bstrQueueType);
	void LogEventError(_com_error& comerr, LPCTSTR pszFormat, ...);
	bstr_t Get(XmlNS::IXMLDOMDocument *ptrDomIn);
	bstr_t RequestQueueInfo(BSTR bstrQueueName, QueueTypeEnum eQueueType, XmlNS::IXMLDOMDocument *ptrDomOut);
	bstr_t RequestAllQueueInfo(XmlNS::IXMLDOMDocument *ptrDomOut);
	bstr_t Delete(XmlNS::IXMLDOMDocument *ptrDomIn);
	void ConvertField(LPBYTE lpByte,LPSTR * ppStr,int iFieldSize,BOOL fRightToLeft);
	HRESULT GUIDFromString(LPWSTR lpWStr, GUID * pGuid);
	int WideToAnsi(LPSTR lpStr,LPWSTR lpWStr,int cchStr);
	int GetDigit(LPSTR lpstr);
	void RemoveEvent(BSTR bstrQueueName, QueueTypeEnum eQueueType,  XmlNS::IXMLDOMDocument	*ptrDomOut, GUID gKey);
	bstr_t LookupConnectionString( BSTR bstrDefaultConnStr );
	bstr_t GetPasswordFromRepository( BSTR bstrRepositoryPath, BSTR bstrInstance, BSTR bstrSchema);
	bstr_t GetConnStrSetting( bstr_t bstrConnString, bstr_t bstrSetting );


private:
	CScheduleManagerDaily m_ScheduleManagerDaily;

	// the lists of queues are held in an STL map
	MapQueueNameToCommon1 m_MapQueueNameToCommon1;

    // PSC 06/08/03 MSMS214
	namespaceMutex::CCriticalSection m_csConfigure;

	};

#endif //__MESSAGEQUEUELISTENER1_H_
