///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerLog.cpp
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

#include "stdafx.h"
#include "MessageQueueListenerLOG.h"
#include "MessageQueueListenerLOG1.h"
#include "ThreadPoolMessageFunctionQueue.h"

CThreadPoolManagerFunctionQueue* CMessageQueueListenerLOG1::s_pThreadPoolManagerFunctionQueue = NULL;
std::ofstream CMessageQueueListenerLOG1::s_LogFile;

///////////////////////////////////////////////////////////////////////////////
// CMessageQueueListenerLOG1

CMessageQueueListenerLOG1::CMessageQueueListenerLOG1()
{
}

CMessageQueueListenerLOG1::~CMessageQueueListenerLOG1()
{
}

HRESULT CMessageQueueListenerLOG1::FinalConstruct()
{
	return CComObjectRootEx<CComMultiThreadModel>::FinalConstruct();
}

void CMessageQueueListenerLOG1::FinalRelease()
{
	CComObjectRootEx<CComMultiThreadModel>::FinalRelease();
}

STDMETHODIMP CMessageQueueListenerLOG1::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IMessageQueueListenerLOG1
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CMessageQueueListenerLOG1::OnLog(long lLOGAREA, BSTR bstr)
{
	// current all areas are logged to file.
	// lLOGAREA is ignored.

	CThreadData* pThreadData = new CThreadData(bstr);

	CThreadPoolMessageFunctionQueue* pThreadPoolMessageFunctionQueue = new CThreadPoolMessageFunctionQueue(
		(CThreadPoolMessageFunctionQueue::funcptr)CMessageQueueListenerLOG1::ExecutionRoutine, 
		(CThreadPoolMessageFunctionQueue::funcparam)pThreadData
		);	
	s_pThreadPoolManagerFunctionQueue->AddFunctionToQueue(pThreadPoolMessageFunctionQueue);

	return S_OK;
}


bool CMessageQueueListenerLOG1::StartUp()
{
	bool bResult = false;

	_bstr_t bstrLogFileName = GetLogFileName();

	_ASSERTE(!s_LogFile.is_open());
	s_LogFile.open(bstrLogFileName, std::ios_base::out | std::ios_base::trunc);
	if (!s_LogFile.fail())
	{
		_ASSERTE(s_pThreadPoolManagerFunctionQueue == NULL);
		s_pThreadPoolManagerFunctionQueue = new CThreadPoolManagerFunctionQueue();
		bResult = s_pThreadPoolManagerFunctionQueue->StartUp(1);
	}
	return bResult;
}

void CMessageQueueListenerLOG1::CloseDown()
{
	_ASSERTE(s_pThreadPoolManagerFunctionQueue != NULL);
	s_pThreadPoolManagerFunctionQueue->CloseDown();
	delete s_pThreadPoolManagerFunctionQueue;
	s_pThreadPoolManagerFunctionQueue = NULL;

	if (s_LogFile.is_open())
	{
		s_LogFile.close();
	}
}

int CMessageQueueListenerLOG1::ExecutionRoutine(void* pv)
{
	CThreadData* pThreadData = reinterpret_cast<CThreadData*>(pv);

	// update the log file
	s_LogFile << (LPCSTR)pThreadData->m_bstr;
	s_LogFile.flush();

	delete pThreadData;
	return 0;	
}

_bstr_t CMessageQueueListenerLOG1::GetLogFileName()
{
	TCHAR szModuleFileName[_MAX_PATH]; 
	GetModuleFileName(NULL, szModuleFileName, _MAX_PATH); 
	
	TCHAR* pchSlash = _tcsrchr(szModuleFileName, '\\');
	if (pchSlash)
	{
		_tcscpy(pchSlash, _T("\\MessageQueueListenerLOG.log"));
	}
	return _bstr_t(szModuleFileName);
}

