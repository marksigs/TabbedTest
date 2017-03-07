///////////////////////////////////////////////////////////////////////////////
//	FILE:			Log.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD		03/12/01	SYS3303 - Add support for logging to file
//  LD		03/12/01	SYS3303 - Correction to direction of arrows
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <sys/timeb.h>					 
#include <time.h>
#include "Log.h"

///////////////////////////////////////////////////////////////////////////////

CLog::CLog()
{

}

CLog::~CLog()
{

}

void CLog::LogToFileWithPrefix(MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr& MessageQueueListenerLOG1Ptr, MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, DWORD dwOriginatingThreadId, LPCTSTR pszString)
{
	_ASSERT(_tcslen(pszString) > 0);
	
	// determine the prefix for the message
	const int nMaxLenPrefix = MAXLOGMESSAGESIZE;
	TCHAR szNewPrefix[nMaxLenPrefix] = _T("");
	FormatTimeNow(szNewPrefix + _tcslen(szNewPrefix), nMaxLenPrefix - _tcslen(szNewPrefix));
	_tcscat(szNewPrefix, _T(": "));

	if (dwOriginatingThreadId != NULL)
	{
		FormatOriginatingAndCurrentThreadID(szNewPrefix + _tcslen(szNewPrefix), nMaxLenPrefix - _tcslen(szNewPrefix), dwOriginatingThreadId);
	}
	else
	{
		FormatThreadID(szNewPrefix + _tcslen(szNewPrefix), nMaxLenPrefix - _tcslen(szNewPrefix));
	}
	_tcscat(szNewPrefix, _T(": "));

	TCHAR* pszLogArea = NULL;
	switch (lLogArea)
	{
		case MESSAGEQUEUELISTENERLOGLib::LOGAREA_MQL:
			pszLogArea = _T("MQL");
			break;
        case MESSAGEQUEUELISTENERLOGLib::LOGAREA_MSMQ1:
			pszLogArea = _T("MSMQ1");
			break;
        case MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1MSMQ:
			pszLogArea = _T("MTS1MSMQ");
			break;
        case MESSAGEQUEUELISTENERLOGLib::LOGAREA_OMMQ1:
			pszLogArea = _T("OMMQ1");
			break;
        case MESSAGEQUEUELISTENERLOGLib::LOGAREA_MTS1OMMQ:
			pszLogArea = _T("MTS1OOMMQ");
			break;
        case MESSAGEQUEUELISTENERLOGLib::LOGAREA_COMPONENT:
			pszLogArea = _T("COMPONENT");
			break;
		default :
			_ASSERT(0);
			pszLogArea = _T("UNKNOWN");
			break;
	}
	// concatenate the message safely
	const int nMaxLenString = MAXLOGMESSAGESIZE;
	TCHAR szString[nMaxLenString] = _T("");
	VERIFY(_sntprintf(szString, nMaxLenString - 1, _T("%s%s %s"), szNewPrefix, pszLogArea, pszString) > -1);

	// and output it
	MessageQueueListenerLOG1Ptr->OnLog(lLogArea, szString);
}

TCHAR* CLog::FormatTimeNow(TCHAR* pszTimeNow, int nMaxLenTimeNow) const
{
	memset(pszTimeNow, 0, nMaxLenTimeNow);

	tm* ptimeLocal;
	_timeb timeNow;

	_ftime(&timeNow);
	ptimeLocal = localtime(&timeNow.time);

	if (ptimeLocal != NULL)
	{
		VERIFY(
			_sntprintf(
				pszTimeNow, 
				nMaxLenTimeNow - 1, 
				_T("%04d/%02d/%02d %02d:%02d:%02d.%03d"), 
				ptimeLocal->tm_year + 1900,
				ptimeLocal->tm_mon + 1,
				ptimeLocal->tm_mday,
				ptimeLocal->tm_hour,
				ptimeLocal->tm_min,
				ptimeLocal->tm_sec,
				timeNow.millitm) > -1);
	}

	return pszTimeNow;
}


TCHAR* CLog::FormatThreadID(TCHAR* pszThreadID, int nMaxLenThreadID) const
{
	memset(pszThreadID, 0, nMaxLenThreadID);

	DWORD dwCurrentThreadID	= ::GetCurrentThreadId();
	int nThreadPriority		= ::GetThreadPriority(::GetCurrentThread());

	VERIFY(
		_sntprintf(
			pszThreadID, 
			nMaxLenThreadID - 1, 
			_T("0x%03x(%d)"),
			dwCurrentThreadID, 
			nThreadPriority) > -1);

	return pszThreadID;
}

TCHAR* CLog::FormatOriginatingAndCurrentThreadID(TCHAR* pszThreadID, int nMaxLenThreadID, DWORD dwOriginatingThreadID) const
{
	memset(pszThreadID, 0, nMaxLenThreadID);

	DWORD dwCurrentThreadID	= ::GetCurrentThreadId();
	int nThreadPriority		= ::GetThreadPriority(::GetCurrentThread());

	VERIFY(
		_sntprintf(
			pszThreadID, 
			nMaxLenThreadID - 1, 
			_T("0x%03x/0x%03x(%d)"),
			dwOriginatingThreadID,
			dwCurrentThreadID, 
			nThreadPriority) > -1);

	return pszThreadID;
}


///////////////////////////////////////////////////////////////////////////////

CLog::CLogInOut::CLogInOut(CLog& rLog, MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, DWORD dwOriginatingThreadId) :
	m_Log(rLog),
	m_lLogArea(lLogArea),
	m_dwOriginatingThreadId(dwOriginatingThreadId)
{
}

CLog::CLogInOut::~CLogInOut()
{
	if (m_MessageQueueListenerLOG1Ptr)
	{
		try
		{
			m_Log.LogToFileWithPrefix(m_MessageQueueListenerLOG1Ptr, m_lLogArea, m_dwOriginatingThreadId, _bstr_t(L"<- ") +  m_bstrTrace);
		}
		catch(...)
		{
			_ASSERT(0);
		}
	}
}
void CLog::CLogInOut::Initialize(LPCTSTR pszFormat, ...)
{
	if (m_Log.m_MessageQueueListenerLOG1Ptr)
	{
		try
		{
			// expand the current message
			TCHAR    chMsg[MAXLOGMESSAGESIZE];
			va_list pArg;

			va_start(pArg, pszFormat);
			VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
			va_end(pArg);

			_ASSERT(_tcslen(chMsg) > 0);
			m_bstrTrace = chMsg;
			m_MessageQueueListenerLOG1Ptr = m_Log.m_MessageQueueListenerLOG1Ptr;
			m_Log.LogToFileWithPrefix(m_MessageQueueListenerLOG1Ptr, m_lLogArea, m_dwOriginatingThreadId, _bstr_t(L"-> ") +  m_bstrTrace);
		}
		catch(...)
		{
			_ASSERT(0);
		}
	}
}

void CLog::CLogInOut::AnyThreadInitialize(LPCTSTR pszFormat, ...)
{
	try
	{
		MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr MessageQueueListenerLOG1Ptr(__uuidof(MESSAGEQUEUELISTENERLOGLib::MessageQueueListenerLOG1));
		if (MessageQueueListenerLOG1Ptr)
		{
			// expand the current message
			TCHAR    chMsg[MAXLOGMESSAGESIZE];
			va_list pArg;

			va_start(pArg, pszFormat);
			VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
			va_end(pArg);

			_ASSERT(_tcslen(chMsg) > 0);
			m_bstrTrace = chMsg;
			m_MessageQueueListenerLOG1Ptr = MessageQueueListenerLOG1Ptr;
			m_Log.LogToFileWithPrefix(m_MessageQueueListenerLOG1Ptr, m_lLogArea, m_dwOriginatingThreadId, _bstr_t(L"-> ") +  chMsg);
		}		
	}
	catch(...)
	{
	}
}
///////////////////////////////////////////////////////////////////////////////

CLog::CLogIn::CLogIn(CLog& rLog, MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, DWORD dwOriginatingThreadId) :
	m_Log(rLog),
	m_lLogArea(lLogArea),
	m_dwOriginatingThreadId(dwOriginatingThreadId)
{
}

CLog::CLogIn::~CLogIn()
{
}
void CLog::CLogIn::Initialize(LPCTSTR pszFormat, ...)
{
	if (m_Log.m_MessageQueueListenerLOG1Ptr)
	{
		try
		{
			// expand the current message
			TCHAR    chMsg[MAXLOGMESSAGESIZE];
			va_list pArg;

			va_start(pArg, pszFormat);
			VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
			va_end(pArg);

			_ASSERT(_tcslen(chMsg) > 0);
			m_Log.LogToFileWithPrefix(m_Log.m_MessageQueueListenerLOG1Ptr, m_lLogArea, m_dwOriginatingThreadId, _bstr_t(L"-- ") +  chMsg);
		}
		catch(...)
		{
			_ASSERT(0);
		}
	}
}

void CLog::CLogIn::AnyThreadInitialize(LPCTSTR pszFormat, ...)
{
	try
	{
		MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr MessageQueueListenerLOG1Ptr(__uuidof(MESSAGEQUEUELISTENERLOGLib::MessageQueueListenerLOG1));
		if (MessageQueueListenerLOG1Ptr)
		{
			// expand the current message
			TCHAR    chMsg[MAXLOGMESSAGESIZE];
			va_list pArg;

			va_start(pArg, pszFormat);
			VERIFY(_vsntprintf(chMsg, MAXLOGMESSAGESIZE - 1, pszFormat, pArg) > -1);
			va_end(pArg);

			_ASSERT(_tcslen(chMsg) > 0);
			m_Log.LogToFileWithPrefix(MessageQueueListenerLOG1Ptr, m_lLogArea, m_dwOriginatingThreadId, _bstr_t(L"-- ") +  chMsg);
		}
	}
	catch(...)
	{
	}
}
