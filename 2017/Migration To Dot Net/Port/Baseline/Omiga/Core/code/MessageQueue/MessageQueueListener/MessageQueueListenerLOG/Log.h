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
//	LD		15/05/02	SYS4618 - Make logging more robust
//  LD		20/06/02	SYS4933 - Add originating thread id
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_LOG_H__B302E0E3_E814_11D5_A235_005004E8D1A7__INCLUDED_)
#define AFX_LOG_H__B302E0E3_E814_11D5_A235_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#import "..\MessageQueueListenerLOG\MessageQueueListenerLOG.tlb"

#define MAXLOGMESSAGESIZE 2048

///////////////////////////////////////////////////////////////////////////////

class CLog
{
public:
	CLog();
	virtual ~CLog();

	TCHAR* FormatTimeNow(TCHAR* pszTimeNow, int nMaxLenTimeNow) const;
	TCHAR* FormatThreadID(TCHAR* pszThreadID, int nMaxLenThreadID) const;
	TCHAR* FormatOriginatingAndCurrentThreadID(TCHAR* pszThreadID, int nMaxLenThreadID, DWORD dwOriginatingThreadID) const;

protected:
	MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr m_MessageQueueListenerLOG1Ptr;

private:
	void LogToFileWithPrefix(
		MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr& MessageQueueListenerLOG1Ptr, 
		MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, 
		DWORD m_dwOriginatingThreadId,
		LPCTSTR pszString);

public:
	class CLogInOut
	{
	public:
		CLogInOut(CLog& rLog, MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, DWORD dwOriginatingThreadId = NULL);
		virtual ~CLogInOut();
		void Initialize(LPCTSTR pszFormat, ...);
		void AnyThreadInitialize(LPCTSTR pszFormat, ...);
	private:
		CLog& m_Log;
		_bstr_t m_bstrTrace;
		MESSAGEQUEUELISTENERLOGLib::LOGAREA m_lLogArea;
		DWORD m_dwOriginatingThreadId;
		MESSAGEQUEUELISTENERLOGLib::IMessageQueueListenerLOG1Ptr m_MessageQueueListenerLOG1Ptr;
	};
	friend class CLogInOut;

	class CLogIn
	{
	public:
		CLogIn(CLog& rLog, MESSAGEQUEUELISTENERLOGLib::LOGAREA lLogArea, DWORD dwOriginatingThreadId = NULL);
		virtual ~CLogIn();
		void Initialize(LPCTSTR pszFormat, ...);
		void AnyThreadInitialize(LPCTSTR pszFormat, ...);
	private:
		CLog& m_Log;
		MESSAGEQUEUELISTENERLOGLib::LOGAREA m_lLogArea;
		DWORD m_dwOriginatingThreadId;
	};
	friend class CLogIn;
};

#endif // !defined(AFX_LOG_H__B302E0E3_E814_11D5_A235_005004E8D1A7__INCLUDED_)
