// Profiler.h: interface for the CProfiler class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PROFILER_H__61C4B05D_87A2_4D50_BAD7_061F63984EC0__INCLUDED_)
#define AFX_PROFILER_H__61C4B05D_87A2_4D50_BAD7_061F63984EC0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>
#include <time.h>
#include "mutex.h"

class CProfiler  
{
private:
	class CTimer;

public:
	CProfiler();
	virtual ~CProfiler();

	static void StartProfiling();
	static void StopProfiling();
	static void StartTimer(LPCTSTR szLabel, bool bReset = false);
	static void StopTimer(LPCTSTR szLabel);

	static const std::_tstring GetTimesCSV();
	static void GetTimesXML(MSXML::IXMLDOMNodePtr ptrNode);

protected:
	typedef std::map<const std::_tstring, CTimer*> MapTimerType;

	static CProfiler* GetProfiler();
	static void RemoveProfiler();

	void _StartTimer(LPCTSTR szLabel, bool bReset = false);
	void _StopTimer(LPCTSTR szLabel);
	const std::_tstring _GetTimesCSV();
	void _GetTimesXML(MSXML::IXMLDOMNodePtr ptrNode);
	void RemoveAllTimers();
	void RemoveTimer(LPCTSTR szLabel);
	void SetTimerTime(LPCTSTR szLabel, int nTotalTime = 0);
	const int GetTimerTime(LPCTSTR szLabel);
	CTimer* GetTimer(LPCTSTR szLabel, bool bAdd = false);
	MapTimerType& GetMapTimers() { return m_MapTimers; }

private:
	class CTimer
	{
	public:
		CTimer(LPCTSTR pszLabel = NULL) : m_nStartTime(0), m_nTotalTime(0), m_nHits(0)
		{
			SetLabel(pszLabel);
		}
		void Start(bool bReset = false) { if (bReset) { SetTotalTime(0); SetHits(0); } m_nStartTime = GetCPUTimeMilliSecs(); SetHits(GetHits() + 1); }
		void Stop() { m_nTotalTime += max(0, GetCPUTimeMilliSecs() - m_nStartTime); m_nStartTime = 0; }
		void SetLabel(LPCTSTR pszLabel)
		{
			if (pszLabel != NULL && _tcslen(pszLabel) > 0)
			{
				m_sLabel = pszLabel;
			}
			else
			{
				m_sLabel = _T("");
			}
		}
		const std::_tstring GetLabel() const { return m_sLabel; }
		void SetTotalTime(int nTotalTime = 0) { m_nTotalTime = nTotalTime; }
		const int GetTotalTime() const { return m_nTotalTime; }
		void SetHits(int nHits = 0) { m_nHits = nHits; }
		const int GetHits() const { return m_nHits; }
		const int GetCPUTimeMilliSecs() const { return static_cast<int>((static_cast<double>(clock()) / CLOCKS_PER_SEC) * 1000); }

	private:
		int					m_nStartTime;
		int					m_nTotalTime;
		int					m_nHits;
		std::_tstring		m_sLabel;
	};

	MapTimerType			m_MapTimers;

	__declspec(thread) static CProfiler*		s_pProfiler;
};

//////////////////////////////////////////////////////////////////////
// Wrapper class which stops timer when object goes out of scope.
//////////////////////////////////////////////////////////////////////
class CProfile
{
public:
	CProfile(LPCTSTR szLabel) : m_sLabel(szLabel) { CProfiler::StartTimer(szLabel, false); }
	virtual ~CProfile() { CProfiler::StopTimer(m_sLabel.c_str()); }

private:
	std::_tstring	m_sLabel;
};

#ifdef PROFILE
#define IF_PROFILE(expr) expr
#else
#define IF_PROFILE(expr) void(0);
#endif

#endif // !defined(AFX_PROFILER_H__61C4B05D_87A2_4D50_BAD7_061F63984EC0__INCLUDED_)
