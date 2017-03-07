///////////////////////////////////////////////////////////////////////////////
//	FILE:			Profiler.cpp
//	DESCRIPTION:	Simple profiler that records its results in XML.
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		28/09/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <sstream>
#include "Profiler.h"


__declspec(thread) CProfiler* CProfiler::s_pProfiler = NULL;

//////////////////////////////////////////////////////////////////////
// Static functions
//////////////////////////////////////////////////////////////////////
void CProfiler::StartProfiling()
{
	GetProfiler();
}

void CProfiler::StopProfiling()
{
	RemoveProfiler();
}

CProfiler* CProfiler::GetProfiler()
{
	if (s_pProfiler == NULL)
	{
		s_pProfiler = new CProfiler;
		_ASSERTE(s_pProfiler);
	}

	return s_pProfiler;
}

void CProfiler::RemoveProfiler()
{
	delete s_pProfiler;
	s_pProfiler = NULL;
}

void CProfiler::StartTimer(LPCTSTR szLabel, bool bReset)
{
	GetProfiler()->_StartTimer(szLabel, bReset);
}

void CProfiler::StopTimer(LPCTSTR szLabel)
{
	GetProfiler()->_StopTimer(szLabel);
}

void CProfiler::GetTimesXML(MSXML::IXMLDOMNodePtr ptrNode)
{
	GetProfiler()->_GetTimesXML(ptrNode);
}

const std::_tstring CProfiler::GetTimesCSV()
{
	return GetProfiler()->_GetTimesCSV();
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CProfiler::CProfiler()
{
}

CProfiler::~CProfiler()
{
	RemoveAllTimers();
}

void CProfiler::RemoveAllTimers()
{
	MapTimerType& MapTimers = GetMapTimers();

	if (MapTimers.size() > 0)
	{
		MapTimerType::iterator it;
		for (it = MapTimers.begin(); it != MapTimers.end(); it++)
		{
			CTimer* pTimer= (*it).second;

			delete pTimer;
		}
		MapTimers.clear();
	}
}

void CProfiler::_StartTimer(LPCTSTR szLabel, bool bReset)
{
	CTimer* pTimer = NULL;

	pTimer = GetTimer(szLabel, true);

	_ASSERTE(pTimer != NULL);

	if (pTimer != NULL)
	{
		pTimer->Start(bReset);
	}
}

void CProfiler::_StopTimer(LPCTSTR szLabel)
{
	CTimer* pTimer = NULL;

	pTimer = GetTimer(szLabel, true);

	_ASSERTE(pTimer != NULL);

	if (pTimer != NULL)
	{
		pTimer->Stop();
	}
}

void CProfiler::SetTimerTime(LPCTSTR szLabel, int nTotalTime)
{
	CTimer* pTimer = NULL;

	pTimer = GetTimer(szLabel, true);

	_ASSERTE(pTimer != NULL);

	if (pTimer != NULL)
	{
		pTimer->SetTotalTime(nTotalTime);
	}
}

const int CProfiler::GetTimerTime(LPCTSTR szLabel)
{
	CTimer* pTimer = NULL;
	int nTotalTime = 0;

	pTimer = GetTimer(szLabel, true);

	_ASSERTE(pTimer != NULL);

	if (pTimer != NULL)
	{
		nTotalTime = pTimer->GetTotalTime();
	}

	return nTotalTime;
}

void CProfiler::RemoveTimer(LPCTSTR szLabel)
{
	MapTimerType& MapTimers = GetMapTimers();

	MapTimerType::iterator it = MapTimers.find(szLabel);
	if (it != MapTimers.end())
	{	
		MapTimers.erase(it);
	}	
}

CProfiler::CTimer* CProfiler::GetTimer(LPCTSTR szLabel, bool bAdd)
{
	CTimer* pTimer = NULL;

	MapTimerType& MapTimers = GetMapTimers();

	MapTimerType::iterator it = MapTimers.find(szLabel);
	if (it == MapTimers.end())
	{
		if (bAdd)
		{
			pTimer = new CTimer(szLabel);

			_ASSERTE(pTimer != NULL);

			if (pTimer != NULL)
			{
				MapTimers.insert(MapTimerType::value_type(pTimer->GetLabel(), pTimer));
			}
		}
	}
	else
	{
		pTimer = (*it).second;
	}	

	return pTimer;
}

const std::_tstring CProfiler::_GetTimesCSV()
{
	std::_tstring sTimesCSV;

	MapTimerType& MapTimers = GetMapTimers();

	if (MapTimers.size() > 0)
	{
		MapTimerType::iterator it;
		for (it = MapTimers.begin(); it != MapTimers.end(); it++)
		{
			CTimer* pTimer= (*it).second;

			_ASSERTE(pTimer != NULL);

			if (pTimer != NULL)
			{
				if (sTimesCSV.size() > 0)
				{
					sTimesCSV += _T(",");
				}
				std::_tostringstream strmTime;
				strmTime << pTimer->GetLabel() << _T(",") << pTimer->GetTotalTime();
				sTimesCSV += strmTime.str();
			}
		}
	}

	return sTimesCSV;
}

void CProfiler::_GetTimesXML(MSXML::IXMLDOMNodePtr ptrNode)
{
	MSXML::IXMLDOMDocumentPtr ptrXMLDOMDocument(__uuidof(MSXML::DOMDocument));

	ptrXMLDOMDocument = ptrNode->ownerDocument;

	if (ptrXMLDOMDocument != NULL)
	{
		MSXML::IXMLDOMElementPtr ptrXMLDOMElementProfile;

		ptrXMLDOMElementProfile = ptrXMLDOMDocument->createElement(L"PROFILE");
		MSXML::IXMLDOMElementPtr ptrXMLDOMDocumentElement = ptrXMLDOMDocument->GetdocumentElement();
		
		if (ptrXMLDOMDocumentElement != NULL)
		{
			ptrXMLDOMDocumentElement->appendChild(ptrXMLDOMElementProfile);

			MapTimerType& MapTimers = GetMapTimers();

			if (MapTimers.size() > 0)
			{
				MapTimerType::iterator it;
				for (it = MapTimers.begin(); it != MapTimers.end(); it++)
				{
					CTimer* pTimer= (*it).second;

					_ASSERTE(pTimer != NULL);

					if (pTimer != NULL)
					{
						MSXML::IXMLDOMElementPtr ptrXMLDOMElementTimer = ptrXMLDOMDocument->createElement(pTimer->GetLabel().c_str());
						ptrXMLDOMElementProfile->appendChild(ptrXMLDOMElementTimer);

						MSXML::IXMLDOMElementPtr ptrXMLDOMElementHits = ptrXMLDOMDocument->createElement(L"HITS");
						std::_tostringstream strmHits;
						strmHits << pTimer->GetHits();
						ptrXMLDOMElementHits->Puttext(strmHits.str().c_str());

						MSXML::IXMLDOMElementPtr ptrXMLDOMElementTotalTime = ptrXMLDOMDocument->createElement(L"TOTALTIME");
						std::_tostringstream strmTotalTime;
						strmTotalTime << pTimer->GetTotalTime();
						ptrXMLDOMElementTotalTime->Puttext(strmTotalTime.str().c_str());

						ptrXMLDOMElementTimer->appendChild(ptrXMLDOMElementHits);
						ptrXMLDOMElementTimer->appendChild(ptrXMLDOMElementTotalTime);
					}
				}
			}
		}
	}
}

