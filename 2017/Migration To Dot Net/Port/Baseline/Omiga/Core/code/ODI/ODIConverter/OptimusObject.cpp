///////////////////////////////////////////////////////////////////////////////
//	FILE:			OptimusObject.cpp
//	DESCRIPTION:	Encapsulates a deserialized Optimus object.
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      22/10/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CodePage.h"
#include "OptimusObject.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusObject::COptimusObject
//	
//	Description:
//		Encapulates an Optimus object. Mainly used for debugging/logging.
//	
//	Parameters:
//		BYTE* pObject:
//			Pointer to an existing, serialized Optimus object.
//		LPCWSTR szTag:
//			Name of the Optimus object.
//		LPCWSTR lpUnicode:
//			Unicode version of Optimus object payload data.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
COptimusObject::COptimusObject(BYTE* pObject, LPCWSTR szTag, LPCWSTR lpUnicode) :
	m_pObject(pObject),
	m_nLength(0),
	m_lNumber(0),
	m_lParent(0),
	m_nId(0),
	m_bstrTag(L""),
	m_lpUnicode(NULL),
	m_nMaxSizeUnicode(0),
	m_bOwnsUnicode(false),
	m_lpEBCDIC(NULL),
	m_nMaxSizeEBCDIC(0)
{
	Decode(pObject, szTag, lpUnicode);
}

COptimusObject::~COptimusObject()
{
	if (m_bOwnsUnicode)
	{
		delete []m_lpUnicode;
		m_lpUnicode = NULL;
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusObject::Decode
//	
//	Description:
//		Converts a serialized Optimus object in a byte stream to a COptimusObje
//		object.ct 
//	
//	Parameters:
//		BYTE* pObject:
//			Pointer to the byte stream containing the serialized Optimus object
//		LPCWSTR szTag:.
//			Name of the Optimus object.
//		LPCWSTR lpUnicode:
//			Buffer that contains/will receive the payload data as Unicode.
//		CCodePage* pCodePage:
//			Code page for converting the payload data from Ebcdic to Unicode.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusObject::Decode(BYTE* pObject, LPCWSTR szTag, LPCWSTR lpUnicode, CCodePage* pCodePage)
{
	if (pObject == NULL)
	{
		pObject = m_pObject;
	}
	else
	{
		m_pObject = pObject;
	}

	if (pObject != NULL)
	{
		m_nLength = GetShort(pObject);
		pObject += sizeof(m_nLength);
		m_lNumber = GetLong(pObject);
		pObject += sizeof(m_lNumber);
		m_lParent = GetLong(pObject);
		pObject += sizeof(m_lParent);
		m_nId = GetShort(pObject);
		pObject += sizeof(m_nId);

		m_bstrTag = szTag;

		m_lpEBCDIC = reinterpret_cast<LPCSTR>(pObject);
		m_nMaxSizeEBCDIC = m_nLength - (sizeof(m_nLength) + sizeof(m_lNumber) + sizeof(m_lParent) + sizeof(m_nId));

		if (m_bOwnsUnicode == true)
		{
			delete []m_lpUnicode;
			m_lpUnicode = NULL;
		}

		if (lpUnicode == NULL)
		{
			m_nMaxSizeUnicode = (m_nMaxSizeEBCDIC + 1) * 2;
			m_lpUnicode = new wchar_t[m_nMaxSizeUnicode];
			pCodePage->RecvMBC2W(reinterpret_cast<LPCSTR>(pObject), m_nMaxSizeEBCDIC, m_lpUnicode, m_nMaxSizeUnicode);
			m_bOwnsUnicode = true;
		}
		else
		{
			m_lpUnicode = const_cast<LPWSTR>(lpUnicode);
			m_bOwnsUnicode = false;
		}
		m_nMaxSizeUnicode = wcslen(lpUnicode);
	}
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusObject::GetShort
//	
//	Description:
//		Retrieves serialized short integer from start of byte stream.
//	
//	Parameters:
//		const BYTE* pStream:
//			The byte stream.
//	
//	Return:
//		short: 	
//			The short integer.
///////////////////////////////////////////////////////////////////////////////
short COptimusObject::GetShort(const BYTE* pStream)
{
	short nShort = MAKEWORD(pStream[1], pStream[0]);
	return nShort;
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusObject::GetLong
//	
//	Description:
//		Retrieves serialized long integer from start of byte stream.
//	
//	Parameters:
//		const BYTE* pStream:
//			The byte stream.
//	
//	Return:
//		long: 	
//			The long integer.
///////////////////////////////////////////////////////////////////////////////
long COptimusObject::GetLong(const BYTE* pStream)
{
	long llong = MAKELONG(MAKEWORD(pStream[3], pStream[2]), MAKEWORD(pStream[1], pStream[0]));
	return llong;
}

void COptimusObject::LogDebugHeader()
{
	_Module.LogDebug(_T("Length,Number,Parent,Id,Tag,Unicode,Ebcdic\n"));
}

///////////////////////////////////////////////////////////////////////////////
//	Function: COptimusObject::LogDebug
//	
//	Description:
//		Logs decoded Optimus object.
//	
//	Parameters:
//		bool bLogHex:
//			If true, then the hex for all decoded members is also logged. The 
//			hex can be used to directly compare the decoded object against the 
//			original byte stream.
//	
//	Return:
//		None
///////////////////////////////////////////////////////////////////////////////
void COptimusObject::LogDebug(bool bLogHex)
{
	const int nMaxSizeString = (m_nLength * 10);
	_TCHAR* pszStream = new _TCHAR[nMaxSizeString];

	if (pszStream != NULL)
	{
		::ZeroMemory(pszStream, nMaxSizeString);

		LogDebugItem(pszStream, m_nLength, bLogHex);
		_tcscat(pszStream, _T(","));
		LogDebugItem(pszStream, m_lNumber, bLogHex);
		_tcscat(pszStream, _T(","));
		LogDebugItem(pszStream, m_lParent, bLogHex);
		_tcscat(pszStream, _T(","));
		LogDebugItem(pszStream, m_nId, bLogHex);
		_tcscat(pszStream, _T(","));
		_stprintf(pszStream + wcslen(pszStream), _T("%s,"), static_cast<LPCWSTR>(m_bstrTag));
		LogDebugItem(pszStream, m_lpUnicode, bLogHex);

		if (bLogHex)
		{
			_tcscat(pszStream, _T(","));
			LogDebugHexString(pszStream, reinterpret_cast<BYTE*>(const_cast<LPSTR>(m_lpEBCDIC)), m_nMaxSizeEBCDIC, bLogHex);
		}

		_Module.LogDebug(_T("%s\n"), pszStream);

		delete []pszStream;
	}
}

void COptimusObject::LogDebugItem(_TCHAR* pszStream, short nItem, bool bLogHex)
{
	_stprintf(pszStream + wcslen(pszStream), _T("%d"), nItem);
	LogDebugHexString(pszStream, reinterpret_cast<BYTE*>(&nItem), sizeof(nItem), bLogHex, true);
}

void COptimusObject::LogDebugItem(_TCHAR* pszStream, long lItem, bool bLogHex)
{
	_stprintf(pszStream + wcslen(pszStream), _T("%ld"), lItem);
	LogDebugHexString(pszStream, reinterpret_cast<BYTE*>(&lItem), sizeof(lItem), bLogHex, true);
}

void COptimusObject::LogDebugItem(_TCHAR* pszStream, LPCWSTR lpItem, bool bLogHex)
{
	_stprintf(pszStream + wcslen(pszStream), _T("%s"), lpItem);
	LogDebugHexString(pszStream, reinterpret_cast<BYTE*>(const_cast<LPWSTR>(lpItem)), wcslen(lpItem) * sizeof(wchar_t), bLogHex);
}

void COptimusObject::LogDebugHexString(_TCHAR* pszStream, BYTE* pHex, int nMaxSizeHex, bool bLogHex, bool bReversed)
{
	if (bLogHex)
	{
		_stprintf(pszStream + wcslen(pszStream), _T("("));
		HexToString(pszStream, pHex, nMaxSizeHex, bReversed);
		_stprintf(pszStream + wcslen(pszStream), _T(")"));
	}
}

_TCHAR* COptimusObject::HexToString(_TCHAR* pszStream, BYTE* pHex, int nMaxSizeHex, bool bReversed)
{
	for (int nByte = 0; nByte < nMaxSizeHex; nByte++)
	{
		if (nByte > 0)
		{
			_tcscat(pszStream, _T(" "));
		}
		_stprintf(pszStream + wcslen(pszStream), _T("%02X"), bReversed ? *((pHex + (nMaxSizeHex - 1)) - nByte) : *(pHex + nByte));
	}

	return pszStream;
}

