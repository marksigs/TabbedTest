// LookUpTable.h: interface for the CLookUpTable class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_LOOKUPTABLE_H__4ABF7FF6_0A82_4769_8A81_5530B8ED69E9__INCLUDED_)
#define AFX_LOOKUPTABLE_H__4ABF7FF6_0A82_4769_8A81_5530B8ED69E9__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <map>

class CLookUpTable  
{
public:
	enum EDirection
	{
		dirNull,
		dirSend,
		dirRecv
	};

	CLookUpTable(LPCWSTR szName = NULL, const MSXML::IXMLDOMElementPtr ptrXMLDOMElement = NULL);
	virtual ~CLookUpTable();

	void Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMElement);
	void Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMElement);

	typedef std::pair<_bstr_t, _bstr_t> PairMapInToOutType;
	typedef std::map<_bstr_t, PairMapInToOutType, Nocase> MapInToOutType;

	MapInToOutType& GetMapInToOut(const EDirection Direction)
	{
		_ASSERTE(Direction != dirNull);
		return Direction == dirSend ? GetMapSendInToSendOut() : GetMapRecvInToRecvOut();
	}
	MapInToOutType& GetMapSendInToSendOut()
	{
		return m_MapSendInToSendOut;
	}
	MapInToOutType& GetMapRecvInToRecvOut()
	{
		return m_MapRecvInToRecvOut;
	}

	_bstr_t GetName() const { return m_bstrName; }
	_bstr_t GetSendSrc() const { return m_bstrSendSrc; }
	_bstr_t GetRecvSrc() const { return m_bstrRecvSrc; }
	_bstr_t GetSendValue(LPCWSTR szKey)	{ return GetValue(szKey, dirSend); }
	_bstr_t GetRecvValue(LPCWSTR szKey)	{ return GetValue(szKey, dirRecv); }
	_bstr_t GetValue(LPCWSTR szKey, const EDirection Direction);

	void WriteCSV(FILE* fp, LPCWSTR szProperty);

private:
	_bstr_t				m_bstrName;				// Unique identifier for this table.
	_bstr_t				m_bstrSendSrc;			// Source for the send conversions. Optional.
	_bstr_t				m_bstrRecvSrc;			// Source for the receive conversions. Optional.

	MapInToOutType		m_MapSendInToSendOut;	// Substitute values when sending the property data.
	MapInToOutType		m_MapRecvInToRecvOut;	// Substitute values when receiving the property data.
};

#endif // !defined(AFX_LOOKUPTABLE_H__4ABF7FF6_0A82_4769_8A81_5530B8ED69E9__INCLUDED_)

