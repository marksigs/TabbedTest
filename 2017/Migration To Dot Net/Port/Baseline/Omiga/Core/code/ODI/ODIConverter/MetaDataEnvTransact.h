// MetaDataEnvTransact.h: interface for the CMetaDataEnvTransactTransact class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_METADATAENVTRANSACT_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_)
#define AFX_METADATAENVTRANSACT_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "MetaDataEnv.h"

#include <map>
#include "WSASocket.h"


class CMetaDataEnvTransact : public CMetaDataEnv
{
public:
	CMetaDataEnvTransact(LPCWSTR szType, LPCWSTR szName);
	virtual ~CMetaDataEnvTransact();

	virtual bool Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent);
	virtual void Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent);
	_bstr_t GetTransactHost() { return m_bstrTransactHost; }
	int GetTransactPort() { return m_nTransactPort; }
	int GetSktTimeOutMs() { return m_nSktTimeOutMs; }
	int GetSktMaxSend() { return m_nSktMaxSend; }
	int GetSktMaxRecv() { return m_nSktMaxRecv; }
	void SetProfile(const bool bProfile) { m_bProfile = bProfile; }
	bool GetProfile() const { return m_bProfile; }
	void SetTestResponse(const bool bTestResponse) { m_bTestResponse = bTestResponse; }
	bool GetTestResponse() const { return m_bTestResponse; }
	void SetTestResponseDelayMs(const int nTestResponseDelayMs) { m_nTestResponseDelayMs = nTestResponseDelayMs; }
	int GetTestResponseDelayMs() const { return m_nTestResponseDelayMs; }
	void SetLogDebug(const bool bLogDebug) { m_bLogDebug = bLogDebug; }
	bool GetLogDebug() const { return m_bLogDebug; }
	void SetLogXML(const int nLogXML) { m_nLogXML = nLogXML; }
	int GetLogXML() const { return m_nLogXML; }

private:
    namespaceMutex::CCriticalSection	m_csInit;

	_bstr_t				m_bstrTransactHost;			// Transact host for this environment.
	int					m_nTransactPort;			// Transact sign on port for this environment.
	int					m_nSktTimeOutMs;			// Send and receive time out for this environment.
	int					m_nSktMaxSend;				// Maximum size of send buffer.
	int					m_nSktMaxRecv;				// Maximum size of recv buffer.
	bool				m_bProfile;					// If true then profile calls.
	bool				m_bTestResponse;			// If true then don't call Optimus - send back test response.
	int					m_nTestResponseDelayMs;		// The delay before sending back the test response
	bool				m_bLogDebug;				// If true then send log output to debug window.
	int					m_nLogXML;					// Number of request/response XML to log.
};


#endif // !defined(AFX_METADATAENVTRANSACT_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_)
