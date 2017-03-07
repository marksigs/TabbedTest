// MetaDataEnvOptimus.h: interface for the CMetaDataEnvOptimusOptimus class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_METADATAENVOPTIMUS_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_)
#define AFX_METADATAENVOPTIMUS_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "MetaDataEnv.h"

#include <map>
#include "CodePages.h"
#include "LookUpTables.h"
#include "OptimusMetaData.h"
#include "WSASocket.h"

class CCodePage;

class CMetaDataEnvOptimus : public CMetaDataEnv
{
public:
	CMetaDataEnvOptimus(LPCWSTR szType, LPCWSTR szName);
	virtual ~CMetaDataEnvOptimus();

	virtual bool Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent);
	virtual void Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent);
	_bstr_t GetOptimusObjectMapPath() { return m_bstrOptimusObjectMapPath; }
	_bstr_t GetCodePageMapPath() { return m_bstrCodePageMapPath; }
	_bstr_t GetLookUpTablesPath() { return m_bstrLookUpTablesPath; }
	_bstr_t GetOptimusHost() { return m_bstrOptimusHost; }
	int GetOptimusPort() { return m_nOptimusPort; }
	_bstr_t GetOptimusHostName() { return m_bstrOptimusHostName; }
	_bstr_t GetOptimusEnvironment() { return m_bstrOptimusEnvironment; }
	int GetListenPort() { return m_nListenPort; }
	int GetSktTimeOutMs() { return m_nSktTimeOutMs; }
	int GetSktMaxSend() { return m_nSktMaxSend; }
	int GetSktMaxRecv() { return m_nSktMaxRecv; }
	CCodePage* GetCodePage(const MSXML::IXMLDOMNodePtr ptrXMLDOMNode);
	COptimusMetaData& GetOptimusMetaData() { return m_OptimusMetaData; }
	void SetProfile(const bool bProfile) { m_bProfile = bProfile; }
	bool GetProfile() const { return m_bProfile; }
	void SetTestResponse(const bool bTestResponse) { m_bTestResponse = bTestResponse; }
	bool GetTestResponse() const { return m_bTestResponse; }
	void SetTestResponseDelayMs(const int nTestResponseDelayMs) { m_nTestResponseDelayMs = nTestResponseDelayMs; }
	int GetTestResponseDelayMs() const { return m_nTestResponseDelayMs; }
	void SetLogDebug(const bool bLogDebug) { m_bLogDebug = bLogDebug; }
	bool GetLogDebug() const { return m_bLogDebug; }
	void SetLogHex(const bool bLogHex) { m_bLogHex = bLogHex; }
	bool GetLogHex() const { return m_bLogHex; }
	void SetLogXML(const int nLogXML) { m_nLogXML = nLogXML; }
	int GetLogXML() const { return m_nLogXML; }
	CCodePages& GetCodePages() { return m_CodePages; }
	CLookUpTables& GetLookUpTables() { return m_LookUpTables; }

	_bstr_t CMetaDataEnvOptimus::GetLookUpTableByNameValue(
		LPCWSTR szTable, 
		LPCWSTR szKey, 
		const CLookUpTable::EDirection Direction)
	{
		return m_LookUpTables.GetLookUpTableByNameValue(szTable, szKey, Direction);
	}
	_bstr_t CMetaDataEnvOptimus::GetLookUpTableByPropertyValue(
		LPCWSTR szTable, 
		LPCWSTR szKey, 
		const CLookUpTable::EDirection Direction)
	{
		return m_LookUpTables.GetLookUpTableByPropertyValue(szTable, szKey, Direction);
	}


protected:
	typedef std::pair<CMetaDataEnvOptimus*, CWSASocket*> PairAcceptDataType;
	void InitListeningSocket(const MSXML::IXMLDOMElementPtr ptrXMLDOMElementParent);
	static DWORD WINAPI ListeningSocketThreadProc(LPVOID lpParameter);
	static DWORD WINAPI AcceptSocketThreadProc(LPVOID lpParameter);
	void InitOptimusObjectMapPath();

private:
    namespaceMutex::CCriticalSection	m_csInit;

	_bstr_t				m_bstrOptimusObjectMapPath;	// Path to Optimus Object Map XML file.
	_bstr_t				m_bstrCodePageMapPath;		// Path to code page XML file.
	_bstr_t				m_bstrLookUpTablesPath;		// Path to look up tables file.
	_bstr_t				m_bstrNewCodePageMapPath;	// New path to code page XML file.
	_bstr_t				m_bstrNewLookUpTablesPath;	// New path to look up tables file.
	COptimusMetaData	m_OptimusMetaData;			// Optimus Object Map meta data.
	_bstr_t				m_bstrOptimusHost;			// Optimus host for this environment.
	int					m_nOptimusPort;				// Optimus sign on port for this environment.
	_bstr_t				m_bstrOptimusEnvironment;	// Optimus environment.
	_bstr_t				m_bstrOptimusHostName;		// Optimus host name as used in startNewSession request.
	int					m_nListenPort;				// The port on which to listen for Optimus push data.
	int					m_nSktTimeOutMs;			// Send and receive time out for this environment.
	int					m_nSktMaxSend;				// Maximum size of send buffer.
	int					m_nSktMaxRecv;				// Maximum size of recv buffer.
	bool				m_bProfile;					// If true then profile calls.
	bool				m_bTestResponse;			// If true then don't call Optimus - send back test response.
	int					m_nTestResponseDelayMs;		// The delay before sending back the test response
	bool				m_bLogDebug;				// If true then send log output to debug window.
	bool				m_bLogHex;					// If true then include hex dump in log output.
	int					m_nLogXML;					// Number of request/response XML to log.
	CCodePages			m_CodePages;				// Container for code pages for this environment.
	CLookUpTables		m_LookUpTables;				// Container for look up tables for this environment.
	CWSASocket			m_sktListener;
};


#endif // !defined(AFX_METADATAENVOPTIMUS_H__B71A4949_11BA_4EC3_9836_1235E26600BA__INCLUDED_)
