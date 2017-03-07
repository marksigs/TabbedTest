///////////////////////////////////////////////////////////////////////////////
//	FILE:			WSASocket.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_WSASOCKET_H__9FDA4768_CC0B_4DFB_AD64_9C208311DD75__INCLUDED_)
#define AFX_WSASOCKET_H__9FDA4768_CC0B_4DFB_AD64_9C208311DD75__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "mutex.h"

class CWSASocket  
{
public:
	CWSASocket();
	CWSASocket(
		const WORD wVersionRequested, 
		const int nAddressFamily, 
		const int nType, 
		const int nProtocol,
		const LPWSAPROTOCOL_INFO lpProtocolInfo = NULL);
	virtual ~CWSASocket();

	virtual bool Exit();
	virtual bool Open(
		const WORD wVersionRequested = 0x02, 
		const int nAddressFamily = PF_INET, 
		const int nType = SOCK_STREAM, 
		const int nProtocol = 0,
		const LPWSAPROTOCOL_INFO lpProtocolInfo = NULL);
	virtual bool Close(const bool bLinger = 1, const int nInterval = 5, const bool bCleanup = FALSE);
	virtual bool Shutdown(int nHow) const;
	virtual bool Bind(int nPort = 0) const;
	virtual bool Listen(int nBackLog = SOMAXCONN) const;
	virtual bool Accept(CWSASocket& WSASocket) const;
	virtual bool Connect(LPCTSTR lpszHostName, const UINT uPort);
	virtual int Send(LPWSABUF wsaBufs, DWORD dwBufferCount, const DWORD dwFlags = 0) const;
	virtual int Send(LPCSTR lpszBuffer, const DWORD dwLen, const DWORD dwFlags = 0) const;
	virtual int SendBuffer(LPCSTR lpszBuffer, const DWORD dwLen, const DWORD dwFlags = 0) const;
	virtual int Recv(LPWSABUF wsaBufs, DWORD dwBufferCount, DWORD dwFlags = 0) const;
	virtual int Recv(LPSTR lpszBuffer, const DWORD dwLen, DWORD dwFlags = 0) const;
	virtual int RecvBuffer(LPSTR lpszBuffer, const DWORD dwLen, DWORD dwFlags = 0) const;
	virtual bool SetSockOpt(const int nLevel, const int nOptName, const void* lpOptVal, const int nOptLen) const;

	static bool WSAStartup(const WORD wVersionRequested = 0x02);
	static bool WSACleanup();
	static bool FindProtocols();
	static LPWSAPROTOCOL_INFO GetProtocols() { return m_pProtocols; }
	static int GetProtocolsNum() { return m_nProtocols; }

	LPHOSTENT GetHostByName(LPCTSTR lpszHostName);
	SOCKET GetSocket() const { return m_Socket; }
	void SetSendBlockSize(const DWORD dwSendBlockSize) { m_dwSendBlockSize = dwSendBlockSize; }
	void SetRecvBlockSize(const DWORD  dwRecvBlockSize) { m_dwRecvBlockSize = dwRecvBlockSize; }
	DWORD GetSendBlockSize() const { return m_dwSendBlockSize; }
	DWORD GetRecvBlockSize() const { return m_dwRecvBlockSize; }
	bool SetKeepAlive(const bool bKeepAlive);
	bool GetKeepAlive() const { return m_bKeepAlive; }
	UINT SetConnectRetries(const UINT nConnectRetries);
	UINT GetConnectRetries() const { return m_nConnectRetries; }
	bool IsValid() { return m_Socket != INVALID_SOCKET; }

protected:
	virtual bool Init();
	static void Error(LPCSTR lpszFile, int nLine, LPCTSTR ppszFormatplusArgs, ...);
	static void Error(int nError, LPCSTR lpszFile, int nLine, LPCTSTR ppszFormatplusArgs, ...);
	static void Error(int nError, LPCSTR lpszFile, int nLine, LPCTSTR* ppszFormatplusArgs);

private:
	LPHOSTENT 			m_lpHostEnt;
	DWORD 				m_dwSendBlockSize;
	DWORD 				m_dwRecvBlockSize;
	UINT				m_nConnectRetries;
	static LPWSAPROTOCOL_INFO m_pProtocols;
	static int			m_nProtocols;
	static bool			s_bStarted;
	static long			s_lReferenceCount;  
	static int			s_nIORetries;
    static namespaceMutex::CCriticalSection m_csStarted;
	static namespaceMutex::CCriticalSection m_csReferenceCount;

protected:
	WORD				m_wVersionRequested;
	int					m_nAddressFamily;
	int					m_nType;
	int					m_nProtocol;
	LPWSAPROTOCOL_INFO	m_lpProtocolInfo;
	LPTSTR				m_lpszHostName;
	UINT				m_uPort;                              
	bool				m_bKeepAlive;
	SOCKET				m_Socket;

protected:
	static const _TCHAR* GetWSAErrorStr(int nWSAErrorCode = NULL);
	static void SetIORetries(int nIORetries) { s_nIORetries = nIORetries; }
	static int GetIORetries(int nIORetries) { return s_nIORetries; }
};

#endif // !defined(AFX_WSASOCKET_H__9FDA4768_CC0B_4DFB_AD64_9C208311DD75__INCLUDED_)
