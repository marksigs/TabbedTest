///////////////////////////////////////////////////////////////////////////////
//	FILE:			WSASocket.cpp
//	DESCRIPTION:	Wrapper for Winsock API.
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS      20/08/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <Winsock2.h>
#include "WSASocket.h"
#include "Exception.h"


bool CWSASocket::s_bStarted				= false;
namespaceMutex::CCriticalSection CWSASocket::m_csStarted;
namespaceMutex::CCriticalSection CWSASocket::m_csReferenceCount;
long CWSASocket::s_lReferenceCount 		= 0;
int CWSASocket::s_nIORetries			= 0;
WSAPROTOCOL_INFO* CWSASocket::m_pProtocols = NULL;
int CWSASocket::m_nProtocols = 0;

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CWSASocket::CWSASocket()
{
	Init();
}

CWSASocket::CWSASocket(
	const WORD wVersionRequested, 
	const int nAddressFamily, 
	const int nType, 
	const int nProtocol,
	const LPWSAPROTOCOL_INFO lpProtocolInfo)
{
	if (Init())
	{
		Open(wVersionRequested, nAddressFamily, nType, nProtocol, lpProtocolInfo);
	}
}

CWSASocket::~CWSASocket()
{
	Exit();
}

bool CWSASocket::Init()
{
	m_wVersionRequested		= 0;
	m_nAddressFamily		= 0;
	m_nType					= 0;
	m_nProtocol				= 0;
	m_lpProtocolInfo		= NULL;
	m_lpHostEnt				= 0;
	m_lpszHostName			= NULL;
	m_uPort					= 0;
	m_dwSendBlockSize		= 1024;
	m_dwRecvBlockSize		= 1024;
	m_bKeepAlive			= false;
	m_nConnectRetries		= 0;
	m_Socket				= INVALID_SOCKET;

	return true;
}

bool CWSASocket::Exit()
{
	bool bSuccess = true;

	if (m_Socket != INVALID_SOCKET)
	{      
		bSuccess = Close();
	}           
	
	delete []m_lpszHostName;
	m_lpszHostName = NULL;

	return bSuccess;
}

bool CWSASocket::Open(
	const WORD wVersionRequested, 
	const int nAddressFamily, 
	const int nType, 
	const int nProtocol,
	const LPWSAPROTOCOL_INFO lpProtocolInfo)
{               
	bool bSuccess 			= false;

	m_wVersionRequested		= wVersionRequested;
	m_nAddressFamily		= nAddressFamily;
	m_nType					= nType;
	m_nProtocol				= nProtocol;
	m_lpProtocolInfo		= lpProtocolInfo;

	if (CWSASocket::WSAStartup(wVersionRequested))
	{
		GROUP Group;
		::ZeroMemory(&Group, sizeof(Group));

		m_Socket = ::WSASocket(m_nAddressFamily, m_nType, m_nProtocol, lpProtocolInfo, Group, WSA_FLAG_OVERLAPPED);

	    if (m_Socket == INVALID_SOCKET)
	    {
	    	Error(__FILE__, __LINE__, _T("WSASocket error"));
	    }
		else
		{                           
			bSuccess = true;

			// AS 15/07/99 enable use of keep alives, to prevent ISDN lines from dropping after time out.
			if (bSuccess && m_bKeepAlive)
			{
				bSuccess = SetSockOpt(SOL_SOCKET, SO_KEEPALIVE, &m_bKeepAlive, sizeof(m_bKeepAlive));
			}

			if (bSuccess)
			{
				::InterlockedIncrement(&s_lReferenceCount);
				bSuccess = true;
			}
		}
	}
		
	return bSuccess;	
}

bool CWSASocket::Close(const bool bLinger, const int nInterval, const bool bCleanup)
{           
	bool bSuccess = false;
	
	_ASSERTE(m_Socket != INVALID_SOCKET);
	
	if (m_Socket != INVALID_SOCKET)
	{
		LINGER Linger;
		Linger.l_onoff 		= (u_short)bLinger;
		Linger.l_linger 	= bLinger ? (u_short)nInterval : 0;
		SetSockOpt(SOL_SOCKET, SO_LINGER, &Linger, sizeof(Linger));

		if (::closesocket(m_Socket) == SOCKET_ERROR)
		{                             
			int nError = NULL;
			if ((nError = ::WSAGetLastError()) != WSAECONNABORTED)
			{
				Error(__FILE__, __LINE__, _T("closesocket(%d) error"), m_Socket);
			}
		}   
		else
		{           
			m_Socket = INVALID_SOCKET;
			            
			::InterlockedDecrement(&s_lReferenceCount);
			if (bCleanup && s_lReferenceCount == 0)
			{                                                    
				namespaceMutex::CSingleLock lck(&m_csReferenceCount, true);

				if (s_lReferenceCount == 0)
				{
					WSACleanup();	
				}
			}
			bSuccess = true;
		}
	}
	
	return bSuccess;
}

bool CWSASocket::Bind(int nPort) const
{
	bool bSuccess = false;
	
	_ASSERTE(m_Socket != INVALID_SOCKET);

	if (m_Socket != INVALID_SOCKET)
	{
		sockaddr_in cli_addr;

	    cli_addr.sin_family 		= AF_INET;
	    cli_addr.sin_addr.s_addr 	= INADDR_ANY;
	    cli_addr.sin_port			= htons((unsigned short)nPort);
	    
    	if (::bind(m_Socket, (LPSOCKADDR)&cli_addr, sizeof(cli_addr)) == SOCKET_ERROR)
    	{                                
    		Error(__FILE__, __LINE__, _T("bind error"));
    	}                     
    	else
    	{
    		bSuccess = true;
    	}
    }
	    
    return bSuccess;
}

bool CWSASocket::Listen(int nBackLog) const
{
	bool bSuccess = false;
	
	_ASSERTE(m_Socket != INVALID_SOCKET);

	if (m_Socket != INVALID_SOCKET)
	{
    	if (::listen(m_Socket, nBackLog) == SOCKET_ERROR)
    	{                                
    		Error(__FILE__, __LINE__, _T("listen error"));
    	}                     
    	else
    	{
    		bSuccess = true;
    	}
    }
	    
    return bSuccess;
}

bool CWSASocket::Accept(CWSASocket& WSASocket) const
{
	bool bSuccess = false;

	_ASSERTE(m_Socket != INVALID_SOCKET);

	if (m_Socket != INVALID_SOCKET)
	{
		sockaddr addr;
		int addrlen = sizeof(addr);

		_ASSERTE(WSASocket.m_Socket == INVALID_SOCKET);
		
		if ((WSASocket.m_Socket = ::WSAAccept(m_Socket, &addr, &addrlen, NULL, NULL)) == INVALID_SOCKET)
		{
   			Error(__FILE__, __LINE__, _T("WSAAccept error"));
		}
		else
		{
   			bSuccess = true;
		}
	}

	return bSuccess;
}

LPHOSTENT CWSASocket::GetHostByName(LPCTSTR lpszHostName)
{            
	if (m_lpHostEnt == NULL || _tcsicmp(lpszHostName, m_lpszHostName) != 0)
	{
		char* pmbsHostName = new char[_tcslen(lpszHostName) + 1];

		if (pmbsHostName != NULL)
		{
#ifdef _UNICODE
			::wcstombs(pmbsHostName, lpszHostName, _tcslen(lpszHostName) + 1); 
#else
			::strcpy(pmbsHostName, lpszHostName);
#endif
			m_lpHostEnt = ::gethostbyname(pmbsHostName);
			
			delete []pmbsHostName;

			if (m_lpHostEnt == NULL)
			{
	   			Error(__FILE__, __LINE__, _T("gethostbyname error"));
	   		}
		}
	}
		
	return m_lpHostEnt;
}

bool CWSASocket::Connect(LPCTSTR lpszHostName, const UINT uPort)
{
	bool bSuccess = false;
	
	_ASSERTE(m_Socket != INVALID_SOCKET);

	if (m_Socket != INVALID_SOCKET)
	{       
		LPHOSTENT lpHostEnt = NULL;
		                       
		if ((lpHostEnt = GetHostByName(lpszHostName)) != NULL)
		{
			sockaddr_in srv_addr;
			
	      	srv_addr.sin_family 		= PF_INET;
	      	srv_addr.sin_addr.s_addr 	= *(unsigned long*)*lpHostEnt->h_addr_list;
	      	srv_addr.sin_port 			= htons((u_short)uPort);
	  
			WSABUF wsabufCalleeData;
			::ZeroMemory(&wsabufCalleeData, sizeof(wsabufCalleeData));

			// AS 02/09/99
			// Under heavy load during load testing, ::connect can return WSAEADDRINUSE, e.g., the TCP/IP
			// stack on the client can't keep up, so loop until connection is OK.
			int nResult = NULL;
			UINT nRetry = 0;
			bSuccess = true;
		    while 
				(
					bSuccess &&
					(nResult = 
						::WSAConnect(
							m_Socket, 
							(LPSOCKADDR)&srv_addr, 
							sizeof(srv_addr),
							NULL, 
							&wsabufCalleeData,
							NULL,
							NULL)) == SOCKET_ERROR &&
					::WSAGetLastError() == WSAEADDRINUSE &&
					nRetry < m_nConnectRetries
				)
			{
				::Sleep(1000);

				bSuccess = 
					Close() &&			// Reopen socket.
					Open(m_wVersionRequested, m_nAddressFamily, m_nType, m_nProtocol) &&
					Bind();

				nRetry++;
			}
			
			if (nResult == SOCKET_ERROR)
		    {
	    		Error(__FILE__, __LINE__, _T("Connect error"));
				bSuccess = false;
	    	}                     
	    	else
	    	{                 
	   			delete []m_lpszHostName;
		   		m_lpszHostName = new _TCHAR[_tcslen(lpszHostName) + 1];
				if (m_lpszHostName != NULL)
				{
		   			_tcscpy(m_lpszHostName, lpszHostName);
				}
	    		m_uPort = uPort;
	    		bSuccess = true;
	    	}
		}
	}
	
	return bSuccess;
}

bool CWSASocket::Shutdown(int nHow) const
{           
	bool bSuccess = false;
	
	_ASSERTE(m_Socket != INVALID_SOCKET);

	if (m_Socket != INVALID_SOCKET)
	{
		if (::shutdown(m_Socket, nHow) == SOCKET_ERROR)
		{                             
			Error(__FILE__, __LINE__, _T("shutdown error"));
		}
		else
		{
			bSuccess = true;
		}
	}

	return bSuccess;
}

bool CWSASocket::SetSockOpt(const int nLevel, const int nOptName, const void* lpOptVal, const int nOptLen) const
{                  
	bool bSuccess = false;
	
	if (::setsockopt(m_Socket, nLevel, nOptName, static_cast<const char*>(lpOptVal), nOptLen) == SOCKET_ERROR)
	{	
		Error(__FILE__, __LINE__, _T("setsockopt(%d) error"), m_Socket);
	}                                    
	else
	{
		bSuccess = true;
	}
	
	return bSuccess;
}

int CWSASocket::Send(LPWSABUF wsaBufs, DWORD dwBufferCount, const DWORD dwFlags) const
{                 
	int nResult = 0;
	DWORD dwNumberOfBytesSent = 0;

	if ((nResult = ::WSASend(m_Socket, wsaBufs, dwBufferCount, &dwNumberOfBytesSent, dwFlags, NULL, NULL)) != SOCKET_ERROR)
	{	
		nResult = dwNumberOfBytesSent;
	}
	else
	{
		Error(__FILE__, __LINE__, _T("WSASend error"));
	}                                    
	
	return nResult;
}

int CWSASocket::Send(LPCSTR lpszBuffer, const DWORD dwLen, const DWORD dwFlags) const
{                 
	WSABUF wsaBufs[1];
	wsaBufs[0].buf = const_cast<LPSTR>(lpszBuffer);
	wsaBufs[0].len = dwLen;

	return Send(wsaBufs, 1, dwFlags);
}

int CWSASocket::SendBuffer(LPCSTR lpszBuffer, const DWORD dwLen, const DWORD dwFlags) const
{
	int nResult					= 0;
	DWORD dwNumberOfBytesLeft	= dwLen;
	DWORD dwNumberOfBytesSent	= 0;
	int nRetry					= 0;

  	while (dwNumberOfBytesLeft > 0) 
  	{
		nResult =
    		Send(lpszBuffer, dwNumberOfBytesLeft > m_dwSendBlockSize ? m_dwSendBlockSize : dwNumberOfBytesLeft, dwFlags);

	    if (nResult != SOCKET_ERROR)
	    {
			dwNumberOfBytesSent = static_cast<DWORD>(nResult);
	    	dwNumberOfBytesLeft -= dwNumberOfBytesSent;
    		lpszBuffer += dwNumberOfBytesSent;
    	}
    	else
    	{
			// WSAETIMEDOUT check - if we get this error then retry the send.
			if (::WSAGetLastError() == WSAETIMEDOUT && s_nIORetries > 0 && nRetry < s_nIORetries)
			{
				nRetry++;
			}
			else
			{
		   		break;
			}
    	}
  	}                     

	if (nResult != SOCKET_ERROR)
	{
		dwNumberOfBytesSent = dwLen - dwNumberOfBytesLeft;
	}                           
	
  	return dwNumberOfBytesSent;
}

int CWSASocket::Recv(LPWSABUF wsaBufs, DWORD dwBufferCount, DWORD dwFlags) const
{                 
	int nResult = 0;
	DWORD dwNumberOfBytesRecvd = 0;

	if ((nResult = ::WSARecv(m_Socket, wsaBufs, dwBufferCount, &dwNumberOfBytesRecvd, &dwFlags, NULL, NULL)) != SOCKET_ERROR)
	{
		nResult = dwNumberOfBytesRecvd;
	}
	else
	{	
		Error(__FILE__, __LINE__, _T("WSARecv error"));
	}                                    
	
	return nResult;
}

int CWSASocket::Recv(LPSTR lpszBuffer, const DWORD dwLen, DWORD dwFlags) const
{                 
	WSABUF wsaBufs[1];
	wsaBufs[0].buf = lpszBuffer;
	wsaBufs[0].len = dwLen;

	return Recv(wsaBufs, 1, dwFlags);
}

int CWSASocket::RecvBuffer(LPSTR lpszBuffer, const DWORD dwLen, DWORD dwFlags) const
{
	int nResult					= 0;
	DWORD dwNumberOfBytesLeft	= dwLen;
	DWORD dwNumberOfBytesRecvd	= 0;
	int nRetry					= 0;

	while (dwNumberOfBytesLeft > 0) 
	{
    	nResult = 
    		Recv(lpszBuffer, dwNumberOfBytesLeft > m_dwRecvBlockSize ? m_dwRecvBlockSize : dwNumberOfBytesLeft, dwFlags);
                
		if (nResult != SOCKET_ERROR)
		{
			dwNumberOfBytesRecvd = static_cast<DWORD>(nResult);

			if (dwNumberOfBytesRecvd == 0)
			{
				break;
			}
			else 
			{    
	    		dwNumberOfBytesLeft -= dwNumberOfBytesRecvd;
	    		lpszBuffer += dwNumberOfBytesRecvd;
			}                            
		}
		else 
		{
			// WSAETIMEDOUT check - if we get this error then retry the recieve.
		                                 
			if (::WSAGetLastError() == WSAETIMEDOUT && nRetry < s_nIORetries)
			{
				nRetry++;
			}
			else
			{
				break;
			}
		}
	}
	
	if (nResult != SOCKET_ERROR)
	{
		dwNumberOfBytesRecvd = dwLen - dwNumberOfBytesLeft;
	}                           
	
	return dwNumberOfBytesRecvd;
}

//////////////////////////////////////////////////////////////////////
// Static members.

bool CWSASocket::WSAStartup(const WORD wVersionRequested)
{
	if (!s_bStarted)
	{	
		namespaceMutex::CSingleLock lck(&m_csStarted, true);

		if (!s_bStarted)
		{	
			WSADATA WSAData;
			int nResult = 0;
			
			if ((nResult = ::WSAStartup(wVersionRequested, &WSAData)) != 0)
			{
				Error(__FILE__, __LINE__, _T("WSAStartup error"));
			}              
			else if 
				(
					LOBYTE(WSAData.wVersion) != LOBYTE(wVersionRequested) ||
        			HIBYTE(WSAData.wVersion) != HIBYTE(wVersionRequested)
				) 
			{        
				WSACleanup();
				Error(__FILE__, __LINE__, _T("WSAStartup error: requested version not supported"));
			}
			else
			{   
				FindProtocols();
				s_bStarted = true;
			}
		}
	}
		
	return s_bStarted;
}

bool CWSASocket::WSACleanup()
{                  
	bool bSuccess = false;

	if (s_bStarted)
	{
		namespaceMutex::CSingleLock lck(&m_csStarted, true);

		if (s_bStarted)
		{
			delete m_pProtocols;
			m_pProtocols = NULL;
			m_nProtocols = 0;

			if (::WSACleanup() == SOCKET_ERROR)
			{
				Error(__FILE__, __LINE__, _T("WSACleanup error"));
			}                     
			else
			{      
	   			s_bStarted = false;
	   			bSuccess = true;
			}
		}
	}
	
    return bSuccess;
}

bool CWSASocket::FindProtocols()
{
	bool bSuccess = false;
	int nResult = 0;
	DWORD dwBufLen = 0;

	delete m_pProtocols;
	m_pProtocols = NULL;
	m_nProtocols = 0;

	// First, have WSAEnumProtocols tell you how big a buffer you need.
	nResult = WSAEnumProtocols(NULL, m_pProtocols, &dwBufLen);
	if (SOCKET_ERROR != nResult)
	{
		Error(__FILE__, __LINE__, _T("WSAEnumProtocols error: should not have succeeded"));
	}
	else if (WSAENOBUFS != WSAGetLastError())
	{
		 // WSAEnumProtocols failed for some reason not relating to buffer
		 // size - also odd.
		 Error(__FILE__, __LINE__, _T("WSAEnumProtocols error"));
	}
	else
    {
		// WSAEnumProtocols failed for the "expected" reason. Therefore,
		// you need to allocate a buffer of the appropriate size.
		m_pProtocols = (WSAPROTOCOL_INFO*)new BYTE[dwBufLen];
		if (m_pProtocols)
		{
			// Now you can call WSAEnumProtocols again with the expectation
			// that it will succeed because you have allocated a big enough
			// buffer.
			nResult = WSAEnumProtocols(NULL, m_pProtocols, &dwBufLen);
			if (SOCKET_ERROR == nResult)
			{
				Error(__FILE__, __LINE__, _T("WSAEnumProtocols error"));
			}
			else
			{
				m_nProtocols = nResult;
				bSuccess = true;
			}
		}
	}

	return bSuccess;
}

void CWSASocket::Error(LPCSTR lpszFile, int nLine, LPCTSTR ppszFormatplusArgs, ...)
{
	Error(::WSAGetLastError(), lpszFile, nLine, &ppszFormatplusArgs);
}

void CWSASocket::Error(int nError, LPCSTR lpszFile, int nLine, LPCTSTR ppszFormatplusArgs, ...)
{
	Error(nError, lpszFile, nLine, &ppszFormatplusArgs);
}

void CWSASocket::Error(int nError, LPCSTR lpszFile, int nLine, LPCTSTR* ppszFormatplusArgs)
{
	if (nError == 0)
	{
		nError = ::WSAGetLastError();
	}

    _TCHAR chMsg[256];
    va_list pArg;

    va_start(pArg, *ppszFormatplusArgs);
    _vstprintf(chMsg, *ppszFormatplusArgs, pArg);
    va_end(pArg);

	throw CException(nError, lpszFile, nLine, _T("%s: %s"), chMsg, GetWSAErrorStr(nError));
}


const _TCHAR* CWSASocket::GetWSAErrorStr(int nWSAErrorCode)
{
	const _TCHAR* pszWSAErrorCode = NULL;

	if (nWSAErrorCode == NULL)
	{
		nWSAErrorCode = ::WSAGetLastError();
	}

	static const _TCHAR* pszWSAEINTR			= _T("WSAEINTR");
	static const _TCHAR* pszWSAEBADF			= _T("WSAEBADF");
	static const _TCHAR* pszWSAEACCES			= _T("WSAEACCES");
	static const _TCHAR* pszWSAEFAULT			= _T("WSAEFAULT");
	static const _TCHAR* pszWSAEINVAL			= _T("WSAEINVAL");
	static const _TCHAR* pszWSAEMFILE			= _T("WSAEMFILE");
	static const _TCHAR* pszWSAEWOULDBLOCK		= _T("WSAEWOULDBLOCK");
	static const _TCHAR* pszWSAEINPROGRESS		= _T("WSAEINPROGRESS");
	static const _TCHAR* pszWSAEALREADY			= _T("WSAEALREADY");
	static const _TCHAR* pszWSAENOTSOCK			= _T("WSAENOTSOCK");
	static const _TCHAR* pszWSAEDESTADDRREQ		= _T("WSAEDESTADDRREQ");
	static const _TCHAR* pszWSAEMSGSIZE			= _T("WSAEMSGSIZE");
	static const _TCHAR* pszWSAEPROTOTYPE		= _T("WSAEPROTOTYPE");
	static const _TCHAR* pszWSAENOPROTOOPT		= _T("WSAENOPROTOOPT");
	static const _TCHAR* pszWSAEPROTONOSUPPORT	= _T("WSAEPROTONOSUPPORT");
	static const _TCHAR* pszWSAESOCKTNOSUPPORT	= _T("WSAESOCKTNOSUPPORT");
	static const _TCHAR* pszWSAEOPNOTSUPP		= _T("WSAEOPNOTSUPP");
	static const _TCHAR* pszWSAEPFNOSUPPORT		= _T("WSAEPFNOSUPPORT");
	static const _TCHAR* pszWSAEAFNOSUPPORT		= _T("WSAEAFNOSUPPORT");
	static const _TCHAR* pszWSAEADDRINUSE		= _T("WSAEADDRINUSE");
	static const _TCHAR* pszWSAEADDRNOTAVAIL	= _T("WSAEADDRNOTAVAIL");
	static const _TCHAR* pszWSAENETDOWN			= _T("WSAENETDOWN");
	static const _TCHAR* pszWSAENETUNREACH		= _T("WSAENETUNREACH");
	static const _TCHAR* pszWSAENETRESET		= _T("WSAENETRESET");
	static const _TCHAR* pszWSAECONNABORTED		= _T("WSAECONNABORTED");
	static const _TCHAR* pszWSAECONNRESET		= _T("WSAECONNRESET");
	static const _TCHAR* pszWSAENOBUFS			= _T("WSAENOBUFS");
	static const _TCHAR* pszWSAEISCONN			= _T("WSAEISCONN");
	static const _TCHAR* pszWSAENOTCONN			= _T("WSAENOTCONN");
	static const _TCHAR* pszWSAESHUTDOWN		= _T("WSAESHUTDOWN");
	static const _TCHAR* pszWSAETOOMANYREFS		= _T("WSAETOOMANYREFS");
	static const _TCHAR* pszWSAETIMEDOUT		= _T("WSAETIMEDOUT");
	static const _TCHAR* pszWSAECONNREFUSED		= _T("WSAECONNREFUSED");
	static const _TCHAR* pszWSAELOOP			= _T("WSAELOOP");
	static const _TCHAR* pszWSAENAMETOOLONG		= _T("WSAENAMETOOLONG");
	static const _TCHAR* pszWSAEHOSTDOWN		= _T("WSAEHOSTDOWN");
	static const _TCHAR* pszWSAEHOSTUNREACH		= _T("WSAEHOSTUNREACH");
	static const _TCHAR* pszWSAENOTEMPTY		= _T("WSAENOTEMPTY");
	static const _TCHAR* pszWSAEPROCLIM			= _T("WSAEPROCLIM");
	static const _TCHAR* pszWSAEUSERS			= _T("WSAEUSERS");
	static const _TCHAR* pszWSAEDQUOT			= _T("WSAEDQUOT");
	static const _TCHAR* pszWSAESTALE			= _T("WSAESTALE");
	static const _TCHAR* pszWSAEREMOTE			= _T("WSAEREMOTE");
	static const _TCHAR* pszWSAEDISCON			= _T("WSAEDISCON");
	static const _TCHAR* pszWSASYSNOTREADY		= _T("WSASYSNOTREADY");
	static const _TCHAR* pszWSAVERNOTSUPPORTED	= _T("WSAVERNOTSUPPORTED");
	static const _TCHAR* pszWSANOTINITIALISED	= _T("WSANOTINITIALISED");
	static const _TCHAR* pszWSAHOST_NOT_FOUND	= _T("WSAHOST_NOT_FOUND");
	static const _TCHAR* pszWSATRY_AGAIN		= _T("WSATRY_AGAIN");
	static const _TCHAR* pszWSANO_RECOVERY		= _T("WSANO_RECOVERY");
	static const _TCHAR* pszWSANO_DATA			= _T("WSANO_DATA");

	switch (nWSAErrorCode)
	{
	case WSAEINTR:
		pszWSAErrorCode = pszWSAEINTR;
		break;
	case WSAEBADF:
		pszWSAErrorCode = pszWSAEBADF;
		break;
	case WSAEACCES:
		pszWSAErrorCode = pszWSAEACCES;
		break;
	case WSAEFAULT:
		pszWSAErrorCode = pszWSAEFAULT;
		break;
	case WSAEINVAL:
		pszWSAErrorCode = pszWSAEINVAL;
		break;
	case WSAEMFILE:
		pszWSAErrorCode = pszWSAEMFILE;
		break;
	case WSAEWOULDBLOCK:
		pszWSAErrorCode = pszWSAEWOULDBLOCK;
		break;
	case WSAEINPROGRESS:
		pszWSAErrorCode = pszWSAEINPROGRESS;
		break;
	case WSAEALREADY:
		pszWSAErrorCode = pszWSAEALREADY;
		break;
	case WSAENOTSOCK:
		pszWSAErrorCode = pszWSAENOTSOCK;
		break;
	case WSAEDESTADDRREQ:
		pszWSAErrorCode = pszWSAEDESTADDRREQ;
		break;
	case WSAEMSGSIZE:
		pszWSAErrorCode = pszWSAEMSGSIZE;
		break;
	case WSAEPROTOTYPE:
		pszWSAErrorCode = pszWSAEPROTOTYPE;
		break;
	case WSAENOPROTOOPT:
		pszWSAErrorCode = pszWSAENOPROTOOPT;
		break;
	case WSAEPROTONOSUPPORT:
		pszWSAErrorCode = pszWSAEPROTONOSUPPORT;
		break;
	case WSAESOCKTNOSUPPORT:
		pszWSAErrorCode = pszWSAESOCKTNOSUPPORT;
		break;
	case WSAEOPNOTSUPP:
		pszWSAErrorCode = pszWSAEOPNOTSUPP;
		break;
	case WSAEPFNOSUPPORT:
		pszWSAErrorCode = pszWSAEPFNOSUPPORT;
		break;
	case WSAEAFNOSUPPORT:
		pszWSAErrorCode = pszWSAEAFNOSUPPORT;
		break;
	case WSAEADDRINUSE:
		pszWSAErrorCode = pszWSAEADDRINUSE;
		break;
	case WSAEADDRNOTAVAIL:
		pszWSAErrorCode = pszWSAEADDRNOTAVAIL;
		break;
	case WSAENETDOWN:
		pszWSAErrorCode = pszWSAENETDOWN;
		break;
	case WSAENETUNREACH:
		pszWSAErrorCode = pszWSAENETUNREACH;
		break;
	case WSAENETRESET:
		pszWSAErrorCode = pszWSAENETRESET;
		break;
	case WSAECONNABORTED:
		pszWSAErrorCode = pszWSAECONNABORTED;
		break;
	case WSAECONNRESET:
		pszWSAErrorCode = pszWSAECONNRESET;
		break;
	case WSAENOBUFS:
		pszWSAErrorCode = pszWSAENOBUFS;
		break;
	case WSAEISCONN:
		pszWSAErrorCode = pszWSAEISCONN;
		break;
	case WSAENOTCONN:
		pszWSAErrorCode = pszWSAENOTCONN;
		break;
	case WSAESHUTDOWN:
		pszWSAErrorCode = pszWSAESHUTDOWN;
		break;
	case WSAETOOMANYREFS:
		pszWSAErrorCode = pszWSAETOOMANYREFS;
		break;
	case WSAETIMEDOUT:
		pszWSAErrorCode = pszWSAETIMEDOUT;
		break;
	case WSAECONNREFUSED:
		pszWSAErrorCode = pszWSAECONNREFUSED;
		break;
	case WSAELOOP:
		pszWSAErrorCode = pszWSAELOOP;
		break;
	case WSAENAMETOOLONG:
		pszWSAErrorCode = pszWSAENAMETOOLONG;
		break;
	case WSAEHOSTDOWN:
		pszWSAErrorCode = pszWSAEHOSTDOWN;
		break;
	case WSAEHOSTUNREACH:
		pszWSAErrorCode = pszWSAEHOSTUNREACH;
		break;
	case WSAENOTEMPTY:
		pszWSAErrorCode = pszWSAENOTEMPTY;
		break;
	case WSAEPROCLIM:
		pszWSAErrorCode = pszWSAEPROCLIM;
		break;
	case WSAEUSERS:
		pszWSAErrorCode = pszWSAEUSERS;
		break;
	case WSAEDQUOT:
		pszWSAErrorCode = pszWSAEDQUOT;
		break;
	case WSAESTALE:
		pszWSAErrorCode = pszWSAESTALE;
		break;
	case WSAEREMOTE:
		pszWSAErrorCode = pszWSAEREMOTE;
		break;
	case WSAEDISCON:
		pszWSAErrorCode = pszWSAEDISCON;
		break;
	case WSASYSNOTREADY:
		pszWSAErrorCode = pszWSASYSNOTREADY;
		break;
	case WSAVERNOTSUPPORTED:
		pszWSAErrorCode = pszWSAVERNOTSUPPORTED;
		break;
	case WSANOTINITIALISED:
		pszWSAErrorCode = pszWSANOTINITIALISED;
		break;
	case WSAHOST_NOT_FOUND:
		pszWSAErrorCode = pszWSAHOST_NOT_FOUND;
		break;
	case WSATRY_AGAIN:
		pszWSAErrorCode = pszWSATRY_AGAIN;
		break;
	case WSANO_RECOVERY:
		pszWSAErrorCode = pszWSANO_RECOVERY;
		break;
	case WSANO_DATA:
		pszWSAErrorCode = pszWSANO_DATA;
		break;
	}

	return pszWSAErrorCode;
}
