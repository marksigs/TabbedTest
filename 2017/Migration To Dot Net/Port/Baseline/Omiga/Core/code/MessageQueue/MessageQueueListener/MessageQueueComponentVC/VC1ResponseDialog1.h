///////////////////////////////////////////////////////////////////////////////
//	FILE:			VC1ResponseDialog1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
///////////////////////////////////////////////////////////////////////////////

#ifndef __VC1RESPONSEDIALOG1_H_
#define __VC1RESPONSEDIALOG1_H_

#include "resource.h"       // main symbols
#include "MessageQueueComponentVC.h"
#include <atlhost.h>

///////////////////////////////////////////////////////////////////////////////

class CVC1ResponseDialog1 : 
	public CAxDialogImpl<CVC1ResponseDialog1>
{
public:
	CVC1ResponseDialog1(BSTR& rbstrMessage) :
		m_rbstrMessage(rbstrMessage),
		m_lMESSQ_RESP(MESSQ_RESP_SUCCESS)
	{
	}

	~CVC1ResponseDialog1()
	{
	}

	enum { IDD = IDD_VC1RESPONSEDIALOG1 };


BEGIN_MSG_MAP(CVC1ResponseDialog1)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_SUCCESS, OnMessqRespSuccess)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_RETRY_NOW, OnMessqRespRetryNow)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_RETRY_LATER, OnMessqRetryLater)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_RETRY_MOVE_MESSAGE, OnMessqMoveMessage)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_STALL_COMPONENT, OnMessqStallComponent)
	COMMAND_ID_HANDLER(IDC_MESSQ_RESP_STALL_QUEUE, OnMessqRespStallQueue)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		::SetWindowTextW(GetDlgItem(IDC_MESSAGE), m_rbstrMessage);
		return 1;  // Let the system set the focus
	}

	LRESULT OnMessqRespSuccess(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_SUCCESS;
		EndDialog(wID);
		return 0;
	}

	LRESULT OnMessqRespRetryNow(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_RETRY_NOW;
		EndDialog(wID);
		return 0;
	}

	LRESULT OnMessqRetryLater(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_RETRY_LATER;
		EndDialog(wID);
		return 0;
	}

	LRESULT OnMessqMoveMessage(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_RETRY_MOVE_MESSAGE;
		EndDialog(wID);
		return 0;
	}

	LRESULT OnMessqStallComponent(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_STALL_COMPONENT;
		EndDialog(wID);
		return 0;
	}

	LRESULT OnMessqRespStallQueue(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		m_lMESSQ_RESP = MESSQ_RESP_STALL_QUEUE;
		EndDialog(wID);
		return 0;
	}

// Dialog Data	
public:
	BSTR& m_rbstrMessage;
	long m_lMESSQ_RESP;
};

#endif //__VC1RESPONSEDIALOG1_H_
