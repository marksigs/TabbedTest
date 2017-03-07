///////////////////////////////////////////////////////////////////////////////
//	FILE:			VC2ResponseDialog1.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      15/11/00    Initial version
///////////////////////////////////////////////////////////////////////////////

#ifndef __VC2RESPONSEDIALOG1_H_
#define __VC2RESPONSEDIALOG1_H_

#include "resource.h"       // main symbols
#include "MessageQueueComponentVC.h"
#include <atlhost.h>

///////////////////////////////////////////////////////////////////////////////

class CVC2ResponseDialog1 : 
	public CAxDialogImpl<CVC2ResponseDialog1>
{
public:
	CVC2ResponseDialog1(BSTR& rbstrConfig, BSTR& rbstrMessage) :
		m_rbstrConfig(rbstrConfig),
		m_rbstrMessage(rbstrMessage),
		m_lMESSQ_RESP(MESSQ_RESP_SUCCESS)
	{
	}

	~CVC2ResponseDialog1()
	{
	}

	enum { IDD = IDD_VC2RESPONSEDIALOG1 };


BEGIN_MSG_MAP(CVC2ResponseDialog1)
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
		::SetWindowTextW(GetDlgItem(IDC_CONFIG), m_rbstrConfig);
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
	BSTR& m_rbstrConfig;
	BSTR& m_rbstrMessage;
	long m_lMESSQ_RESP;
};

#endif //__VC2RESPONSEDIALOG1_H_
