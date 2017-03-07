VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pbProgress 
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   820
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   0
   End
   Begin VB.Label lbltext 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Workfile:      frmMessage.bas
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:
'
'   Modeless forms in an in-process ActiveX DLL are not possible from
'   Internet Explorer (see MSDN Q176468), therefore to display a wait prompt when
'   printing we have to use a modal form. Because it is modal, we have to call the
'   actual printing function from this form, in order that the printing occurs whilst
'   the form and the "Please wait" message are displayed.
'
'   In addition a timer is used to ensure the form window is painted first, before
'   kicking off the printing.
'
'Dependencies:  None
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DJB    08/05/02 Created.
'------------------------------------------------------------------------------------------

Option Explicit

Private Const gstrObjectName = "frmMessage.bas"

Private m_messageOp As IMessageOp
Private m_bSuccess As Boolean

Private Sub Form_Initialize()
    Set m_messageOp = Nothing
    lbltext.Caption = "Please wait..."
    pbProgress.Value = 0
    m_bSuccess = False
End Sub

Private Sub Form_Load()
   
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    tmrMain.Interval = 10
    tmrMain.Enabled = True
            
End Sub

Private Sub tmrMain_Timer()
On Error GoTo ExitPoint

    Const strFunctionName As String = "tmrMain_Timer"

    tmrMain.Enabled = False
    m_bSuccess = False
    
    MousePointer = vbHourglass
    
    m_bSuccess = m_messageOp.Execute
    
ExitPoint:
    Set m_messageOp = Nothing
    MousePointer = vbDefault
    Unload Me
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False, True
End Sub

Public Property Let MessageOp(ByVal vNewValue As Variant)
    Set m_messageOp = vNewValue
End Property

Public Property Let Message(ByVal vNewValue As Variant)
    lbltext.Caption = vNewValue
End Property

Public Property Let Progress(ByVal vNewValue As Variant)
    pbProgress.Value = vNewValue
End Property

Public Property Get Success() As Boolean
    Success = m_bSuccess
End Property

