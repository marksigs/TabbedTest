VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSGProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_frmParent As Form
Private m_sTitle As String

Public Sub ClearProgress()
    
    m_frmParent.lblTitle.Caption = ""

Failed:
End Sub
Public Sub StartProgress(frmParent As Form, sTitle)
    On Error GoTo Failed
    frmParent.lblTitle.Caption = sTitle
    frmParent.Refresh
    Set m_frmParent = frmParent
Failed:

End Sub
