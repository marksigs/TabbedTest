VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "Logging"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   7200
      Width           =   1455
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmText
' Description   : Displays the Supervisor log file
'
' Change history
' Prog      Date        Description
' DJP       11/03/03    Created for BM0316
' TW        02/01/2007  EP2_640 - Form not displaying the whole of the log if > 32k
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Sub cmdOK_Click()
    Hide
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    
    ReadLog
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub ReadLog()
Dim fs As FileSystemObject
Dim strFileName As String
Dim strWork As String
Dim arrLines

Dim F As Integer
Dim X As Long
    
    On Error GoTo Failed
    
    Set fs = New FileSystemObject
    strFileName = App.Path & "\" & PROMOTION_LOG_FILE
    List1.Clear
    If fs.FileExists(strFileName) Then
        strWork = Space$(FileLen(strFileName))
        F = FreeFile()
        Open strFileName For Binary Access Read As F
        Get F, , strWork
        Close F
        arrLines = Split(strWork, vbCrLf)
        For X = 0 To UBound(arrLines, 1)
            List1.AddItem arrLines(X)
        Next X
    
    End If
    List1.Visible = True
    Set fs = Nothing
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Form_Resize()
' TW 02/01/2007 EP2_640 - Added resizing capability to form
Dim W As Integer
Dim H As Integer
    W = Me.Width - 360
    H = Me.Height - 1460
    If H > 0 And W > 1 Then
        List1.Move 120, 240, W, H
        cmdOK.Top = Me.Height - cmdOK.Height - 540
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseLogging
End Sub
