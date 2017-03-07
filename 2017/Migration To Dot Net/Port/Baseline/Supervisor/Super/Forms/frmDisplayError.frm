VERSION 5.00
Begin VB.Form frmDisplayError 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmDisplayError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   1575
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdExtra 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      Picture         =   "frmDisplayError.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   180
      Width           =   555
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmDisplayError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmDisplayError
' Description   :   Used for displaying errors to the user. There is the general error
'                   message text and the "extra" error text which can be viewed when the
'                   extra button is enabled (when SetExtra() has been called to set an
'                   extra value)
'
' Change history
' Prog      Date        Description
' DJP       29/11/00    Set focus when form is activated.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GD APPLIED BMIDS00010
' STB       16/05/02    SYS4620 Adjusted the colour profile so error messages won't be made invisible.


Option Explicit
Private m_sExtra As String
Public Sub SetExtra(sExtra As String)
    m_sExtra = sExtra
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExtra_Click()
    MsgBox m_sExtra, vbInformation
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
' DJP       29/11/00    Set focus when form is activated.
Private Sub Form_Activate()
    cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    EndWaitCursor
    Caption = App.Title
    If Len(m_sExtra) > 0 Then
        Me.cmdExtra.Enabled = True
    Else
        Me.cmdExtra.Enabled = False
    End If
    

    
End Sub
