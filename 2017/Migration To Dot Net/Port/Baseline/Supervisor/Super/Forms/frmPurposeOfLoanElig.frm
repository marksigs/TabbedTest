VERSION 5.00

Object = "{2728C331-6208-11D3-AD6C-0060087A1BF0}#1.0#0"; "msgtextedit.ocx"
Begin VB.Form frmPurposeOfLoanElig 
   Caption         =   "Purpose of Loan Eligilibility"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9015
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGEditBox txtPurposeOfLoan 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4380
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   4380
      Width           =   1215
   End
   Begin MSGOCX.MSGHorizontalSwapList MSGHorizontalSwapList1 
      Height          =   3510
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8235
      _ExtentX        =   13996
      _ExtentY        =   6191
   End
   Begin VB.Label lblPurposeOfLoan 
      Caption         =   "Product Code"
      Height          =   255
      Index           =   0
      Left            =   540
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmPurposeOfLoanElig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub Form_Load()
    MSGHorizontalSwapList1.SetFirstTitle "Purpose of Loan List"
    MSGHorizontalSwapList1.SetSecondTitle "Purpose of Loan Eligibility"
    SetHeaders

End Sub

Private Sub SetHeaders()
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim line As Collection
    
    lvHeaders.nWidth = 100
    lvHeaders.sName = "A Top Heading"
    headers.Add lvHeaders
    MSGHorizontalSwapList1.SetFirstColumnHeaders headers
    
    Set headers = New Collection
    lvHeaders.sName = "Another Top Heading"
    headers.Add lvHeaders
    MSGHorizontalSwapList1.SetSecondColumnHeaders headers
    Set line = New Collection
    line.Add "Some Text"

    MSGHorizontalSwapList1.AddLineFirst line
    Set line = New Collection
    line.Add "Some More Text"
    MSGHorizontalSwapList1.AddLineFirst line
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
