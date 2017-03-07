VERSION 5.00
Object = "{2728C331-6208-11D3-AD6C-0060087A1BF0}#1.0#0"; "msgtextedit.ocx"
Begin VB.Form frmEditStampDuty 
   Caption         =   "Stamp Duty"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   9945
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   4140
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtStampDuty 
      Height          =   315
      Index           =   1
      Left            =   1620
      TabIndex        =   7
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtStampDuty 
      Height          =   315
      Index           =   2
      Left            =   1620
      TabIndex        =   8
      Top             =   540
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtStampDuty 
      Height          =   315
      Index           =   3
      Left            =   5340
      TabIndex        =   9
      Top             =   120
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtStampDuty 
      Height          =   315
      Index           =   4
      Left            =   5340
      TabIndex        =   0
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin VB.Label lblStampDuty 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblStampDuty 
      Caption         =   "Set ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label lblStampDuty 
      Caption         =   "End Date"
      Height          =   255
      Index           =   4
      Left            =   4020
      TabIndex        =   4
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label lblStampDuty 
      Caption         =   "Set Description"
      Height          =   255
      Index           =   3
      Left            =   4020
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditStampDuty"
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

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
