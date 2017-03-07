VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmEditOtherFees 
   Caption         =   "Other Fees"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   13
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   12
      Top             =   3180
      Width           =   1215
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   660
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Top             =   1500
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtOtherFees 
      Height          =   315
      Index           =   5
      Left            =   1800
      TabIndex        =   10
      Top             =   2340
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "Percentage"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   11
      Top             =   2460
      Width           =   1275
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "Amount"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   9
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "Fee Type"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "End Date"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label lblOtherFees 
      Caption         =   "Fee Name"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   360
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditOtherFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Hide
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
