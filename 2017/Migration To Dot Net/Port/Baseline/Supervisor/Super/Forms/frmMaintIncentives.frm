VERSION 5.00
Object = "{2728C331-6208-11D3-AD6C-0060087A1BF0}#1.0#0"; "msgtextedit.ocx"
Object = "{945A1F21-620A-11D3-AD6C-0060087A1BF0}#1.0#0"; "MSGFlexGrid.ocx"
Begin VB.Form frmMaintIncentives 
   Caption         =   "Maintain Incentives"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   4140
      Width           =   1215
   End
   Begin MSGFlex.MSGFlexGrid MSGFlexIncentives 
      Height          =   3195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5636
      FormatString    =   $"frmMaintIncentives.frx":0000
   End
   Begin MSGOCX.MSGEditBox txtMaintIncentives 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   180
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin VB.Label lblMaintIncentives 
      Caption         =   "Set Type"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMaintIncentives"
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

Private Sub Form_Load()
    
    MSGFlexIncentives.Rows = 2
    MSGFlexIncentives.Cols = 5
    MSGFlexIncentives.FormatString = "|^       Set Type       |^             Description            |^    Amount    |^   %age   |^    %age Max.     ;   "

End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
