VERSION 5.00

Object = "{2728C331-6208-11D3-AD6C-0060087A1BF0}#1.0#0"; "msgtextedit.ocx"

Begin VB.Form frmEditLegalFees 
   Caption         =   "Legal Fees"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameLegalFees 
      Caption         =   "Fee Set Bands"
      Height          =   3375
      Left            =   60
      TabIndex        =   15
      Top             =   2640
      Width           =   5775
      Begin MSGOCX.MSGDataGrid MSGDataGrid1 
         Height          =   2715
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4789
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSGOCX.MSGComboBox cboFeeType 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   6300
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6300
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtLegalFees 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtLegalFees 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtLegalFees 
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtLegalFees 
      Height          =   315
      Index           =   3
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGComboBox cboApplicationType 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Application Type"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   1395
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Fee Type"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   1395
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "End Date"
      Height          =   255
      Index           =   4
      Left            =   4020
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Set Description"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label lblLegalFees 
      Caption         =   "Lender Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditLegalFees"
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
'    MSGFlexGrid1.Rows = 2
'    MSGFlexGrid1.FormatString = "|^      Start Date       |^      End Date      |^      Type of Application      |^      From £      |^      To £      |^      Legal Fee Amount £      ;   "
    
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
