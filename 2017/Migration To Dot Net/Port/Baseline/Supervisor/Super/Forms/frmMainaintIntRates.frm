VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmMainaintIntRates 
   Caption         =   "Maintain Interest Rates"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Another"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1740
      TabIndex        =   16
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   3720
      Width           =   1275
   End
   Begin MSGOCX.MSGComboBox cboRateType 
      Height          =   315
      Left            =   2100
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
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
      Text            =   ""
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   240
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   1
      Left            =   2100
      TabIndex        =   2
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   2
      Left            =   2100
      TabIndex        =   4
      Top             =   1080
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   3
      Left            =   2100
      TabIndex        =   6
      Top             =   1980
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   5
      Left            =   2100
      TabIndex        =   8
      Top             =   2400
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin MSGOCX.MSGEditBox txtMaintIntRates 
      Height          =   315
      Index           =   6
      Left            =   2100
      TabIndex        =   10
      Top             =   2820
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Months"
      Height          =   255
      Index           =   8
      Left            =   3300
      TabIndex        =   14
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Rate Type"
      Height          =   255
      Index           =   7
      Left            =   300
      TabIndex        =   12
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Ceiling Rate"
      Height          =   255
      Index           =   6
      Left            =   300
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Floor Rate"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   9
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Interest Rate"
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Rate End Date"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Rate Period"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblMaintIntRates 
      Caption         =   "Sequence Number"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "frmMainaintIntRates"
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
