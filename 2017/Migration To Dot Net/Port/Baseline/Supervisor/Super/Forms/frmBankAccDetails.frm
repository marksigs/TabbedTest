VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"
Begin VB.Form frmBankAccDetails 
   Caption         =   "Bank Account Details"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox cboBankAccDetails 
      Height          =   315
      Index           =   0
      Left            =   1260
      ScaleHeight     =   255
      ScaleWidth      =   1875
      TabIndex        =   17
      Top             =   2340
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4515
      TabIndex        =   16
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5835
      TabIndex        =   15
      Top             =   2940
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   2
      Left            =   1260
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      _ExtentX        =   2037
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   2
      Top             =   1020
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   4
      Left            =   1260
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   5
      Left            =   1260
      TabIndex        =   6
      Top             =   1860
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   6
      Left            =   4380
      TabIndex        =   8
      Top             =   600
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   7
      Left            =   4380
      TabIndex        =   9
      Top             =   1020
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   8
      Left            =   4380
      TabIndex        =   10
      Top             =   1440
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtBankAccDetails 
      Height          =   315
      Index           =   9
      Left            =   4380
      TabIndex        =   11
      Top             =   1860
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin VB.PictureBox cboBankAccDetails 
      Height          =   315
      Index           =   1
      Left            =   4380
      ScaleHeight     =   255
      ScaleWidth      =   1875
      TabIndex        =   18
      Top             =   2340
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Currency"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "Second Bank Account"
      Height          =   315
      Index           =   1
      Left            =   4320
      TabIndex        =   13
      Top             =   180
      Width           =   2055
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "First Bank Account"
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   12
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "A/C Name"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "A/C No."
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "Bank Name"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label lblBankAccDetails 
      Caption         =   "Sort Code"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1035
   End
End
Attribute VB_Name = "frmBankAccDetails"
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
