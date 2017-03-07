VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmLegalMaintenance 
   Caption         =   "Legal Rep. Maintenance"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   10035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Bank Details"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8595
      TabIndex        =   54
      Top             =   5670
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&PAF"
      Height          =   315
      Left            =   5220
      TabIndex        =   4
      Top             =   540
      Width           =   615
   End
   Begin VB.OptionButton optHeadOffice 
      Caption         =   "Yes"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   3900
      Width           =   855
   End
   Begin VB.OptionButton optHeadOffice 
      Caption         =   "No"
      Height          =   315
      Index           =   1
      Left            =   2700
      TabIndex        =   2
      Top             =   3900
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5985
      TabIndex        =   1
      Top             =   5670
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7290
      TabIndex        =   0
      Top             =   5670
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   540
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
   Begin MSGOCX.MSGComboBox cboMaintenance 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   1380
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   540
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   1800
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   5
      Left            =   1800
      TabIndex        =   11
      Top             =   2220
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   6
      Left            =   1800
      TabIndex        =   12
      Top             =   2640
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   7
      Left            =   1800
      TabIndex        =   13
      Top             =   3060
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   8
      Left            =   1800
      TabIndex        =   14
      Top             =   3480
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
   Begin MSGOCX.MSGComboBox cboMaintenance 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   5160
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   9
      Left            =   7500
      TabIndex        =   16
      Top             =   120
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   10
      Left            =   7500
      TabIndex        =   17
      Top             =   540
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   11
      Left            =   7500
      TabIndex        =   18
      Top             =   960
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   12
      Left            =   7500
      TabIndex        =   19
      Top             =   1380
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   13
      Left            =   7500
      TabIndex        =   20
      Top             =   1800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   14
      Left            =   7500
      TabIndex        =   21
      Top             =   2220
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   15
      Left            =   7440
      TabIndex        =   22
      Top             =   4140
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   16
      Left            =   7440
      TabIndex        =   23
      Top             =   4560
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   17
      Left            =   7440
      TabIndex        =   24
      Top             =   4980
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
   Begin MSGOCX.MSGComboBox cboMaintenance 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   46
      Top             =   4320
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   18
      Left            =   1800
      TabIndex        =   48
      Top             =   4740
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSGOCX.MSGComboBox cboMaintenance 
      Height          =   315
      Index           =   3
      Left            =   7500
      TabIndex        =   50
      Top             =   2700
      Width           =   1395
      _ExtentX        =   2461
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   19
      Left            =   7500
      TabIndex        =   52
      Top             =   3120
      Width           =   435
      _ExtentX        =   767
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
   Begin VB.Label lblMaintenance 
      Caption         =   "No. of Partners"
      Height          =   255
      Index           =   24
      Left            =   6000
      TabIndex        =   53
      Top             =   3180
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Country of Practice"
      Height          =   435
      Index           =   23
      Left            =   6000
      TabIndex        =   51
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Indemnity Ins. Expiry"
      Height          =   435
      Index           =   22
      Left            =   300
      TabIndex        =   49
      Top             =   4740
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Type"
      Height          =   255
      Index           =   21
      Left            =   300
      TabIndex        =   47
      Top             =   4380
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Panel Number"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   45
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   44
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Post Code"
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   43
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Building Name && No."
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   42
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Street"
      Height          =   255
      Index           =   4
      Left            =   300
      TabIndex        =   41
      Top             =   1860
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "District"
      Height          =   255
      Index           =   5
      Left            =   300
      TabIndex        =   40
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Town"
      Height          =   255
      Index           =   6
      Left            =   300
      TabIndex        =   39
      Top             =   2700
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "County"
      Height          =   255
      Index           =   7
      Left            =   300
      TabIndex        =   38
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Country"
      Height          =   255
      Index           =   8
      Left            =   300
      TabIndex        =   37
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Head Office"
      Height          =   255
      Index           =   9
      Left            =   300
      TabIndex        =   36
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Free Pay Method"
      Height          =   255
      Index           =   10
      Left            =   300
      TabIndex        =   35
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Telephone"
      Height          =   255
      Index           =   11
      Left            =   6000
      TabIndex        =   34
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Fax"
      Height          =   255
      Index           =   12
      Left            =   6000
      TabIndex        =   33
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "E-Mail"
      Height          =   255
      Index           =   13
      Left            =   6000
      TabIndex        =   32
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "DX Number"
      Height          =   255
      Index           =   14
      Left            =   6000
      TabIndex        =   31
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Active From"
      Height          =   255
      Index           =   15
      Left            =   6000
      TabIndex        =   30
      Top             =   1860
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Active To"
      Height          =   255
      Index           =   16
      Left            =   6000
      TabIndex        =   29
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Contact Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   5940
      TabIndex        =   28
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Name"
      Height          =   255
      Index           =   18
      Left            =   5940
      TabIndex        =   27
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Telephone"
      Height          =   255
      Index           =   19
      Left            =   5940
      TabIndex        =   26
      Top             =   4620
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Extension"
      Height          =   255
      Index           =   20
      Left            =   5940
      TabIndex        =   25
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmLegalMaintenance"
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

Private Sub Command2_Click()
    frmBankAccDetails.Show vbModal, Me
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
