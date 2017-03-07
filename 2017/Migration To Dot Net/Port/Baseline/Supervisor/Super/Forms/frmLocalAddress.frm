VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmLocalAddress 
   Caption         =   "Local Name & Address Maintenance"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBankDetails 
      Caption         =   "&Bank Details"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8505
      TabIndex        =   46
      Top             =   5040
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   45
      Top             =   5040
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5895
      TabIndex        =   44
      Top             =   5040
      Width           =   1230
   End
   Begin VB.OptionButton optHeadOffice 
      Caption         =   "No"
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   21
      Top             =   4020
      Width           =   855
   End
   Begin VB.OptionButton optHeadOffice 
      Caption         =   "Yes"
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   20
      Top             =   4020
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&PAF"
      Height          =   315
      Left            =   5100
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   480
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
      Left            =   1680
      TabIndex        =   1
      Top             =   60
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
      Left            =   1680
      TabIndex        =   5
      Top             =   900
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
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
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
      Left            =   4200
      TabIndex        =   8
      Top             =   480
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
      Left            =   1680
      TabIndex        =   11
      Top             =   1740
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
      Left            =   1680
      TabIndex        =   13
      Top             =   2160
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
      Left            =   1680
      TabIndex        =   15
      Top             =   2580
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
      Left            =   1680
      TabIndex        =   17
      Top             =   3000
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
      Left            =   1680
      TabIndex        =   19
      Top             =   3420
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
      Left            =   1680
      TabIndex        =   24
      Top             =   4500
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
      Left            =   7380
      TabIndex        =   26
      Top             =   60
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
      Left            =   7380
      TabIndex        =   28
      Top             =   480
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
      Left            =   7380
      TabIndex        =   30
      Top             =   900
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
      Left            =   7380
      TabIndex        =   32
      Top             =   1320
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
      Left            =   7380
      TabIndex        =   34
      Top             =   1740
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
      Left            =   7380
      TabIndex        =   36
      Top             =   2160
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
      Left            =   7320
      TabIndex        =   39
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
      Index           =   16
      Left            =   7320
      TabIndex        =   41
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
   Begin MSGOCX.MSGEditBox txtMaintenance 
      Height          =   315
      Index           =   17
      Left            =   7320
      TabIndex        =   43
      Top             =   3900
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
   Begin VB.Label lblMaintenance 
      Caption         =   "Extension"
      Height          =   255
      Index           =   20
      Left            =   5820
      TabIndex        =   42
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Telephone"
      Height          =   255
      Index           =   19
      Left            =   5820
      TabIndex        =   40
      Top             =   3540
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Name"
      Height          =   255
      Index           =   18
      Left            =   5820
      TabIndex        =   38
      Top             =   3120
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
      Left            =   5820
      TabIndex        =   37
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Active To"
      Height          =   255
      Index           =   16
      Left            =   5880
      TabIndex        =   35
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Active From"
      Height          =   255
      Index           =   15
      Left            =   5880
      TabIndex        =   33
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "DX Number"
      Height          =   255
      Index           =   14
      Left            =   5880
      TabIndex        =   31
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "E-Mail"
      Height          =   255
      Index           =   13
      Left            =   5880
      TabIndex        =   29
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Fax"
      Height          =   255
      Index           =   12
      Left            =   5880
      TabIndex        =   27
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Telephone"
      Height          =   255
      Index           =   11
      Left            =   5880
      TabIndex        =   25
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Free Pay Method"
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Head Office"
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   22
      Top             =   4020
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Country"
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "County"
      Height          =   255
      Index           =   7
      Left            =   180
      TabIndex        =   16
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Town"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "District"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   12
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Street"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Building Name && No."
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Post Code"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label lblMaintenance 
      Caption         =   "Type"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLocalAddress"
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

Private Sub cmdBankDetails_Click()
    frmBankAccDetails.Show vbModal, Me
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
