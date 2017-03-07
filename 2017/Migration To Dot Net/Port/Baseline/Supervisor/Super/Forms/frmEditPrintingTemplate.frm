VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPrintingTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Printing Template"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmEditPrintingTemplate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   8985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   8180
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7620
      TabIndex        =   22
      Top             =   8180
      Width           =   1215
   End
   Begin VB.Frame fraStages 
      Height          =   2440
      Left            =   120
      TabIndex        =   37
      Top             =   5660
      Width           =   8760
      Begin MSGOCX.MSGHorizontalSwapList swapStage 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4048
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   5655
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   8760
      Begin VB.Frame fraPrecreateForPack 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6240
         TabIndex        =   64
         Top             =   5190
         Width           =   1455
         Begin VB.OptionButton optPrecreate 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optPrecreate 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   65
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame fraPrintToWeb 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6240
         TabIndex        =   49
         Top             =   3945
         Width           =   1395
         Begin VB.OptionButton optPrintToWeb 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   51
            Top             =   20
            Width           =   615
         End
         Begin VB.OptionButton optPrintToWeb 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   50
            Top             =   20
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame fmePreventEdit 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1800
         TabIndex        =   46
         Top             =   3930
         Width           =   1440
         Begin VB.OptionButton optPreventEdit 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   48
            Top             =   20
            Width           =   615
         End
         Begin VB.OptionButton optPreventEdit 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   47
            Top             =   20
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin MSGOCX.MSGComboBox cboPrinterDestination 
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   3140
         Width           =   2400
         _ExtentX        =   4233
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
      Begin VB.Frame fraViewBeforePrint 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6240
         TabIndex        =   44
         Top             =   3500
         Width           =   1395
         Begin VB.OptionButton optViewPrint 
            Caption         =   "No"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   765
            TabIndex        =   19
            Top             =   90
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optViewPrint 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   18
            Top             =   90
            Width           =   615
         End
      End
      Begin VB.Frame fraEditBeforePrint 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1800
         TabIndex        =   43
         Top             =   3500
         Width           =   1440
         Begin VB.OptionButton optEditPrint 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   780
            TabIndex        =   17
            Top             =   90
            Width           =   615
         End
         Begin VB.OptionButton optEditPrint 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   16
            Top             =   90
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin MSGOCX.MSGMulti txtPrintMethodName 
         Height          =   315
         Left            =   6240
         TabIndex        =   13
         Top             =   2700
         Width           =   2400
         _ExtentX        =   4233
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
         Text            =   ""
         MaxLength       =   80
      End
      Begin VB.Frame fraPrintMenuAccess 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   6240
         TabIndex        =   39
         Top             =   2262
         Width           =   1410
         Begin VB.OptionButton optPrintMenu 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   10
            Top             =   90
            Width           =   615
         End
         Begin VB.OptionButton optPrintMenu 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   765
            TabIndex        =   11
            Top             =   90
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.CheckBox chkInactive 
         Enabled         =   0   'False
         Height          =   330
         Left            =   6240
         TabIndex        =   5
         Top             =   1402
         Width           =   525
      End
      Begin VB.Frame fraCustInd 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   1800
         TabIndex        =   38
         Top             =   2255
         Width           =   1875
         Begin VB.OptionButton optCustInd 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   780
            TabIndex        =   9
            Top             =   90
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optCustInd 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   8
            Top             =   90
            Width           =   615
         End
      End
      Begin MSGOCX.MSGComboBox cboDocumentType 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   615
         Width           =   2400
         _ExtentX        =   4233
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
         Mandatory       =   -1  'True
         Text            =   ""
      End
      Begin MSGOCX.MSGEditBox txtDetail 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
         Top             =   210
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
         BackColor       =   16777215
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   10
      End
      Begin MSGOCX.MSGEditBox txtDPSTemplateID 
         Height          =   315
         Left            =   6240
         TabIndex        =   1
         Top             =   210
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   10
      End
      Begin MSGOCX.MSGEditBox txtDetail 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   1845
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         PromptInclude   =   0   'False
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
         MaxLength       =   5
      End
      Begin MSGOCX.MSGEditBox txtDetail 
         Height          =   315
         Index           =   4
         Left            =   6240
         TabIndex        =   3
         Top             =   615
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   50
      End
      Begin MSGOCX.MSGComboBox cboMinRoleLevel 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   2400
         _ExtentX        =   4233
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
         Mandatory       =   -1  'True
         Text            =   ""
      End
      Begin MSGOCX.MSGComboBox cboRecipientType 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   2700
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSGOCX.MSGEditBox txtDetail 
         Height          =   315
         Index           =   6
         Left            =   6240
         TabIndex        =   7
         Top             =   1845
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         TextType        =   5
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   5
      End
      Begin MSGOCX.MSGMulti txtRemoteLocation 
         Height          =   315
         Left            =   6240
         TabIndex        =   15
         Top             =   3140
         Width           =   2400
         _ExtentX        =   4233
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
         Text            =   ""
         MaxLength       =   100
      End
      Begin MSGOCX.MSGComboBox cboDocDelType 
         Height          =   315
         Left            =   1800
         TabIndex        =   56
         Top             =   4320
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSGOCX.MSGMulti txtDescription 
         Height          =   330
         Left            =   1800
         TabIndex        =   57
         Top             =   1020
         Width           =   6860
         _ExtentX        =   12091
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         MaxLength       =   255
      End
      Begin MSGOCX.MSGComboBox cboFirstPagePrinterTray 
         Height          =   315
         Left            =   1800
         TabIndex        =   59
         Top             =   4740
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSGOCX.MSGComboBox cboOtherPagesPrinterTray 
         Height          =   315
         Left            =   6240
         TabIndex        =   60
         Top             =   4740
         Width           =   2400
         _ExtentX        =   4233
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
      Begin MSGOCX.MSGComboBox cboGeminiPrintMode 
         Height          =   315
         Left            =   1800
         TabIndex        =   63
         Top             =   5160
         Width           =   2415
         _ExtentX        =   4260
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
         Mandatory       =   -1  'True
         Text            =   ""
      End
      Begin MSGOCX.MSGComboBox cboEmailRecipientType 
         Height          =   315
         Left            =   6240
         TabIndex        =   67
         Top             =   3140
         Width           =   2400
         _ExtentX        =   4233
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
      Begin VB.Label lblCreatForPack 
         AutoSize        =   -1  'True
         Caption         =   "Pre-create for Pack"
         Height          =   195
         Left            =   4680
         TabIndex        =   62
         Top             =   5220
         Width           =   1380
      End
      Begin VB.Label lblGeminiPrintMode 
         AutoSize        =   -1  'True
         Caption         =   "Gemini Print Mode"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   5220
         Width           =   1290
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   1088
         Width           =   795
      End
      Begin VB.Label lblDocDelType 
         Caption         =   "Document Delivery Type"
         Height          =   420
         Left            =   120
         TabIndex        =   55
         Top             =   4267
         Width           =   1395
      End
      Begin VB.Label lblFirstPrintTray 
         AutoSize        =   -1  'True
         Caption         =   "First Page Printer Tray"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   4800
         Width           =   1560
      End
      Begin VB.Label lblOtherPrintTray 
         AutoSize        =   -1  'True
         Caption         =   "Other Pages Printer Tray"
         Height          =   390
         Left            =   4680
         TabIndex        =   53
         Top             =   4702
         Width           =   1500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSendPrintToWeb 
         Caption         =   "Send Print to Web"
         Height          =   195
         Left            =   4680
         TabIndex        =   52
         Top             =   3990
         Width           =   1305
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Prevent Edit in DMS"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   45
         Top             =   3970
         Width           =   1440
      End
      Begin VB.Label lblViewBeforePrinting 
         AutoSize        =   -1  'True
         Caption         =   "View Before Printing"
         Height          =   195
         Left            =   4680
         TabIndex        =   42
         Top             =   3605
         Width           =   1425
      End
      Begin VB.Label lblEditBeforePrinting 
         AutoSize        =   -1  'True
         Caption         =   "Edit Before Printing"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   3605
         Width           =   1350
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Print Data Manager Method Name"
         Height          =   375
         Index           =   14
         Left            =   4680
         TabIndex        =   40
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label lblPrinterDestinationDetail 
         Caption         =   "Remote Location"
         Height          =   195
         Left            =   4680
         TabIndex        =   36
         Top             =   3200
         Width           =   1335
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Printer Destination"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   3200
         Width           =   1335
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Recipient Type"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   2730
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Print Menu Access"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   33
         Top             =   2325
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Customer Indicator"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   2330
         Width           =   1455
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Default Copies"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   31
         Top             =   1905
         Width           =   1155
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Maximum Copies"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   1905
         Width           =   1185
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Set Inactive"
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   29
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Template Id"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   840
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Document Group"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label lblPrintingDetail 
         Caption         =   "Template Name"
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   26
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "Minimum Role Level"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1500
         Width           =   1425
      End
      Begin VB.Label lblPrintingDetail 
         AutoSize        =   -1  'True
         Caption         =   "DPS Template Id"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   24
         Top             =   270
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditPrintingTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPrintingTemplates
' Description   : Form which allows the adding/editing of Print Templates
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added Form
' DB        04/11/02    BMIDS00840  Added edit and view before printing option buttons
' MV        07/03/03    BM0341  Modified : cmdOk_Onclick();cboPrinterDestination_Click();ValidatePrinterDestination()
' MV        01/04/03    BM0341  Modified : cboPrinterDestination_Click();ValidatePrinterDestination()
' MV        16/04/03    BM0341  Modified : cboPrinterDestination_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Specific history:

' GHun      13/07/2005  MAR7    Integrate local printing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'EPSOM HISTORY
'
' Prog      Date        Description
' TW        25/01/2007  EP2_990 - Apply EP1257 to Epsom 2 Supervisor (CC56 -  Gemini printing  Supervisor changes)
' GHun      13/02/2007  EP2_1363 - Added EmailRecipientType combo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private m_clsPrintTemplate As PrintingTemplateTable
Private m_bIsEdit As Boolean
Private m_colKeys As Collection
Private m_ReturnCode As MSGReturnCode
Private m_sTemplateID As String
Private m_clsComboValidation As ComboValidationTable
Private m_bScreenUpdated As Boolean
Private m_clsAvailableTemplates As AvailableTemplatesTable
Private m_bTaskManagementExists As Boolean

'Form Constants
Private Const TEMPLATEID As Long = 0
Private Const MAXIMUM_COPIES As Long = 2
Private Const TEMPLATE_NAME As Long = 4
Private Const DEFAULT_COPIES As Long = 6

'ComboGroup Constants
Private Const m_sComboMinimumRole As String = "UserRole"
Private Const m_sComboDocumentType As String = "PrintDocumentType"
Private Const m_sComboRecipientType As String = "PrintRecipientType"
Private Const m_sComboPrinterDestination As String = "PrinterDestination"
Private Const m_sComboApplicationStage As String = "ApplicationStage"
Private Const m_sComboEmailRecipientType As String = "EmailRecipientType"   'EP2_1363 GHun

'MAR7 GHun
Private Const m_sComboDocumentDeliveryType As String = "DocumentDeliveryType"
Private Const m_sComboDefaultPrinterTray As String = "DefaultPrinterTray"
'MAR7 End

' TW 25/01/2007 EP2_990
Private Const m_sGeminiPrintMode As String = "GeminiPrintMode"
' TW 25/01/2007 EP2_990 End

Private m_sFormLoaded  As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormEditState
' Description   :   Initialises the form in Edit Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormEditState()
    On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim bRet As Boolean
    
    Set clsTableAccess = m_clsPrintTemplate
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData(POPULATE_KEYS)
    bRet = True
    
    If Not rs Is Nothing Then
        If rs.RecordCount = 0 Then
            bRet = False
        End If
    Else
        bRet = False
    End If
    
    If Not bRet Then
        g_clsErrorHandling.RaiseError errRecordSetEmpty
    Else
        PopulateScreenFields
    End If
    
    txtDetail(TEMPLATEID).Enabled = False
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidatePrinterDestination() As Boolean
    
    Dim vVal As Variant
    Dim clsComboTable As ComboValidationTable
    Dim bRet As Boolean
    Dim col As New Collection
    Dim i  As Integer
    Dim sValidationType As String
    
    On Error GoTo Failed
    
    ValidatePrinterDestination = True
    
    g_clsFormProcessing.HandleComboExtra cboPrinterDestination, vVal, GET_CONTROL_VALUE
    
    If vVal <> "" Then
        Set clsComboTable = New ComboValidationTable
        Set col = clsComboTable.GetComboValidationAsCollection("PrinterDestination", CStr(vVal))
        
        If col.Count > 0 Then
            For i = 1 To col.Count
                sValidationType = col.Item(i)
                
                If sValidationType = "R" Or sValidationType = "F" Or sValidationType = "X" Then
                    If Me.txtRemoteLocation.Text = "" Then
                        'EP2_1363 GHun
                        frmEditPrintingTemplate.txtRemoteLocation.SetFocus
                        MsgBox "Remote location must be entered.", vbExclamation
                        'EP2_1363 End
                        ValidatePrinterDestination = False
                        Exit For
                    End If
                'EP2_1363 GHun
                ElseIf sValidationType = "EF" Then
                    If cboEmailRecipientType.List(cboEmailRecipientType.ListIndex) = COMBO_NONE Then
                        MsgBox "Please select an Email Recipient.", vbExclamation
                        cboEmailRecipientType.SetFocus
                        ValidatePrinterDestination = False
                    End If
                End If
                'EP2_1363 End
            Next
        End If
    End If
    
    Exit Function
    
Failed:
    
End Function

'MAR7 GHun
Private Sub cboDocDelType_Click()
    
    m_bScreenUpdated = True

    Dim vVal As Variant
    Dim clsComboTable As ComboValidationTable
    Dim col  As New Collection
    Dim i As Integer
    Dim strValidationType  As String
    
    On Error GoTo Failed
    
    g_clsFormProcessing.HandleComboExtra cboDocDelType, vVal, GET_CONTROL_VALUE
    
    If m_sFormLoaded = True Then
        
        If vVal <> "" Then
            
            Set clsComboTable = New ComboValidationTable
            Set col = clsComboTable.GetComboValidationAsCollection("DocumentDeliveryType", CStr(vVal))
            
            If col.Count > 0 Then
            
                For i = 1 To col.Count
                    
                    strValidationType = col.Item(i)
                    
                    If (strValidationType = "pdf") Then
                        optEditPrint(0).Enabled = False
                        optEditPrint(1).Enabled = False
                        Exit For
                    Else
                        optEditPrint(0).Enabled = True
                        optEditPrint(1).Enabled = True
                    End If
                Next
                
            End If
        End If
    End If
    Exit Sub
Failed:

     g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub
'MAR7 End

Private Sub cboDocumentType_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cboMinRoleLevel_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cboPrinterDestination_Click()
    
    m_bScreenUpdated = True

    Dim sValidation As String
    Dim vVal As Variant
    Dim clsComboTable As ComboValidationTable
    Dim col  As New Collection
    Dim i As Integer
    Dim strValidationType  As String
    Dim sTempVal As String
    Dim TempCol As New Collection
    Dim j As Integer
    
    On Error GoTo Failed
    
    g_clsFormProcessing.HandleComboExtra cboPrinterDestination, vVal, GET_CONTROL_VALUE
    
    'If m_sFormLoaded = True Then   'MAR7 GHun
        
        If vVal <> "" Then
            
            Set clsComboTable = New ComboValidationTable
            Set col = clsComboTable.GetComboValidationAsCollection("PrinterDestination", CStr(vVal))
            
            If col.Count > 0 Then
            
                For i = 1 To col.Count
                    
                    strValidationType = col.Item(i)
                    
                    If (strValidationType = "R" Or strValidationType = "F" Or strValidationType = "X") Then
                           
                        txtRemoteLocation.Text = ""
                        txtRemoteLocation.Enabled = True
                        txtRemoteLocation.Mandatory = True
                        'EP2_1363 GHun
                        txtRemoteLocation.Visible = True
                        lblPrinterDestinationDetail.Visible = True
                        'EP2_1363 End
                        'MAR7 GHun
                        txtDetail(2).Enabled = True
                        txtDetail(6).Enabled = True
                        cboFirstPagePrinterTray.Enabled = True
                        cboOtherPagesPrinterTray.Enabled = True
                        'MAR7 End
                        
                        'EP2_1363 GHun
                        cboEmailRecipientType.Visible = False
                        cboEmailRecipientType.ListIndex = cboEmailRecipientType.ListCount - 1
                        If strValidationType = "F" Then
                            lblPrinterDestinationDetail.Caption = "File Location"
                        Else
                            lblPrinterDestinationDetail.Caption = "Remote Location"
                        End If
                        'EP2_1363 End
                        
                        If m_clsPrintTemplate.GetPrinterDestinationType = vVal Then
                        
                            Set TempCol = clsComboTable.GetComboValidationAsCollection("PrinterDestination", CStr(m_clsPrintTemplate.GetPrinterDestinationType))
                            
                            If TempCol.Count > 0 Then
                            
                                For j = 1 To TempCol.Count
                                    
                                    If strValidationType = TempCol(j) Then
                                        txtRemoteLocation.Text = m_clsPrintTemplate.GetRemotePrinterLocation
                                        Exit For
                                    End If
                                
                                Next
                                
                            End If
                            
                        End If
                        
                        Exit For
                    'MAR7 GHun
                    ElseIf strValidationType = "W" Then
                        
                        txtDetail(2).Text = ""
                        txtDetail(2).Enabled = False
                        
                        txtDetail(6).Enabled = False
                        
                        cboFirstPagePrinterTray.Enabled = False
                        cboFirstPagePrinterTray.ListIndex = 1
                        
                        cboOtherPagesPrinterTray.Enabled = False
                        cboOtherPagesPrinterTray.ListIndex = 1
                            
                        txtRemoteLocation.Text = ""
                        txtRemoteLocation.Enabled = False
                    'MAR7 End
                    
                    'EP2_1363 GHun
                        txtRemoteLocation.Visible = False
                        lblPrinterDestinationDetail.Visible = False
                        txtRemoteLocation.Visible = False
                        cboEmailRecipientType.Visible = False
                        cboEmailRecipientType.ListIndex = cboEmailRecipientType.ListCount - 1
                    
                    ElseIf strValidationType = "EF" Then
                        txtRemoteLocation.Visible = False
                        lblPrinterDestinationDetail.Visible = True
                        lblPrinterDestinationDetail.Caption = "Email Recipient"
                        
                        cboEmailRecipientType.Visible = True
                        cboEmailRecipientType.Enabled = True
                        If cboEmailRecipientType.Visible Then
                            cboEmailRecipientType.SetFocus
                        End If
                    'EP2_1363 End
                    
                    Else
                        
                        txtRemoteLocation.Text = ""
                        txtRemoteLocation.Enabled = False
                        txtRemoteLocation.Mandatory = False
                        'MAR7 GHun
                        txtDetail(2).Enabled = True
                        txtDetail(6).Enabled = True
                        cboFirstPagePrinterTray.Enabled = True
                        cboOtherPagesPrinterTray.Enabled = True
                        'MAR7 End
                    
                        'EP2_1363 GHun
                        txtRemoteLocation.Visible = False
                        lblPrinterDestinationDetail.Visible = False
                        cboEmailRecipientType.Visible = False
                        cboEmailRecipientType.ListIndex = cboEmailRecipientType.ListCount - 1
                        'EP2_1363 End
                    End If
                Next
            End If
            
        End If
    'End If 'MAR7 GHun
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Private Sub cboRecipientType_Click()
    m_bScreenUpdated = True
End Sub

Private Sub chkInactive_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    Dim bRet As Boolean
    On Error GoTo Failed
    
    
    bRet = ValidatePrinterDestination()
    If bRet Then
        bRet = DoOKProcessing
    End If
    If bRet Then
        SetReturnCode
        Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Activate()
    m_bScreenUpdated = False
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    
    Set m_clsPrintTemplate = New PrintingTemplateTable
    Set m_clsComboValidation = New ComboValidationTable
    Set m_clsAvailableTemplates = New AvailableTemplatesTable
    
    m_sFormLoaded = False
    m_sTemplateID = ""
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetFormEditState
    Else
        SetFormAddState
    End If
       
    m_sFormLoaded = True
    m_bScreenUpdated = False
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormAddState
' Description   :   Initialises the form in Edit Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormAddState()
    
    On Error GoTo Failed
        Dim rs As ADODB.Recordset
        
        'Populate Selected Items with an empty Recordset
        Set rs = TableAccess(m_clsAvailableTemplates).GetTableData(POPULATE_EMPTY)
        TableAccess(m_clsAvailableTemplates).SetRecordSet rs
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenControls()
    On Error GoTo Failed
    Dim sField As String
    Dim clsComboValue As ComboValueTable
    Dim rs As ADODB.Recordset
    
    Set clsComboValue = New ComboValueTable
    'Document Group
    g_clsFormProcessing.PopulateCombo m_sComboDocumentType, cboDocumentType, False
    
    'Minimun Role
    g_clsFormProcessing.PopulateCombo m_sComboMinimumRole, cboMinRoleLevel, False
    
    'Recipient Type
    g_clsFormProcessing.PopulateCombo m_sComboRecipientType, cboRecipientType
    
    g_clsFormProcessing.PopulateCombo m_sComboPrinterDestination, cboPrinterDestination
    
    'MAR7 GHun  Add Document Delivery Type combo
    'Document Delivery Type
    g_clsFormProcessing.PopulateCombo m_sComboDocumentDeliveryType, cboDocDelType
    
    ' First Page Printer Tray
    g_clsFormProcessing.PopulateCombo m_sComboDefaultPrinterTray, cboFirstPagePrinterTray
    
    ' Other Page Printer Tray
    g_clsFormProcessing.PopulateCombo m_sComboDefaultPrinterTray, cboOtherPagesPrinterTray
    'MAR7 End
    
' TW 25/01/2007 EP2_990
    g_clsFormProcessing.PopulateCombo m_sGeminiPrintMode, cboGeminiPrintMode
' TW 25/01/2007 EP2_990 End
    
    'EP2_1363 GHun
    'EmailRecipient Type
    g_clsFormProcessing.PopulateCombo m_sComboEmailRecipientType, cboEmailRecipientType
    'EP2_1363 End
    
    'Printer Destination
    'Returns a Recordset with the corresponding validation types
    'Set rs = m_clsComboValidation.GetComboGroupWithValidationType(m_sComboPrinterDestination)
    ' DJP SQL Server port, not needed
    'TableAccess(m_clsComboValidation).SetRecordSet rs
    
    'Set rs = clsComboValue.GetComboGroupValues(m_sComboPrinterDestination)
    'Set cboPrinterDestination.RowSource = rs
    'sField = m_clsPrintTemplate.GetListField
    'cboPrinterDestination.ListField = sField
    
    'Populate Stages SwapList
    SetStageHeaders
    'Does the task management schema exist?
    
    m_bTaskManagementExists = g_clsMainSupport.DoesTaskManagementExist
    If m_bTaskManagementExists Then
        PopulateAvailableItemsFromRS
    Else
        PopulateAvailableItemsFromComboGroup
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Private Sub SetStageHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim colLine As Collection

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Stage Name"
    headers.Add lvHeaders

    swapStage.SetFirstColumnHeaders headers
    swapStage.SetSecondColumnHeaders headers

    swapStage.SetFirstTitle "Select Stage(s)"
    swapStage.SetSecondTitle "Selected Stage(s)"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim sStageName As String
    Dim sStageID As String
    Dim colLine As Collection
    Dim bExists As Boolean
    
    Set colLine = New Collection
    Set clsSwapExtra = New SwapExtra
    
    If TableAccess(m_clsAvailableTemplates).RecordCount > 0 Then
        Set rs = TableAccess(m_clsAvailableTemplates).GetRecordSet
        
        Do While Not rs.EOF
            
            Set colLine = New Collection
            sStageID = m_clsAvailableTemplates.GetStageID
            sStageName = m_clsAvailableTemplates.GetStageName
            
            bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
            
            If Not bExists Then
                Set clsSwapExtra = New SwapExtra
                clsSwapExtra.SetValueID sStageID
                
                colLine.Add sStageName
                swapStage.AddLineSecond colLine, clsSwapExtra
            End If
            
            rs.MoveNext
        Loop
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim vVal As Variant
    Dim clsComboUtils As New ComboUtils
    
    'TEXTBOXES - Template ID
    txtDetail(TEMPLATEID).Text = m_clsPrintTemplate.GetHostTemplateID
    m_sTemplateID = txtDetail(TEMPLATEID).Text
    
    'DPS TemplateID
    txtDPSTemplateID.Text = m_clsPrintTemplate.GetDPSTemplateID
    
    'Template Name
    txtDetail(TEMPLATE_NAME).Text = m_clsPrintTemplate.GetHostTemplateName
    
    'Template Description
    txtDescription.Text = m_clsPrintTemplate.GetHostTemplateDescription
    
    ' Maximum Copies
    txtDetail(MAXIMUM_COPIES).Text = m_clsPrintTemplate.GetMaxCopies
    
    'Default Copies
    txtDetail(DEFAULT_COPIES).Text = m_clsPrintTemplate.GetDefaultCopies
    
    'Remote Location
    txtRemoteLocation.Text = m_clsPrintTemplate.GetRemotePrinterLocation
    
    'COMBO'S - Document Group
    vVal = m_clsPrintTemplate.GetTemplateGroupID
    g_clsFormProcessing.HandleComboExtra cboDocumentType, vVal, SET_CONTROL_VALUE
    
    'Minimum Role
    vVal = m_clsPrintTemplate.GetMinRoleLevel
    g_clsFormProcessing.HandleComboExtra cboMinRoleLevel, vVal, SET_CONTROL_VALUE
    
    'Recipient Type
    vVal = m_clsPrintTemplate.GetRecipientTypeID
    g_clsFormProcessing.HandleComboExtra cboRecipientType, vVal, SET_CONTROL_VALUE
    
    'Printer Destination
    'vVal = clsComboUtils.GetValueNameFromValueID(m_clsPrintTemplate.GetPrinterDestinationType, m_sComboPrinterDestination)
    vVal = m_clsPrintTemplate.GetPrinterDestinationType
    g_clsFormProcessing.HandleComboExtra cboPrinterDestination, vVal, SET_CONTROL_VALUE
    cboPrinterDestination_Click 'EP2_1363 GHun
    
    'MAR7 GHun Document Delivery Type
    vVal = m_clsPrintTemplate.GetDocumentDeliveryType
    g_clsFormProcessing.HandleComboExtra cboDocDelType, vVal, SET_CONTROL_VALUE
    
    ' First Page Printer Tray
    vVal = m_clsPrintTemplate.GetFirstPagePrinterTray
    g_clsFormProcessing.HandleComboExtra cboFirstPagePrinterTray, vVal, SET_CONTROL_VALUE
    
    ' Other Page Printer Tray
    vVal = m_clsPrintTemplate.GetOtherPagesPrinterTray
    g_clsFormProcessing.HandleComboExtra cboOtherPagesPrinterTray, vVal, SET_CONTROL_VALUE
    'MAR7 End
    
' TW 25/01/2007 EP2_990
    'Gemini Print Mode
    vVal = m_clsPrintTemplate.GetGeminiPrintMode
    g_clsFormProcessing.HandleComboExtra cboGeminiPrintMode, vVal, SET_CONTROL_VALUE
    
    'Pre-create for pack?
    g_clsFormProcessing.HandleRadioButtons optPrecreate(OPT_YES), optPrecreate(OPT_NO), m_clsPrintTemplate.GetPrecreateForPackInd, SET_CONTROL_VALUE
' TW 25/01/2007 EP2_990 End
    
    'Print Manager Method Name
    txtPrintMethodName.Text = m_clsPrintTemplate.GetPrintManagerMethod
    
    'Customer Indicator
    g_clsFormProcessing.HandleRadioButtons optCustInd(OPT_YES), optCustInd(OPT_NO), m_clsPrintTemplate.GetCustSpecificInd, SET_CONTROL_VALUE
    
    'Print Menu Access?
    g_clsFormProcessing.HandleRadioButtons optPrintMenu(OPT_YES), optPrintMenu(OPT_NO), m_clsPrintTemplate.GetPrintMenuAccessInd, SET_CONTROL_VALUE
    
    'DB BMIDS00840 - Added edit and view before print
    'Edit before print
    g_clsFormProcessing.HandleRadioButtons optEditPrint(OPT_YES), optEditPrint(OPT_NO), m_clsPrintTemplate.GetEditBeforePrintInd, SET_CONTROL_VALUE
    
    'View before print
    g_clsFormProcessing.HandleRadioButtons optViewPrint(OPT_YES), optViewPrint(OPT_NO), m_clsPrintTemplate.GetViewBeforePrintInd, SET_CONTROL_VALUE
    'DB END
    
    '*=[MC]BMIDS770 - CC072 - Prevent Edit for documents stored in DMS
    g_clsFormProcessing.HandleRadioButtons optPreventEdit(OPT_YES), optPreventEdit(OPT_NO), m_clsPrintTemplate.GetPreventEditInDMS, SET_CONTROL_VALUE
    
    '*=Enable if Edit Before print is selected "Yes"
    If optEditPrint(0).Value = True Then
        optPreventEdit(0).Enabled = True
        optPreventEdit(1).Enabled = True
    End If
    '*=[MC]SECTION END - CC072
    
    'MAR7 GHun Start
    'Print To Web
    g_clsFormProcessing.HandleRadioButtons optPrintToWeb(OPT_YES), optPrintToWeb(OPT_NO), m_clsPrintTemplate.GetPrintToWeb, SET_CONTROL_VALUE
    'MAR7 End
    
    'Set Inative
    vVal = m_clsPrintTemplate.GetInactiveIndicator
    g_clsFormProcessing.HandleCheckBox chkInactive, vVal, SET_CONTROL_VALUE
    
    'EP2_1363 GHun
    'Email Recipient Type
    vVal = m_clsPrintTemplate.GetEmailRecipientType
    g_clsFormProcessing.HandleComboExtra cboEmailRecipientType, vVal, SET_CONTROL_VALUE
    
    If IsNumeric(vVal) Then
        cboEmailRecipientType.Enabled = True
    Else
        cboEmailRecipientType.Enabled = False
    End If
    'EP2_1363 End
    
    'Does the task managment schema exist?
    If m_bTaskManagementExists Then
        'Populate from Task Table
        m_clsAvailableTemplates.GetSelectedStages m_sTemplateID, True
        PopulateSelectedItems
        m_clsAvailableTemplates.GetSelectedStages m_sTemplateID, True
    Else
        'Populate From Combo Group
        m_clsAvailableTemplates.GetSelectedStagesFromComboGroup m_sTemplateID
        PopulateSelectedItems
        m_clsAvailableTemplates.GetSelectedStagesFromComboGroup m_sTemplateID, False
    End If
    chkInactive.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveScreenFields()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim vVal As Variant
    
    Set clsTableAccess = m_clsPrintTemplate
    
    'Next Set the TemplateID
    m_sTemplateID = txtDetail(TEMPLATEID).Text
    
    If Not m_bIsEdit Then
        g_clsFormProcessing.CreateNewRecord m_clsPrintTemplate
    End If
    
    m_clsPrintTemplate.SetTemplateID m_sTemplateID
    
    m_clsPrintTemplate.SetDPSTemplateID txtDPSTemplateID.Text
    
    g_clsFormProcessing.HandleComboExtra cboDocumentType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetTemplateGroupID vVal
    
    m_clsPrintTemplate.SetHostTemplateName txtDetail(TEMPLATE_NAME).Text
    
    m_clsPrintTemplate.SetHostTemplateDescription txtDescription.Text
    
    g_clsFormProcessing.HandleComboExtra cboMinRoleLevel, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetMinRoleLevel vVal
    
    g_clsFormProcessing.HandleCheckBox chkInactive, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetInactiveIndicator vVal
    
    m_clsPrintTemplate.SetMaxCopies txtDetail(MAXIMUM_COPIES).Text
    m_clsPrintTemplate.SetDefaultCopies txtDetail(DEFAULT_COPIES).Text
    
    g_clsFormProcessing.HandleRadioButtons optCustInd(OPT_YES), optCustInd(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetCustomerSpecificInd vVal
    
    'Print Menu
    g_clsFormProcessing.HandleRadioButtons optPrintMenu(OPT_YES), optPrintMenu(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPrintMenuAccessInd vVal
    
    'DB BMIDS00840 - Added edit & view before printing
    g_clsFormProcessing.HandleRadioButtons optEditPrint(OPT_YES), optEditPrint(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetEditBeforePrintInd vVal
    
    g_clsFormProcessing.HandleRadioButtons optViewPrint(OPT_YES), optViewPrint(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetViewBeforePrintInd vVal
    'DB END
    
    '*=[MC]BMIDS770 - CC072
    g_clsFormProcessing.HandleRadioButtons optPreventEdit(OPT_YES), optPreventEdit(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPreventEditInDMS vVal
    
    '*=[MC]BMIDS770 SECTION END
    
    'MAR7 GHun Start
    'Print To Web
    g_clsFormProcessing.HandleRadioButtons optPrintToWeb(OPT_YES), optPrintToWeb(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPrintToWeb vVal
    'MAR7 End
    
    'Recipient Type
    g_clsFormProcessing.HandleComboExtra cboRecipientType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetRecipientType vVal
    
    'Printer Destination
    'vVal = m_clsComboValidation.GetValueID
    g_clsFormProcessing.HandleComboExtra cboPrinterDestination, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPrintDestinationType vVal
    
    'Remote Location
    m_clsPrintTemplate.SetRemotePrinterLocation txtRemoteLocation.Text
    
    'Print Manager Method Name
    m_clsPrintTemplate.SetPrintManagerMethodName txtPrintMethodName.Text
    
    'MAR7 GHun Start  Document Delivery Type
    g_clsFormProcessing.HandleComboExtra cboDocDelType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetDocumentDeliveryType vVal
    
    ' First Page Printer Tray
    g_clsFormProcessing.HandleComboExtra cboFirstPagePrinterTray, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetFirstPagePrinterTray vVal
    
    ' Other Page Printer Tray
    g_clsFormProcessing.HandleComboExtra cboOtherPagesPrinterTray, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetOtherPagesPrinterTray vVal
    'MAR7 End
    
' TW 25/01/2007 EP2_990
    g_clsFormProcessing.HandleComboExtra cboGeminiPrintMode, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetGeminiPrintMode vVal
    
    g_clsFormProcessing.HandleRadioButtons optPrecreate(OPT_YES), optPrecreate(OPT_NO), vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetPrecreateForPackInd vVal
' TW 25/01/2007 EP2_990 End
    
    'EP2_1363 GHun
    'Recipient Type
    g_clsFormProcessing.HandleComboExtra cboEmailRecipientType, vVal, GET_CONTROL_VALUE
    m_clsPrintTemplate.SetEmailRecipientType vVal
    'EP2_1363 End
    
    clsTableAccess.Update
    
    'Stages
    SaveSelectedStages
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean

    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim vSaveChanges As Variant
    
    bShowError = True

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet Then
        bRet = ValidateScreenData
    End If
    
    If m_bIsEdit And bRet Then
        'This record is an edit and Valid
        vSaveChanges = MsgBox("You are about to update the selected Printing Template. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
        Select Case vSaveChanges
            Case vbYes
                If bRet Then
                    SaveScreenFields
                    SaveChangeRequest
                End If
            Case vbNo
                bRet = False
        End Select
    ElseIf Not m_bIsEdit And bRet Then
        SaveScreenFields
        SaveChangeRequest
        bRet = True
    End If
        
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub optCustInd_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub optPrintMenu_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub txtDetail_Change(Index As Integer)
    m_bScreenUpdated = True
End Sub
Private Sub optEditPrint_Click(Index As Integer)
    
    If optEditPrint(0).Value = True Then
        optViewPrint(0).Enabled = False
        optViewPrint(1).Enabled = False
        '*=[MC]BMIDS770 Enable if optViewPrint selected
        optPreventEdit(0).Enabled = True
        optPreventEdit(1).Enabled = True
    Else
        optViewPrint(0).Enabled = True
        optViewPrint(1).Enabled = True
        '*=[MC]BMIDS770 Disable If optviewprint not selected to yes
        optPreventEdit(0).Enabled = False
        optPreventEdit(1).Enabled = False
        'If optViewPrint not selected, Unselect prevent edit
        optPreventEdit(0).Value = True
    End If
    
m_bScreenUpdated = True
        
End Sub
Private Sub optViewPrint_Click(Index As Integer)

    If optViewPrint(0).Value = True Then
        optEditPrint(0).Enabled = False
        optEditPrint(1).Enabled = False
    Else
        optEditPrint(0).Enabled = True
        optEditPrint(1).Enabled = True
    End If
m_bScreenUpdated = True

End Sub

Private Sub SaveSelectedStages()
    On Error GoTo Failed
    
    Dim nThisItem       As Integer
    Dim colValues       As Collection
    Dim sStageID        As String
    Dim clsSwapExtra    As SwapExtra
    
    If TableAccess(m_clsAvailableTemplates).RecordCount > 0 Then
        'Delete all Existing Records, need to repopulate available template (Without the stagename), so not to delete the stage
        m_clsAvailableTemplates.GetSelectedStages m_sTemplateID, False
        TableAccess(m_clsAvailableTemplates).DeleteAllRows
        TableAccess(m_clsAvailableTemplates).Update
    End If
    
    If swapStage.GetSecondCount > 0 Then
        For nThisItem = 1 To swapStage.GetSecondCount
            'Add a row and set the fields
            TableAccess(m_clsAvailableTemplates).AddRow
        
            Set colValues = swapStage.GetLineSecond(nThisItem, clsSwapExtra)
    
            sStageID = clsSwapExtra.GetValueID
            
            m_clsAvailableTemplates.SetTemplateID m_sTemplateID
            m_clsAvailableTemplates.SetStageID sStageID
        Next
    End If
    
    If TableAccess(m_clsAvailableTemplates).RecordCount > 0 Then
        TableAccess(m_clsAvailableTemplates).Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateAvailableItemsFromComboGroup()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsComboValidation As ComboValidationTable
    Dim nThisItem As Long
    Dim clsSwapExtra As SwapExtra
    Dim sStageID As String
    Dim sStageName As String
    Dim bExists As Boolean
    Dim colLine As Collection
    
    Set clsComboValidation = New ComboValidationTable
    
    Set rs = clsComboValidation.GetComboGroupWithValidationType(m_sComboApplicationStage)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For nThisItem = 1 To rs.RecordCount
                
                Set colLine = New Collection
                
                sStageID = clsComboValidation.GetValueID()
                sStageName = clsComboValidation.GetValueName()
                
                bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
                
                If Not bExists Then
                    Set clsSwapExtra = New SwapExtra
                    clsSwapExtra.SetValueID sStageID
                    
                    colLine.Add sStageName
                    swapStage.AddLineFirst colLine, clsSwapExtra
                
                End If
                rs.MoveNext
            Next
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesDPSTemplateIDExist
' Description   :   Checks to see if a DPS Document Record exists with the ID entered
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesDPSTemplateIDExist() As Boolean
    On Error GoTo Failed
    Dim clsTemplate As TemplateTable
    Dim bRet As Boolean
    Dim colValues As Collection
    
    Set colValues = New Collection
    
    Set clsTemplate = New TemplateTable
    
    colValues.Add txtDPSTemplateID.Text
    
    bRet = TableAccess(clsTemplate).DoesRecordExist(colValues, TableAccess(clsTemplate).GetKeyMatchFields)
    
    
    DoesDPSTemplateIDExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateAvailableItemsFromRS
' Description   :   If the Task Management Schema exists then the available Tasks must be populated from the Task Table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableItemsFromRS()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsComboValidation As ComboValidationTable
    Dim nThisItem As Long
    Dim clsSwapExtra As SwapExtra
    Dim sStageID As String
    Dim sStageName As String
    Dim bExists As Boolean
    Dim colLine As Collection
    
    Set clsComboValidation = New ComboValidationTable
    If m_bIsEdit Then
        m_sTemplateID = m_colKeys(1)
    End If
    
    Set rs = m_clsAvailableTemplates.GetUnassignedStages(m_sTemplateID)
    TableAccess(m_clsAvailableTemplates).SetRecordSet rs
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For nThisItem = 1 To rs.RecordCount
                
                Set colLine = New Collection
                
                sStageID = m_clsAvailableTemplates.GetStageID()
                sStageName = m_clsAvailableTemplates.GetStageName()
                
                bExists = g_clsFormProcessing.DoesSwapValueExist(swapStage, sStageName)
                
                If Not bExists Then
                    Set clsSwapExtra = New SwapExtra
                    clsSwapExtra.SetValueID sStageID
                    
                    colLine.Add sStageName
                    swapStage.AddLineFirst colLine, clsSwapExtra
                
                End If
                rs.MoveNext
            Next
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub txtDPSTemplateID_Validate(Cancel As Boolean)
    On Error GoTo Failed
    
    Dim bRet As Boolean
        
    If Len(txtDPSTemplateID.Text) > 0 Then
        bRet = DoesDPSTemplateIDExist
        
        If Not bRet Then
            Cancel = True
            txtDPSTemplateID.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Your DPS template Id is not valid, please re-enter."
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Common Function used to setup a promotion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sDesc = m_sTemplateID
    
    colMatchValues.Add m_sTemplateID
    Set clsTableAccess = m_clsPrintTemplate
    
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess, sDesc
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates the data on the screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = True
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number
    
End Function
Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sTemplateID As String
    Dim colValues As New Collection

    sTemplateID = txtDetail(TEMPLATEID).Text

    bRet = False
    If Len(sTemplateID) > 0 Then
        colValues.Add sTemplateID

        bRet = TableAccess(m_clsPrintTemplate).DoesRecordExist(colValues)

        If bRet = True Then
           g_clsErrorHandling.DisplayError "TemplateID must be unique"
           txtDetail(TEMPLATEID).SetFocus
        End If
    End If
    DoesRecordExist = bRet
End Function

'MAR7 Start
Private Sub optPrintToWeb_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub
'MAR7 End

'EP2_1363 GHun
Private Sub cboEmailRecipientType_Click()
    m_bScreenUpdated = True
End Sub
'EP2_1363 End

