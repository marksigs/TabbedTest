VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBroker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit AR Firm"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmEditBroker.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8220
      TabIndex        =   30
      Top             =   5640
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   5640
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Firm Details"
      TabPicture(0)   =   "frmEditBroker.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Contact Details"
      TabPicture(1)   =   "frmEditBroker.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Trading Names"
      TabPicture(2)   =   "frmEditBroker.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Networks"
      TabPicture(3)   =   "frmEditBroker.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Linked Networks"
         Height          =   4695
         Left            =   -74760
         TabIndex        =   59
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdAddNetwork 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   4080
            Width           =   1335
         End
         Begin VB.CommandButton cmdRemoveNetwork 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   4080
            Width           =   1335
         End
         Begin VB.CheckBox chkSelectedNetwork 
            Alignment       =   1  'Right Justify
            Caption         =   "Selected Network"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6870
            TabIndex        =   28
            Top             =   4080
            Width           =   1695
         End
         Begin MSGOCX.MSGListView lvCurrentAssociations 
            Height          =   3615
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   6376
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Trading Names"
         Height          =   4695
         Left            =   240
         TabIndex        =   58
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdDeleteTradingName 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   24
            Top             =   4080
            Width           =   1275
         End
         Begin VB.CommandButton cmdEditTradingName 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   23
            Top             =   4080
            Width           =   1275
         End
         Begin VB.CommandButton cmdAddTradingName 
            Caption         =   "&Add"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   4080
            Width           =   1275
         End
         Begin MSGOCX.MSGListView lvTradingNames 
            Height          =   3615
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   6376
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Procuration Fee Adjustments"
         Height          =   855
         Left            =   -74760
         TabIndex        =   54
         Top             =   3000
         Width           =   8895
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3240
            ScaleHeight     =   255
            ScaleWidth      =   2655
            TabIndex        =   55
            Top             =   360
            Width           =   2655
            Begin VB.OptionButton optIncludeOnlineSubmissionLoading 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   7
               Top             =   0
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optIncludeOnlineSubmissionLoading 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Include Online Submission Loading ?"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   2610
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Fax Contact Details"
         Height          =   1575
         Left            =   -70200
         TabIndex        =   50
         Top             =   3600
         Width           =   4335
         Begin MSGOCX.MSGEditBox txtCountryCode2 
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAreaCode2 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtTelephoneNumber2 
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Number"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   53
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Telephone Contact Details"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   46
         Top             =   3600
         Width           =   4335
         Begin MSGOCX.MSGEditBox txtCountryCode1 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAreaCode1 
            Height          =   285
            Left            =   1440
            TabIndex        =   16
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtTelephoneNumber1 
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Number"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   49
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   48
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contact Details"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   8895
         Begin MSGOCX.MSGEditBox txtPostCode 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   2520
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine2 
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Top             =   720
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine3 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Top             =   1080
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine4 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   1440
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine5 
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Top             =   1800
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine6 
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   2160
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin MSGOCX.MSGEditBox txtAddressLine1 
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   503
            TextType        =   4
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
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 1"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Postcode"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 2"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 4"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   36
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 5"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 6"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   34
            Top             =   2160
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Address Line 3"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Firm Details"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   40
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdCreateAsOmigaUnit 
            Caption         =   "Create as Omiga Unit"
            Height          =   375
            Left            =   6600
            TabIndex        =   5
            Top             =   1800
            Width           =   1935
         End
         Begin MSGOCX.MSGEditBox txtUnitID 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            Enabled         =   0   'False
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
         Begin MSGOCX.MSGEditBox txtARFirmName 
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   1080
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   503
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
         End
         Begin MSGOCX.MSGComboBox cboFirmStatus 
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   1440
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSGOCX.MSGComboBox cboListingStatus 
            Height          =   315
            Left            =   1440
            TabIndex        =   4
            Top             =   1800
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSGOCX.MSGEditBox txtFSAARFirmRef 
            Height          =   285
            Left            =   1440
            TabIndex        =   0
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            Enabled         =   0   'False
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FSA Reference"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Listing Status"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   44
            Top             =   1800
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Status"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   43
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Name"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unit ID"
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Width           =   495
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Packager Indicator"
      Height          =   195
      Index           =   30
      Left            =   4800
      TabIndex        =   45
      Top             =   1920
      Width           =   1350
   End
End
Attribute VB_Name = "frmEditBroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditBroker
' Description   : To Add and Edit AR Firms
'
' Change history
' Prog      Date        Description
' TW        17/10/2006  EP2_15 - Created
' TW        18/11/2006  EP2_132 ECR20/21 Added 'Updated by' and 'Date last updated'
' TW        05/12/2006  EP2_276 - Deal with mandatory fields.
' TW        06/12/2006  EP2_330 - Error linking a Packager to an a Association
' TW        08/12/2006  EP2_360 - AR firm networks is currently showing mortgage club. Should be showing Network.
' TW        09/01/2007  EP2_728 - Caching unit numbers and Ids
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        05/02/2007  EP2_706 - Should  network be mandatory data for ar firms
' TW        13/02/2007  EP2_1334 - AR firms not holding value in firm status
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' TW        12/03/2007  EP2_1190 - Change Control 281 - AR Firm to nominate a single Network for Mortgage Business
' TW        16/05/2007  VR86 - Removed 'Firm Activities' tab and all associated 'Permissions' code
' AW        02/10/2007  OMIGA00003231   Disable edit/delete when no row present
' AW        08/10/2007  OMIGA00003355   Make dialog style fixed single
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Private data
Private m_clsARFirmTable As ARFirmTable

Private m_clsAppointmentsTable As AppointmentsTable

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_colKeys As Collection

Dim strActivityID As String
Dim strARFirmID As String
Dim strPackagerID As String
Dim strFirmPermissionsSeqNo As String
Dim strFirmClubSeqNo As String
' TW 18/11/2006 EP2_132
Dim strUnitID As String
' TW 18/11/2006 EP2_132 End

Private Sub PopulateTradingNamesListView()
Dim m_clsFirmTradingNameTable As New FirmTradingNameTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strARFirmID
    
    m_clsFirmTradingNameTable.SetFindLinkedBrokerTradingNameRecords strARFirmID
    lvTradingNames.LoadColumnDetails TypeName(m_clsFirmTradingNameTable)
    
    TableAccess(m_clsFirmTradingNameTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditBroker.lvTradingNames, m_clsFirmTradingNameTable
    
    cmdEditTradingName.Enabled = False
    cmdDeleteTradingName.Enabled = False
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub SetupTradingNameHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 0
    lvHeaders.sName = "Sequence"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Type"
    headers.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Effective Date"
    headers.Add lvHeaders

    lvHeaders.nWidth = 40
    lvHeaders.sName = "Alternative Name"
    headers.Add lvHeaders

    lvHeaders.nWidth = 0
    lvHeaders.sName = "" 'Dummy - "Alternative Name" value is not returned without this !
    headers.Add lvHeaders

    lvTradingNames.AddHeadings headers
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cboFirmStatus_Validate(Cancel As Boolean)
    Cancel = Not cboFirmStatus.ValidateData()
End Sub



Private Sub cboListingStatus_Validate(Cancel As Boolean)
    Cancel = Not cboListingStatus.ValidateData()
End Sub



Private Sub chkSelectedNetwork_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim clsTableAccess As TableAccess

Dim conn As ADODB.Connection
Dim recSet As ADODB.Recordset
    
Dim intValue As Integer
    
Dim strAppointmentID As String
Dim strSQL As String

    On Error GoTo Failed
    If CheckIfOKToSave() Then
        strAppointmentID = lvCurrentAssociations.SelectedItem.SubItems(5)
    
        If chkSelectedNetwork.Value = 1 Then
            
            ' If value is 1, set the MORTGAGENETWORKIND on the selected appointment record to 1
            ' and then set the MORTGAGENETWORKIND on all other linked appointment records to 0
            
            Set conn = New ADODB.Connection
        
            With conn
                .ConnectionString = g_clsDataAccess.GetActiveConnection
                .Open
            End With
    
            Set recSet = conn.Execute("SELECT APPOINTMENTID, ISNULL(MORTGAGENETWORKIND, 0) AS MORTGAGENETWORKIND FROM APPOINTMENTS WHERE APPOINTMENTID IN (SELECT APPOINTMENTID FROM APPOINTMENTS A INNER JOIN PRINCIPALFIRM M ON A.PRINCIPALFIRMID = M.PRINCIPALFIRMID WHERE A.ARFIRMID = '" & strARFirmID & "')")
            Do While Not recSet.EOF
                If recSet!APPOINTMENTID = strAppointmentID Then
                    intValue = 1
                Else
                    intValue = 0
                End If
                If recSet!MORTGAGENETWORKIND <> intValue Then
                    conn.Execute "UPDATE APPOINTMENTS SET MORTGAGENETWORKIND = " & intValue & " WHERE APPOINTMENTID = '" & recSet!APPOINTMENTID & "'"
                    Call SaveChangeRequest(recSet!APPOINTMENTID, 3)
                End If
                recSet.MoveNext
            Loop
            conn.Close
            Set conn = Nothing
            Set recSet = Nothing
        Else
            ' If value is 0, set the MORTGAGENETWORKIND on the selected appointment record to 0
            Set clsTableAccess = m_clsAppointmentsTable
            
            m_clsAppointmentsTable.SetMortgageNetworkInd 0
            clsTableAccess.Update
        
            Call SaveChangeRequest(strAppointmentID, 3)
        End If
        PopulateAppointmentsListView
    End If
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub cmdAddNetwork_Click()
Dim colKeys As New Collection
    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        colKeys.Add strARFirmID
        frmMaintainFirmLinks.SetKeys colKeys
        frmMaintainFirmLinks.SetIsEdit True
        frmMaintainFirmLinks.Show 1
        If frmMaintainFirmLinks.GetReturnCode() Then
            PopulateAppointmentsListView
        End If
        
        Unload frmMaintainFirmLinks
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvCurrentAssociations_Click()
' TW 12/03/2007 EP2_1190
Dim colFields As New Collection
Dim colMatchValues As New Collection

    On Error GoTo Failed:
    Set colFields = New Collection
    Set colMatchValues = New Collection

    colFields.Add "APPOINTMENTID"
    colMatchValues.Add lvCurrentAssociations.SelectedItem.SubItems(5)

    TableAccess(m_clsAppointmentsTable).SetKeyMatchValues colFields
    TableAccess(m_clsAppointmentsTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsAppointmentsTable).GetTableData POPULATE_KEYS
    
    If m_clsAppointmentsTable.GetMortgageNetworkInd = "Yes" Then
        chkSelectedNetwork.Value = 1
    Else
        chkSelectedNetwork.Value = 0
    End If
    
    chkSelectedNetwork.Enabled = True
' TW 12/03/2007 EP2_1190 End
    
    cmdRemoveNetwork.Enabled = True
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub lvTradingNames_Click()
    'AW OMIGA00003231
    If frmEditBroker.lvTradingNames.ListItems.Count > 0 Then
        cmdEditTradingName.Enabled = True
        cmdDeleteTradingName.Enabled = True
    Else
        cmdEditTradingName.Enabled = False
        cmdDeleteTradingName.Enabled = False
    End If
End Sub


Private Sub cmdEditTradingName_Click()
Dim colMatchValues As Collection
    
    Set colMatchValues = New Collection
    colMatchValues.Add lvTradingNames.getValueFromName("Sequence")
    
    frmEditBrokerTradingName.SetKeys colMatchValues
    frmEditBrokerTradingName.SetIsEdit (True)

    frmEditBrokerTradingName.Show 1
    If frmEditBrokerTradingName.GetReturnCode() Then
        PopulateTradingNamesListView
    End If
    Unload frmEditBrokerTradingName
End Sub

Private Sub cmdDeleteTradingName_Click()
Dim colMatchValues As Collection
Dim colValues As Collection
Dim intIndex As Integer
Dim clsTableAccess As TableAccess
Dim m_clsFirmTradingNameTable As New FirmTradingNameTable

    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add lvTradingNames.getValueFromName("Sequence")
    
    TableAccess(m_clsFirmTradingNameTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsFirmTradingNameTable).DeleteRecords
    
' TW 06/12/2006 EP2_330
    Set clsTableAccess = m_clsFirmTradingNameTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsFirmTradingNameTable, , PromoteDelete
' TW 06/12/2006 EP2_330 End
    
    PopulateTradingNamesListView
    
    Set colValues = Nothing
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub PopulateAppointmentsListView()
Dim m_clsAppointmentsTable As New AppointmentsTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    cmdRemoveNetwork.Enabled = False
' TW 12/03/2007 EP2_1190
    chkSelectedNetwork.Enabled = False
' TW 12/03/2007 EP2_1190 End

    Set colMatchValues = New Collection
    colMatchValues.Add strARFirmID

    m_clsAppointmentsTable.SetFindLinkedBrokerAppointmentRecords strARFirmID
    lvCurrentAssociations.LoadColumnDetails TypeName(m_clsAppointmentsTable)
    TableAccess(m_clsAppointmentsTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditBroker.lvCurrentAssociations, m_clsAppointmentsTable
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupAppointmentHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Firm Details"
    headers.Add lvHeaders

    lvHeaders.nWidth = 30
    lvHeaders.sName = "Address"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Agreed Proc Fee Rate"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Vol Proc Fee Adjustment"
    headers.Add lvHeaders

' TW 12/03/2007 EP2_1190
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Selected Network"
    headers.Add lvHeaders
' TW 12/03/2007 EP2_1190 End
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders

    lvHeaders.nWidth = 0
    lvHeaders.sName = "Sequence"
    headers.Add lvHeaders

    lvCurrentAssociations.AddHeadings headers
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub cmdAddTradingName_Click()
    If CheckIfOKToSave() Then
        frmEditBrokerTradingName.SetIsEdit (False)
        frmEditBrokerTradingName.txtPrincipalFirmID.Text = ""
        frmEditBrokerTradingName.txtARFirmID.Text = strARFirmID
        frmEditBrokerTradingName.Show 1
        If frmEditBrokerTradingName.GetReturnCode() Then
            PopulateTradingNamesListView
        End If
        Unload frmEditBrokerTradingName
    End If
End Sub

Private Function CheckIfOKToSave() As Boolean
Dim bRet As Boolean
    On Error GoTo Failed
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    If bRet Then
        SaveScreenData
    End If
    CheckIfOKToSave = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    CheckIfOKToSave = False
End Function


Private Sub cmdCreateAsOmigaUnit_Click()
    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        GetNewUnitID
' TW 09/01/2007 EP2_728
        frmEditUnit.SetIsEdit False
' TW 09/01/2007 EP2_728 End
        frmEditUnit.txtUnit(0).Text = strUnitID
        frmEditUnit.txtUnit(0).Enabled = False
        frmEditUnit.Show 1
        txtUnitID.Text = frmEditUnit.txtUnit(0).Text
        Unload frmEditUnit
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub


Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New ARFirmTable
    Set m_clsAppointmentsTable = New AppointmentsTable
    Set m_clsARFirmTable = New ARFirmTable
    
    g_clsFormProcessing.PopulateCombo "FirmStatus", Me.cboFirmStatus
    g_clsFormProcessing.PopulateCombo "ListingStatus", Me.cboListingStatus
    
    SetupTradingNameHeaders
    SetupAppointmentHeaders
    
    If m_bIsEdit = True Then
        strARFirmID = m_colKeys(1)
        SetEditState
    Else
        SetAddState
    End If
    
    cmdCreateAsOmigaUnit.Enabled = (Len(txtUnitID.Text) = 0)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub cmdRemoveNetwork_Click()
Dim X As Long
Dim colMatchValues As Collection
Dim colFields As Collection
Dim clsTableAccess As TableAccess
Dim strSQL As String

    On Error GoTo Failed

    Set clsTableAccess = m_clsAppointmentsTable

    If CheckIfOKToSave() Then
        For X = 1 To lvCurrentAssociations.ListItems.Count
            If lvCurrentAssociations.ListItems.Item(X).Selected Then

                Set colFields = New Collection
                Set colMatchValues = New Collection

                colFields.Add "APPOINTMENTID"
' TW 12/03/2007 EP2_1190
'                colMatchValues.Add lvCurrentAssociations.ListItems.Item(X).SubItems(4)
                colMatchValues.Add lvCurrentAssociations.ListItems.Item(X).SubItems(5)
' TW 12/03/2007 EP2_1190 End
                clsTableAccess.SetKeyMatchFields colFields
                clsTableAccess.SetKeyMatchValues colMatchValues

                clsTableAccess.DeleteRow colMatchValues
                clsTableAccess.Update

                g_clsHandleUpdates.SaveChangeRequest m_clsAppointmentsTable, , PromoteDelete
            End If
        Next X

        Call PopulateAppointmentsListView
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError


End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
' TW 05/02/2007 EP2_706
        If lvCurrentAssociations.ListItems.Count = 0 Then
            bRet = False
            MsgBox "There are no Network(s) associated with this AR Firm. Please make an association"
            SSTab1.Tab = 4
        Else
' TW 05/02/2007 EP2_706 End
            SaveScreenData
            SaveChangeRequest strARFirmID, 1
        End If
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub SaveChangeRequest(strKey As String, intClassType As Integer)
    On Error GoTo Failed
    Dim sProductNumber As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add strKey
    
    Select Case intClassType
        Case 1
            Set clsTableAccess = m_clsARFirmTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsARFirmTable
        Case 3
            Set clsTableAccess = m_clsAppointmentsTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsAppointmentsTable
    End Select
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
Dim vTmp As Variant
Dim clsTableAccess As TableAccess
Dim vNextNumber As Variant
    
    On Error GoTo Failed
    Set clsTableAccess = m_clsARFirmTable
    
    m_clsARFirmTable.SetARFirmID strARFirmID
    'Firm details tab
    m_clsARFirmTable.SetUnitID txtUnitID.Text
    m_clsARFirmTable.SetARFirmName txtARFirmName.Text
    If cboFirmStatus.ListIndex > -1 Then
        m_clsARFirmTable.SetFirmStatus cboFirmStatus.GetExtra(cboFirmStatus.ListIndex)
    End If
    
    If cboListingStatus.ListIndex > -1 Then
        m_clsARFirmTable.SetListingStatus cboListingStatus.GetExtra(cboListingStatus.ListIndex)
    End If
    If optIncludeOnlineSubmissionLoading(0).Value = True Then
        m_clsARFirmTable.SetProcFeeOnlineInd 1
    Else
        m_clsARFirmTable.SetProcFeeOnlineInd 0
    End If
      
    'Contact Details tab
    
    m_clsARFirmTable.SetAddressLine1 txtAddressLine1.Text
    m_clsARFirmTable.SetAddressLine2 txtAddressLine2.Text
    m_clsARFirmTable.SetAddressLine3 txtAddressLine3.Text
    m_clsARFirmTable.SetAddressLine4 txtAddressLine4.Text
    m_clsARFirmTable.SetAddressLine5 txtAddressLine5.Text
    m_clsARFirmTable.SetAddressLine6 txtAddressLine6.Text
    m_clsARFirmTable.SetPostcode txtPostCode.Text
    m_clsARFirmTable.SetTelephoneCountryCode txtCountryCode1.Text
    m_clsARFirmTable.SetTelephoneAreaCode txtAreaCode1.Text
    m_clsARFirmTable.SetTelephoneNumber txtTelephoneNumber1.Text
    m_clsARFirmTable.SetFaxCountryCode txtCountryCode2.Text
    m_clsARFirmTable.SetFaxAreaCode txtAreaCode2.Text
    m_clsARFirmTable.SetFaxNumber txtTelephoneNumber2.Text
' TW 18/11/2006 EP2_132
    m_clsARFirmTable.SetLastupdateddate Now
    m_clsARFirmTable.SetLastUpdatedBy g_sSupervisorUser
' TW 18/11/2006 EP2_132 End
      
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new Base Rate Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed
    
    Call GetNewARFirmID
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add strARFirmID
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsARFirmTable
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetNewUnitID()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo Failed

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETNEXTSEQUENCENUMBER"
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "UnitID")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strUnitID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strUnitID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub

Private Sub GetNewARFirmID()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo Failed

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETNEXTSEQUENCENUMBER"
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "ARFirmID")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strARFirmID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strARFirmID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing MortgageClubNetworkAssociation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
Dim colDataSet As New Collection

    colDataSet.Add m_colKeys.Item(1)
    TableAccess(m_clsARFirmTable).SetKeyMatchValues colDataSet
    TableAccess(m_clsARFirmTable).GetTableData POPULATE_KEYS
      
    'Firm details tab
' TW 16/01/2007 EP2_859
    txtFSAARFirmRef.Text = m_clsARFirmTable.GetFSAARFirmRef
' TW 16/01/2007 EP2_859 End
    txtUnitID.Text = m_clsARFirmTable.GetUnitId
    txtARFirmName.Text = m_clsARFirmTable.GetARFirmName()
' TW 13/02/2007 EP2_1334
'    g_clsFormProcessing.HandleComboIndex cboFirmStatus, m_clsARFirmTable.GetFirmStatus, SET_CONTROL_VALUE
'    g_clsFormProcessing.HandleComboIndex cboListingStatus, m_clsARFirmTable.GetListingStatus, SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboFirmStatus, m_clsARFirmTable.GetFirmStatus, SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboListingStatus, m_clsARFirmTable.GetListingStatus, SET_CONTROL_VALUE
' TW 13/02/2007 EP2_1334 End
    If Val(m_clsARFirmTable.GetProcFeeOnlineInd) = 1 Then
        optIncludeOnlineSubmissionLoading(0) = True
    Else
        optIncludeOnlineSubmissionLoading(1) = True
    End If
      
      
    'Contact Details tab
    txtAddressLine1.Text = m_clsARFirmTable.GetAddressLine1()
    txtAddressLine2.Text = m_clsARFirmTable.GetAddressLine2()
    txtAddressLine3.Text = m_clsARFirmTable.GetAddressLine3()
    txtAddressLine4.Text = m_clsARFirmTable.GetAddressLine4()
    txtAddressLine5.Text = m_clsARFirmTable.GetAddressLine5()
    txtAddressLine6.Text = m_clsARFirmTable.GetAddressLine6()
    txtPostCode.Text = m_clsARFirmTable.GetPostcode()
    
    txtCountryCode1.Text = m_clsARFirmTable.GetTelephoneCountryCode()
    txtAreaCode1.Text = m_clsARFirmTable.GetTelephoneAreaCode()
    txtTelephoneNumber1.Text = m_clsARFirmTable.GetTelephoneNumber()
    txtCountryCode2.Text = m_clsARFirmTable.GetFaxCountryCode()
    txtAreaCode2.Text = m_clsARFirmTable.GetFaxAreaCode()
    txtTelephoneNumber2.Text = m_clsARFirmTable.GetFaxNumber
    Call PopulateTradingNamesListView
    Call PopulateAppointmentsListView
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetARFirmTableClass(clsARFirmTable As ARFirmTable)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetARFirmTableClass
' Description   : Sets the Processing class to be used by this form for all events - i.e., this
'                 form will call into m_clsARFirmTable for all business processing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo Failed
    
    If clsARFirmTable Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "ARFirmTable table class emtpy"
    End If
    
    Set m_clsARFirmTable = clsARFirmTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, TypeName(Me) & ".SetARFirmTable " & Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Private Sub cmdOK_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user presses the OK button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bRet As Boolean
    On Error GoTo Failed
    
    'm_clsMortgageClubNetAssocTable.OK
    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Called when the user presses the Cancel button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_clsARFirmTable = Nothing
End Sub


Private Sub txtAreaCode1_Validate(Cancel As Boolean)
    Cancel = Not txtAreaCode1.ValidateData()
End Sub


Private Sub txtAreaCode2_Validate(Cancel As Boolean)
    Cancel = Not txtAreaCode2.ValidateData()
End Sub


Private Sub txtAddressLine2_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine2.ValidateData()
End Sub


Private Sub txtAddressLine1_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine1.ValidateData()
End Sub


Private Sub txtCountryCode1_Validate(Cancel As Boolean)
    Cancel = Not txtCountryCode1.ValidateData()
End Sub


Private Sub txtCountryCode2_Validate(Cancel As Boolean)
    Cancel = Not txtCountryCode2.ValidateData()
End Sub


Private Sub txtAddressLine5_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine5.ValidateData()
End Sub


Private Sub txtAddressLine3_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine3.ValidateData()
End Sub


Private Sub txtPostCode_Validate(Cancel As Boolean)
    Cancel = Not txtPostCode.ValidateData()
End Sub


Private Sub txtARFirmName_Validate(Cancel As Boolean)
    Cancel = Not txtARFirmName.ValidateData()
End Sub


Private Sub txtAddressLine4_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine4.ValidateData()
End Sub


Private Sub txtTelephoneNumber1_Validate(Cancel As Boolean)
    Cancel = Not txtTelephoneNumber1.ValidateData()
End Sub


Private Sub txtTelephoneNumber2_Validate(Cancel As Boolean)
    Cancel = Not txtTelephoneNumber2.ValidateData()
End Sub


Private Sub txtAddressLine6_Validate(Cancel As Boolean)
    Cancel = Not txtAddressLine6.ValidateData()
End Sub


Private Sub txtUnitID_Validate(Cancel As Boolean)
    Cancel = Not txtUnitID.ValidateData()
End Sub



