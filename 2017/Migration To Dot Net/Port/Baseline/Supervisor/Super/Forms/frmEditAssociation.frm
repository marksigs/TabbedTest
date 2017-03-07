VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditAssociation 
   Caption         =   "Add/Edit Association"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   Icon            =   "frmEditAssociation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9330
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   32
      Top             =   7320
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7980
      TabIndex        =   33
      Top             =   7320
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General Details"
      TabPicture(0)   =   "frmEditAssociation.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFindExclusives"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Bank Details"
      TabPicture(1)   =   "frmEditAssociation.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Proc Fee Details"
      TabPicture(2)   =   "frmEditAssociation.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdFindExclusives 
         Caption         =   "Find Exclusives"
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Frame Frame8 
         Caption         =   "Telephone Contact Details"
         Height          =   1575
         Left            =   240
         TabIndex        =   59
         Top             =   4680
         Width           =   4215
         Begin MSGOCX.MSGEditBox txtCountryCode1 
            Height          =   285
            Left            =   2160
            TabIndex        =   14
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
            Left            =   2160
            TabIndex        =   15
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
            Left            =   2160
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
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
            TabIndex        =   62
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   61
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   23
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Fax Contact Details"
         Height          =   1575
         Left            =   4680
         TabIndex        =   55
         Top             =   4680
         Width           =   4215
         Begin MSGOCX.MSGEditBox txtCountryCode2 
            Height          =   285
            Left            =   2160
            TabIndex        =   17
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
            Left            =   2160
            TabIndex        =   18
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
            Left            =   2160
            TabIndex        =   19
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
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
            Index           =   11
            Left            =   240
            TabIndex        =   58
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   57
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Proc Fee Details"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   52
         Top             =   480
         Width           =   8655
         Begin VB.OptionButton optIncludeLoading 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   27
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton optIncludeLoading 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   28
            Top             =   720
            Value           =   -1  'True
            Width           =   735
         End
         Begin MSGOCX.MSGComboBox cboProcFeePaymentMethod 
            Height          =   315
            Left            =   3240
            TabIndex        =   29
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
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
         Begin MSGOCX.MSGEditBox txtAssociationFeeAmount 
            Height          =   285
            Left            =   3240
            TabIndex        =   30
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAssociationFeeRate 
            Height          =   285
            Left            =   3240
            TabIndex        =   31
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAgreedProcFeeRate 
            Height          =   285
            Left            =   3240
            TabIndex        =   26
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            TextType        =   2
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MinValue        =   "-99.99"
            MaxValue        =   "99.99"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Adjustment to Basic Proc Fee"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Association Fee Rate %"
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   64
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Association Fee Value"
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   63
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Include Loan Amount / LTV Loading ?"
            Height          =   195
            Index           =   35
            Left            =   240
            TabIndex        =   54
            Top             =   720
            Width           =   2730
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proc Fee Payment Method"
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   53
            Top             =   1080
            Width           =   1890
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bank Acount"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton cmdBankWizard 
            Caption         =   "Bank Wizard"
            Height          =   375
            Left            =   4560
            TabIndex        =   25
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox chkBankWizardIndicator 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Wizard Indicator"
            Enabled         =   0   'False
            Height          =   255
            Left            =   210
            TabIndex        =   24
            Top             =   1800
            Width           =   2010
         End
         Begin MSGOCX.MSGEditBox txtAccountName 
            Height          =   285
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   6375
            _ExtentX        =   11245
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
         Begin MSGOCX.MSGEditBox txtAccountNumber 
            Height          =   285
            Left            =   2040
            TabIndex        =   21
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtBranchName 
            Height          =   285
            Left            =   2040
            TabIndex        =   22
            Top             =   1080
            Width           =   6375
            _ExtentX        =   11245
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
         Begin MSGOCX.MSGEditBox txtSortCode 
            Height          =   285
            Left            =   2040
            TabIndex        =   23
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sort Code"
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   51
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Branch Name"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   50
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Number"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   49
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Details"
         Height          =   4095
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   8655
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   5160
            TabIndex        =   13
            Top             =   3600
            Width           =   3135
            _ExtentX        =   5530
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
         Begin VB.OptionButton optPackagerIndicator 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   7800
            TabIndex        =   2
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optPackagerIndicator 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   7080
            TabIndex        =   1
            Top             =   360
            Width           =   615
         End
         Begin MSGOCX.MSGEditBox txtAssociationName 
            Height          =   285
            Left            =   2160
            TabIndex        =   0
            Top             =   360
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSGOCX.MSGEditBox txtAssociationDescription 
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Top             =   720
            Width           =   6135
            _ExtentX        =   10821
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
         Begin MSGOCX.MSGEditBox txtPostCode 
            Height          =   285
            Left            =   2160
            TabIndex        =   12
            Top             =   3600
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
         Begin MSGOCX.MSGEditBox txtBuildingName 
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Top             =   1800
            Width           =   4335
            _ExtentX        =   7646
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
         Begin MSGOCX.MSGEditBox txtBuildingNumber 
            Height          =   285
            Left            =   7200
            TabIndex        =   7
            Top             =   1800
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
         Begin MSGOCX.MSGEditBox txtStreet 
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Top             =   2160
            Width           =   6135
            _ExtentX        =   10821
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
         Begin MSGOCX.MSGEditBox txtDistrict 
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            Top             =   2520
            Width           =   6135
            _ExtentX        =   10821
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
         Begin MSGOCX.MSGEditBox txtTown 
            Height          =   285
            Left            =   2160
            TabIndex        =   10
            Top             =   2880
            Width           =   6135
            _ExtentX        =   10821
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
         Begin MSGOCX.MSGEditBox txtCounty 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Top             =   3240
            Width           =   6135
            _ExtentX        =   10821
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
         Begin MSGOCX.MSGEditBox txtFlatNumber 
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Top             =   1440
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
         Begin MSGOCX.MSGComboBox cboFirmIdentifier 
            Height          =   315
            Left            =   2160
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Identifier"
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   67
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Flat Number"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   65
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No."
            Height          =   195
            Index           =   10
            Left            =   6720
            TabIndex        =   46
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Packager Indicator"
            Height          =   195
            Index           =   9
            Left            =   5520
            TabIndex        =   45
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            Height          =   195
            Index           =   8
            Left            =   4200
            TabIndex        =   44
            Top             =   3600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "County"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   43
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Town"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   42
            Top             =   2880
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "District"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   41
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Street"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   40
            Top             =   2160
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Building Name"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   39
            Top             =   1800
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Postcode"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   3600
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Association Description"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Association Name"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   1275
         End
      End
   End
End
Attribute VB_Name = "frmEditAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditAssociation
' Description   : To Add and Edit Mortgage Club Network Associations
'
' Change history
' Prog      Date        Description
' TW        17/10/2006  EP2_15 - Created
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        07/12/2006  EP2_153 - Error Adding Association
'                               - Made Bank Account Number and Sort code mandatory
' TW        07/12/2006  EP2_359 - Supervisor E2CR24_26 - Changes for Firm Identifier
' TW        21/12/2006  EP2_626 - Make Firm Identifier available for Clubs only
' TW        29/01/2007  EP2_1034 - Proc fee calculation processing to allow for firm-level basic rate adjustment
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Private data
Private m_clsMortgageClubNetAssocTable As MortgageClubNetAssocTable

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_colKeys As Collection

Private strClubNetworkAssociationID As String

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


Private Sub cboCountry_Validate(Cancel As Boolean)
    Cancel = Not cboCountry.ValidateData()
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


Private Sub cboProcFeePaymentMethod_Validate(Cancel As Boolean)
    Cancel = Not cboProcFeePaymentMethod.ValidateData()
End Sub


Private Sub cmdBankWizard_Click()
    
Dim bRet As Boolean
Dim sSortCode As String
Dim sBankAccountNumber As String
Dim xdcBankWizardDoc As FreeThreadedDOMDocument
Dim strXMLData As String
Dim strError As String
Dim xndDescription As IXMLDOMNode
Dim iCount As Integer ' EP511

    Const strFunctionName As String = "cmdBankWizard_Click"
    
    On Error GoTo Failed
        
    sSortCode = txtSortCode.Text
    sBankAccountNumber = txtAccountNumber.Text
    chkBankWizardIndicator.Value = 0
    If sSortCode = "" Or sBankAccountNumber = "" Then
        MsgBox "A Bank Sort Code and Bank Account Number must be entered"
    Else
        'Call bank wizard with BankSortCode and BankAccountNumber
        strXMLData = ValidateBankDetails(sSortCode, sBankAccountNumber)
    
        Set xdcBankWizardDoc = New FreeThreadedDOMDocument
        xdcBankWizardDoc.async = False         ' wait until XML is fully loaded
        xdcBankWizardDoc.validateOnParse = False
        xdcBankWizardDoc.loadXML strXMLData
    
        Select Case xdcBankWizardDoc.selectSingleNode("RESPONSE/@TYPE").Text
            Case "SYSERR"
                Set xndDescription = xdcBankWizardDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION")
                MsgBox "System error when trying to run Bank Wizard error - " + xndDescription.Text
            Case "BANKWIZARDERROR"
    
                Select Case xdcBankWizardDoc.selectSingleNode("RESPONSE/ERROR/NUMBER").Text
                    Case 2
                        MsgBox "Sort-code not allocated to any bank branch"
                    Case 4
                        MsgBox "Unknown Account Number Specified"
                    Case 32
                        MsgBox "Account number format incorrect"
                    Case 64
                        MsgBox "Sort-code format incorrect"
                    Case Else
                        MsgBox "Unidentified error from Bank Wizard"
                End Select
            Case "APPERR"
                Set xndDescription = xdcBankWizardDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION")
                If xndDescription Is Nothing Then
                    MsgBox "Bank Wizard failed to complete the search without returning an error"
                Else
                    MsgBox "Bank Wizard error - " + xndDescription.Text
                End If
            Case Else
                chkBankWizardIndicator.Value = 1
        End Select
    End If
    
    Set xdcBankWizardDoc = Nothing
    Set xndDescription = Nothing
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateBankDetails
' Description   : Validate bank details
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateBankDetails(ByVal sSortCode As String, ByVal sBankAccountNumber As String) As String
Dim strXMLData As String
Dim xmlDoc As New MSXML2.DOMDocument
Dim clsOmiga4 As New Omiga4Support
Dim sResponse As String

    On Error GoTo Failed

    xmlDoc.async = False
    xmlDoc.loadXML ("<REQUEST><BANKDETAILS SORTCODE=""" + sSortCode + """ ACCOUNTNUMBER=""" + sBankAccountNumber + """/></REQUEST>")
    
    sResponse = clsOmiga4.RunASP(xmlDoc.xml, "GetBankDetails.asp")

    ValidateBankDetails = sResponse
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub cmdFindExclusives_Click()
Dim colKeys As New Collection
    On Error GoTo Failed:
    colKeys.Add strClubNetworkAssociationID
    colKeys.Add "Introducer"
    colKeys.Add frmMain.tvwDB.SelectedItem
    frmMaintainProductExclusivity.SetKeys colKeys
    frmMaintainProductExclusivity.SetIsEdit True
    frmMaintainProductExclusivity.Show 1
    
    Unload frmMaintainProductExclusivity
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Initialize()
    Set m_colKeys = New Collection
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New MortgageClubNetAssocTable
    
    Set m_clsMortgageClubNetAssocTable = New MortgageClubNetAssocTable
    g_clsFormProcessing.PopulateCombo "Country", Me.cboCountry
' TW 07/12/2006 EP2_153
    g_clsFormProcessing.PopulateCombo "ProcurationFeePaymentMethod", Me.cboProcFeePaymentMethod
' TW 07/12/2006 EP2_153 End
' TW 21/12/2006 EP2_626
'' TW 07/12/2006 EP2_359
'    g_clsFormProcessing.PopulateCombo "MortgageClubIdType", Me.cboFirmIdentifier
'' TW 07/12/2006 EP2_359 End
' TW 21/12/2006 EP2_626 End
    
    Select Case UCase$(frmMain.tvwDB.SelectedItem)
        Case "ASSOCIATIONS"
            Me.Caption = "Add/Edit Association"
            Label1(0).Caption = "Association Name"
            Label1(1).Caption = "Association Description"
            optPackagerIndicator(0).Value = True
' TW 21/12/2006 EP2_626
            cboFirmIdentifier.Enabled = False
' TW 21/12/2006 EP2_626 End
        Case "CLUBS"
            Me.Caption = "Add/Edit Club"
            Label1(0).Caption = "Club Name"
            Label1(1).Caption = "Club Description"
            optPackagerIndicator(1).Value = True
' TW 21/12/2006 EP2_626
            g_clsFormProcessing.PopulateCombo "MortgageClubIdType", Me.cboFirmIdentifier
' TW 21/12/2006 EP2_626 End
    End Select
    optPackagerIndicator(0).Enabled = False
    optPackagerIndicator(1).Enabled = False
    
' TW 23/11/2006 EP2_172
    cmdFindExclusives.Visible = m_bIsEdit
' TW 23/11/2006 EP2_172 End
    
    If m_bIsEdit = True Then
        strClubNetworkAssociationID = m_colKeys(1)
        SetEditState
    Else
        SetAddState
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing MortgageClubNetworkAssociation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
Dim colDataSet As New Collection
    colDataSet.Add m_colKeys.Item(1)
    TableAccess(m_clsMortgageClubNetAssocTable).SetKeyMatchValues colDataSet
    TableAccess(m_clsMortgageClubNetAssocTable).GetTableData POPULATE_KEYS
      
    txtAssociationName.Text = m_clsMortgageClubNetAssocTable.GetMortgageClubNetworkAssocName()
    If Val(m_clsMortgageClubNetAssocTable.GetPackagerIndicator()) = 1 Then
        optPackagerIndicator(0) = True
    Else
        optPackagerIndicator(1) = True
    End If
    
    txtAssociationDescription.Text = m_clsMortgageClubNetAssocTable.GetMortgageClubNetworkAssocDescription()
' TW 21/12/2006 EP2_626
'' TW 07/12/2006 EP2_359
'    g_clsFormProcessing.HandleComboExtra cboFirmIdentifier, m_clsMortgageClubNetAssocTable.GetIdentifier, SET_CONTROL_VALUE
'' TW 07/12/2006 EP2_359 End
' TW 21/12/2006 EP2_626 End
' TW 21/12/2006 EP2_626
    If optPackagerIndicator(1).Value = True Then
        g_clsFormProcessing.HandleComboExtra cboFirmIdentifier, m_clsMortgageClubNetAssocTable.GetIdentifier, SET_CONTROL_VALUE
    End If
' TW 21/12/2006 EP2_626 End
    
    txtPostCode.Text = m_clsMortgageClubNetAssocTable.GetPostcode()
    txtBuildingName.Text = m_clsMortgageClubNetAssocTable.GetBuildingOrHouseName()
    txtBuildingNumber.Text = m_clsMortgageClubNetAssocTable.GetBuildingOrHouseNumber()
    txtFlatNumber.Text = m_clsMortgageClubNetAssocTable.GetFlatNumber()
    txtStreet.Text = m_clsMortgageClubNetAssocTable.GetStreet()
    txtDistrict.Text = m_clsMortgageClubNetAssocTable.GetDistrict()
    txtTown.Text = m_clsMortgageClubNetAssocTable.GetTown()
    txtCounty.Text = m_clsMortgageClubNetAssocTable.GetCounty()
    g_clsFormProcessing.HandleComboExtra cboCountry, m_clsMortgageClubNetAssocTable.GetCountry, SET_CONTROL_VALUE
    
    txtCountryCode1.Text = m_clsMortgageClubNetAssocTable.GetTelephoneCountryCode()
    txtAreaCode1.Text = m_clsMortgageClubNetAssocTable.GetTelephoneAreaCode()
    txtTelephoneNumber1.Text = m_clsMortgageClubNetAssocTable.GetTelephoneNumber()
    txtCountryCode2.Text = m_clsMortgageClubNetAssocTable.GetFaxCountryCode()
    txtAreaCode2.Text = m_clsMortgageClubNetAssocTable.GetFaxAreaCode()
    txtTelephoneNumber2.Text = m_clsMortgageClubNetAssocTable.GetFaxNumber()
    
    txtAccountName.Text = m_clsMortgageClubNetAssocTable.GetBankAccountName()
    txtAccountNumber.Text = m_clsMortgageClubNetAssocTable.GetBankAccountNumber()
    txtBranchName.Text = m_clsMortgageClubNetAssocTable.GetBankAccountBranchName()
    txtSortCode.Text = m_clsMortgageClubNetAssocTable.GetBankSortCode()
    chkBankWizardIndicator.Value = Val(m_clsMortgageClubNetAssocTable.GetBankWizardIndicator())

    If Val(m_clsMortgageClubNetAssocTable.GetProcLoadingInd) = 1 Then
        optIncludeLoading(0) = True
    Else
        optIncludeLoading(1) = True
    End If
' TW 07/12/2006 EP2_153
    g_clsFormProcessing.PopulateCombo "ProcurationFeePaymentMethod", Me.cboProcFeePaymentMethod
    g_clsFormProcessing.HandleComboExtra cboProcFeePaymentMethod, m_clsMortgageClubNetAssocTable.GetPaymentMethod, SET_CONTROL_VALUE
    
    txtAssociationFeeAmount.Text = m_clsMortgageClubNetAssocTable.GetAssociationFeeAmount
    txtAssociationFeeRate.Text = m_clsMortgageClubNetAssocTable.GetAssociationFeeRate
' TW 07/12/2006 EP2_153 End
' TW 29/01/2007 EP2_1034
    txtAgreedProcFeeRate.Text = m_clsMortgageClubNetAssocTable.GetAgreedProcFeeRate
' TW 29/01/2007 EP2_1034 End
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new Base Rate Set
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed
    
    Call GetNewClubNetworkAssociationID
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add strClubNetworkAssociationID
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsMortgageClubNetAssocTable
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetNewClubNetworkAssociationID()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "ClubNetworkAssociationID")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strClubNetworkAssociationID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strClubNetworkAssociationID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:

End Sub


Public Sub SetMortgageClubNetAssocTableClass(clsMortgageClubNetAssocTable As MortgageClubNetAssocTable)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetMortgageClubNetAssocTableClass
' Description   : Sets the Processing class to be used by this form for all events - i.e., this
'                 form will call into m_clsMortgageClubNetAssocTable for all business processing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo Failed
    
    If clsMortgageClubNetAssocTable Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "MortgageClubNetAssocTable table class emtpy"
    End If
    
    Set m_clsMortgageClubNetAssocTable = clsMortgageClubNetAssocTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, TypeName(Me) & ".SetMortgageClubNetAssocTable " & Err.DESCRIPTION
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
       bRet = g_clsValidation.ValidateSortCode(txtSortCode)
    End If
    If bRet = True Then
        SaveScreenData
        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
Dim clsTableAccess As TableAccess

    On Error GoTo Failed
    
    Set clsTableAccess = m_clsMortgageClubNetAssocTable
    
    m_clsMortgageClubNetAssocTable.SetClubNetworkAssociationId strClubNetworkAssociationID
    m_clsMortgageClubNetAssocTable.SetMortgageClubNetworkAssocName txtAssociationName.Text
    If optPackagerIndicator(0).Value = True Then
        m_clsMortgageClubNetAssocTable.SetPackagerIndicator 1
    Else
        m_clsMortgageClubNetAssocTable.SetPackagerIndicator 0
    End If
    m_clsMortgageClubNetAssocTable.SetMortgageClubNetworkAssocDescription txtAssociationDescription.Text
' TW 21/12/2006 EP2_626
'' TW 07/12/2006 EP2_359
'    If cboFirmIdentifier.ListIndex > -1 Then
'        m_clsMortgageClubNetAssocTable.SetIdentifier cboFirmIdentifier.GetExtra(cboFirmIdentifier.ListIndex)
'    End If
'' TW 07/12/2006 EP2_359 End
' TW 21/12/2006 EP2_626 End
' TW 21/12/2006 EP2_626
    If optPackagerIndicator(1).Value = True Then
        If cboFirmIdentifier.ListIndex > -1 Then
            m_clsMortgageClubNetAssocTable.SetIdentifier cboFirmIdentifier.GetExtra(cboFirmIdentifier.ListIndex)
        End If
    End If
' TW 21/12/2006 EP2_626 End
    m_clsMortgageClubNetAssocTable.SetPostcode txtPostCode.Text
    m_clsMortgageClubNetAssocTable.SetBuildingOrHouseName txtBuildingName.Text
    m_clsMortgageClubNetAssocTable.SetBuildingOrHouseNumber txtBuildingNumber.Text
    m_clsMortgageClubNetAssocTable.SetFlatNumber txtFlatNumber.Text
    m_clsMortgageClubNetAssocTable.SetStreet txtStreet.Text
    m_clsMortgageClubNetAssocTable.SetDistrict txtDistrict.Text
    m_clsMortgageClubNetAssocTable.SetTown txtTown.Text
    m_clsMortgageClubNetAssocTable.SetCounty txtCounty.Text
    If cboCountry.ListIndex > -1 Then
        m_clsMortgageClubNetAssocTable.SetCountry cboCountry.GetExtra(cboCountry.ListIndex)
    End If
    m_clsMortgageClubNetAssocTable.SetTelephoneCountryCode txtCountryCode1.Text
    m_clsMortgageClubNetAssocTable.SetTelephoneAreaCode txtAreaCode1.Text
    m_clsMortgageClubNetAssocTable.SetTelephoneNumber txtTelephoneNumber1.Text
    m_clsMortgageClubNetAssocTable.SetFaxCountryCode txtCountryCode2.Text
    m_clsMortgageClubNetAssocTable.SetFaxAreaCode txtAreaCode2.Text
    m_clsMortgageClubNetAssocTable.SetFaxNumber txtTelephoneNumber2.Text
    
    m_clsMortgageClubNetAssocTable.SetBankAccountName txtAccountName.Text
    m_clsMortgageClubNetAssocTable.SetBankAccountNumber txtAccountNumber.Text
    m_clsMortgageClubNetAssocTable.SetBankAccountBranchName txtBranchName.Text
    m_clsMortgageClubNetAssocTable.SetBankSortCode txtSortCode.Text
    m_clsMortgageClubNetAssocTable.SetBankWizardIndicator chkBankWizardIndicator.Value
    
    If optIncludeLoading(0).Value = True Then
        m_clsMortgageClubNetAssocTable.SetProcLoadingInd 1
    Else
        m_clsMortgageClubNetAssocTable.SetProcLoadingInd 0
    End If

    m_clsMortgageClubNetAssocTable.SetLastUpdatedBy g_sSupervisorUser
    m_clsMortgageClubNetAssocTable.SetLastUpdatedDate Now
    If cboProcFeePaymentMethod.ListIndex > -1 Then
        m_clsMortgageClubNetAssocTable.SetPaymentMethod cboProcFeePaymentMethod.GetExtra(cboProcFeePaymentMethod.ListIndex)
    End If
    m_clsMortgageClubNetAssocTable.SetAssociationFeeAmount txtAssociationFeeAmount.Text
    m_clsMortgageClubNetAssocTable.SetAssociationFeeRate txtAssociationFeeRate.Text

' TW 29/01/2007 EP2_1034
    m_clsMortgageClubNetAssocTable.SetAgreedProcFeeRate Val(Format$(txtAgreedProcFeeRate.Text))
' TW 29/01/2007 EP2_1034 End
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sProductNumber As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add strClubNetworkAssociationID
    
    Set clsTableAccess = m_clsMortgageClubNetAssocTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsMortgageClubNetAssocTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
    Set m_clsMortgageClubNetAssocTable = Nothing
End Sub

Private Sub txtAccountName_Validate(Cancel As Boolean)
    Cancel = Not txtAccountName.ValidateData()
End Sub


Private Sub txtAccountNumber_Validate(Cancel As Boolean)
    Cancel = Not txtAccountNumber.ValidateData()
End Sub


Private Sub txtAgreedProcFeeRate_Validate(Cancel As Boolean)
    Cancel = Not txtAgreedProcFeeRate.ValidateData()
End Sub


Private Sub txtAreaCode1_Validate(Cancel As Boolean)
    Cancel = Not txtAreaCode1.ValidateData()
End Sub


Private Sub txtAreaCode2_Validate(Cancel As Boolean)
    Cancel = Not txtAreaCode2.ValidateData()
End Sub


Private Sub txtAssociationDescription_Validate(Cancel As Boolean)
    Cancel = Not txtAssociationDescription.ValidateData()
End Sub



Private Sub txtAssociationFeeAmount_Change()
    If Len(txtAssociationFeeAmount.Text) > 0 Then
        txtAssociationFeeRate.Text = ""
    End If
End Sub

Private Sub txtAssociationFeeAmount_Validate(Cancel As Boolean)
    Cancel = Not txtAssociationFeeAmount.ValidateData()
End Sub


Private Sub txtAssociationFeeRate_Change()
    If Len(txtAssociationFeeRate.Text) > 0 Then
        txtAssociationFeeAmount.Text = ""
    End If
End Sub

Private Sub txtAssociationFeeRate_Validate(Cancel As Boolean)
    Cancel = Not txtAssociationFeeRate.ValidateData()
End Sub


Private Sub txtAssociationName_Validate(Cancel As Boolean)
    Cancel = Not txtAssociationName.ValidateData()
End Sub


Private Sub txtBranchName_Validate(Cancel As Boolean)
    Cancel = Not txtBranchName.ValidateData()
End Sub


Private Sub txtBuildingName_Validate(Cancel As Boolean)
    Cancel = Not txtBuildingName.ValidateData()
End Sub


Private Sub txtBuildingNumber_Validate(Cancel As Boolean)
    Cancel = Not txtBuildingNumber.ValidateData()
End Sub


Private Sub txtCountryCode1_Validate(Cancel As Boolean)
    Cancel = Not txtCountryCode1.ValidateData()
End Sub


Private Sub txtCountryCode2_Validate(Cancel As Boolean)
    Cancel = Not txtCountryCode2.ValidateData()
End Sub


Private Sub txtCounty_Validate(Cancel As Boolean)
    Cancel = Not txtCounty.ValidateData()
End Sub


Private Sub txtDistrict_Validate(Cancel As Boolean)
    Cancel = Not txtDistrict.ValidateData()
End Sub


Private Sub txtFlatNumber_Validate(Cancel As Boolean)
    Cancel = Not txtFlatNumber.ValidateData()
End Sub


Private Sub txtPostCode_Validate(Cancel As Boolean)
    Cancel = Not txtPostCode.ValidateData()
End Sub


Private Sub txtSortCode_Validate(Cancel As Boolean)
    Cancel = Not txtSortCode.ValidateData()
End Sub


Private Sub txtStreet_Validate(Cancel As Boolean)
    Cancel = Not txtStreet.ValidateData()
End Sub


Private Sub txtTelephoneNumber1_Validate(Cancel As Boolean)
    Cancel = Not txtTelephoneNumber1.ValidateData()
End Sub


Private Sub txtTelephoneNumber2_Validate(Cancel As Boolean)
    Cancel = Not txtTelephoneNumber2.ValidateData()
End Sub


Private Sub txtTown_Validate(Cancel As Boolean)
    Cancel = Not txtTown.ValidateData()
End Sub


