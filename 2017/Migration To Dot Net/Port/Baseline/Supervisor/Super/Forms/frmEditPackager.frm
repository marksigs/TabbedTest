VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPackager 
   Caption         =   "Add/Edit Principal Firm"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "frmEditPackager.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   6855
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8220
      TabIndex        =   48
      Top             =   6360
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   47
      Top             =   6360
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Firm Details"
      TabPicture(0)   =   "frmEditPackager.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdFindExclusives"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contact Details"
      TabPicture(1)   =   "frmEditPackager.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Bank Details"
      TabPicture(2)   =   "frmEditPackager.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Firm Activities"
      TabPicture(3)   =   "frmEditPackager.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Trading Names"
      TabPicture(4)   =   "frmEditPackager.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Associations"
      TabPicture(5)   =   "frmEditPackager.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(1)=   "Frame5"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton cmdFindExclusives 
         Caption         =   "Find Exclusives"
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Available Associations"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   86
         Top             =   480
         Width           =   8895
         Begin MSGOCX.MSGListView lvAvailableAssociations 
            Height          =   2055
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3625
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Current Activities"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   85
         Top             =   3240
         Width           =   8895
         Begin VB.CommandButton cmdAddActivity 
            Caption         =   "&Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   2040
            Width           =   1275
         End
         Begin VB.CommandButton cmdRemoveActivity 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   38
            Top             =   2040
            Width           =   1275
         End
         Begin MSGOCX.MSGListView lvCurrentActivities 
            Height          =   1575
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2778
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Categories"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   84
         Top             =   480
         Width           =   8895
         Begin MSGOCX.MSGListView lvCategories 
            Height          =   2055
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   3625
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin MSGOCX.MSGListView lvAvailableActivities 
            Height          =   2055
            Left            =   3480
            TabIndex        =   35
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3625
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Linked Associations"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   83
         Top             =   3240
         Width           =   8895
         Begin VB.CommandButton cmdAddAssociation 
            Caption         =   "&Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   2040
            Width           =   1275
         End
         Begin VB.CommandButton cmdRemoveAssociation 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   46
            Top             =   2040
            Width           =   1275
         End
         Begin MSGOCX.MSGListView lvCurrentAssociations 
            Height          =   1575
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2778
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Trading Names"
         Height          =   5415
         Left            =   -74760
         TabIndex        =   82
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdDeleteTradingName 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3120
            TabIndex        =   42
            Top             =   4800
            Width           =   1275
         End
         Begin VB.CommandButton cmdEditTradingName 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   41
            Top             =   4800
            Width           =   1275
         End
         Begin VB.CommandButton cmdAddTradingName 
            Caption         =   "&Add"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   4800
            Width           =   1275
         End
         Begin MSGOCX.MSGListView lvTradingNames 
            Height          =   4215
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7435
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Fax Contact Details"
         Height          =   1575
         Left            =   -70200
         TabIndex        =   78
         Top             =   3600
         Width           =   4335
         Begin MSGOCX.MSGEditBox txtCountryCode2 
            Height          =   285
            Left            =   1560
            TabIndex        =   25
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
            Left            =   1560
            TabIndex        =   26
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
            Left            =   1560
            TabIndex        =   27
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
            TabIndex        =   81
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   80
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Telephone Contact Details"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   74
         Top             =   3600
         Width           =   4335
         Begin MSGOCX.MSGEditBox txtCountryCode1 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
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
            TabIndex        =   23
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
            TabIndex        =   24
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
            TabIndex        =   77
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   76
            Top             =   720
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   75
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Address Details"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   55
         Top             =   480
         Width           =   8895
         Begin MSGOCX.MSGEditBox txtPostCode 
            Height          =   285
            Left            =   1440
            TabIndex        =   21
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   15
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
            Caption         =   "Line 1"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Postcode"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   61
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 2"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   60
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 4"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   59
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 5"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   58
            Top             =   1800
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 6"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   57
            Top             =   2160
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line 3"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   56
            Top             =   1080
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Procuration Fee Adjustments"
         Height          =   1695
         Left            =   240
         TabIndex        =   72
         Top             =   3600
         Width           =   8895
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3240
            ScaleHeight     =   255
            ScaleWidth      =   2655
            TabIndex        =   88
            Top             =   1080
            Width           =   2655
            Begin VB.OptionButton optIncludeOnlineSubmissionLoading 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton optIncludeOnlineSubmissionLoading 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   14
               Top             =   0
               Value           =   -1  'True
               Width           =   735
            End
         End
         Begin VB.OptionButton optIncludeLoading 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   12
            Top             =   720
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optIncludeLoading 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   11
            Top             =   720
            Width           =   735
         End
         Begin MSGOCX.MSGEditBox txtAgreedProcFeeRate 
            Height          =   285
            Left            =   3240
            TabIndex        =   10
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
            Index           =   15
            Left            =   240
            TabIndex        =   90
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Include Online Submission Loading ?"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   87
            Top             =   1080
            Width           =   2610
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Include Loan Amount / LTV Loading ?"
            Height          =   195
            Index           =   35
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   2730
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Firm Details"
         Height          =   3015
         Left            =   240
         TabIndex        =   63
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtPrincipalFirmName 
            Height          =   285
            Left            =   2280
            MaxLength       =   130
            TabIndex        =   2
            Top             =   1080
            Width           =   6255
         End
         Begin VB.PictureBox FSAReferencePanel 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   4800
            ScaleHeight     =   285
            ScaleWidth      =   3855
            TabIndex        =   91
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
            Begin MSGOCX.MSGEditBox txtFSARef 
               Height          =   285
               Left            =   1800
               TabIndex        =   1
               Top             =   0
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
               Left            =   0
               TabIndex        =   92
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CommandButton cmdCreateAsOmigaUnit 
            Caption         =   "Create as Omiga Unit"
            Height          =   375
            Left            =   6600
            TabIndex        =   9
            Top             =   2400
            Width           =   1935
         End
         Begin VB.OptionButton optPackagerIndicator 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   7920
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.OptionButton optPackagerIndicator 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   7200
            TabIndex        =   4
            Top             =   1440
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSGOCX.MSGEditBox txtUnitID 
            Height          =   285
            Left            =   2280
            TabIndex        =   0
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            Enabled         =   0   'False
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
         Begin MSGOCX.MSGComboBox cboFirmStatus 
            Height          =   315
            Left            =   2280
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
         Begin MSGOCX.MSGComboBox cboFirmLegalStatus 
            Height          =   315
            Left            =   2280
            TabIndex        =   6
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
         Begin MSGOCX.MSGComboBox cboFirmType 
            Height          =   315
            Left            =   2280
            TabIndex        =   7
            Top             =   2160
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
            Left            =   2280
            TabIndex        =   8
            Top             =   2520
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
         Begin MSGOCX.MSGEditBox txtId 
            Height          =   285
            Left            =   2280
            TabIndex        =   94
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            Enabled         =   0   'False
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
            Height          =   195
            Left            =   240
            TabIndex        =   93
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Packager ?"
            Height          =   195
            Index           =   31
            Left            =   6120
            TabIndex        =   71
            Top             =   1440
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Listing Status"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   69
            Top             =   2520
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Type"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   68
            Top             =   2160
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Legal Status"
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   67
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Status"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   66
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Firm Name"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unit ID"
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   64
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bank Acount"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   50
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdBankWizard 
            Caption         =   "Bank Wizard"
            Height          =   375
            Left            =   4560
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox chkBankWizardIndicator 
            Alignment       =   1  'Right Justify
            Caption         =   "Bank Wizard Indicator"
            Height          =   255
            Left            =   210
            TabIndex        =   32
            Top             =   1800
            Width           =   2010
         End
         Begin MSGOCX.MSGEditBox txtAccountName 
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Top             =   360
            Width           =   6615
            _ExtentX        =   11668
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
            TabIndex        =   29
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
            TabIndex        =   30
            Top             =   1080
            Width           =   6615
            _ExtentX        =   11668
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
            TabIndex        =   31
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
            TabIndex        =   54
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Branch Name"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   53
            Top             =   1080
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Number"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   52
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1065
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Packager Indicator"
      Height          =   195
      Index           =   30
      Left            =   4800
      TabIndex        =   70
      Top             =   1920
      Width           =   1350
   End
End
Attribute VB_Name = "frmEditPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPackager
' Description   : To Add and Edit Principal Firms
'
' Change history
' Prog      Date        Description
' TW        17/10/2006  EP2_15 - Created
' TW        18/11/2006  EP2_132 ECR20/21
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        05/12/2006  EP2_276 - Deal with mandatory fields.
' TW        05/12/2006  EP2_258 - When Packager Unit created in Supervisor a Bucket user needs creating
' TW        06/12/2006  EP2_330 - Error linking a Packager to an a Association
' TW        07/12/2006  EP2_153 - Made Bank Account Number and Sort code mandatory
' TW        21/12/2006  EP2_623 - Fault when promoting automatically created users
' TW        09/01/2007  EP2_728 - Caching unit numbers and Ids
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        29/01/2007  EP2_1034 - Proc fee calculation processing to allow for firm-level basic rate adjustment
' TW        01/02/2007  EP2_1036 - Principal Firms/Network display and search (Follow-on)
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' TW        13/03/2007  EP2_1934 - DBM341 - Satellite Packagers - Allow firm name of up to 130 characters.
' TW        05/04/2007  EP2_2292 - Add functionality to create INTRODUCER and INTRODUCERFIRM entries
' TW        25/04/2007  EP2_2568 - Packagers to use sequence number 'PackagerFirmID' prefixed by 'P'
' TW        16/05/2007  VR86 - Set FirmPermissions.FRMPermissions = '4' ('Authorised') when adding permissions
' AW        02/10/2007  OMIGA00003231   Disable edit/delete when no row present
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Private data
Private m_clsPrincipalFirmTable As PrincipalFirmTable
Private m_clsActivityFSATable As ActivityFSATable
Private m_clsFirmPermissionsTable As FirmPermissionsTable
Private m_clsFirmClubNetAssocTable As FirmClubNetAssocTable
' TW 21/12/2006 EP2_623
Private m_clsUserTable As OmigaUserTable
Private m_clsUnitTable As UnitTable
' TW 21/12/2006 EP2_623 End
' TW 05/04/2007 EP2_2292
Private m_clsIntroducerTable As IntroducerTable
' TW 05/04/2007 EP2_2292 End

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
' TW 05/12/2006 EP2_258
Dim strUserID As String
' TW 05/12/2006 EP2_258 End
' TW 05/04/2007 EP2_2292
Dim strIntroducerID As String
' TW 05/04/2007 EP2_2292 End
' TW 25/04/2007 EP2_2568
Dim strHeading As String
' TW 25/04/2007 EP2_2568 End


Private Function CheckIfOKToSave() As Boolean
Dim bRet As Boolean
    On Error GoTo Failed
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
       bRet = g_clsValidation.ValidateSortCode(txtSortCode)
    End If
    If bRet Then
        SaveScreenData
    End If
    CheckIfOKToSave = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    CheckIfOKToSave = False
End Function

Private Sub CreateNewUser()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
' TW 05/04/2007 EP2_2292
Dim strSQL As String
Dim strListingStatus As String
Dim strEmailAddress As String
' TW 05/04/2007 EP2_2292 End

    On Error GoTo Failed
    
    strUserID = ""
    strUnitID = ""

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "usp_CreatePackagerUser"
' TW 13/03/2007 EP2_1934
'        .Parameters.Append .CreateParameter("UnitName", adVarChar, adParamInput, 45, txtPrincipalFirmName.Text)
'        .Parameters.Append .CreateParameter("UserSurname", adVarChar, adParamInput, 70, txtPrincipalFirmName.Text)
'        .Parameters.Append .CreateParameter("UserForename", adVarChar, adParamInput, 60, "")
        .Parameters.Append .CreateParameter("UnitName", adVarChar, adParamInput, 130, txtPrincipalFirmName.Text)
        .Parameters.Append .CreateParameter("UserSurname", adVarChar, adParamInput, 35, Left$(txtPrincipalFirmName.Text, 35))
        .Parameters.Append .CreateParameter("UserForename", adVarChar, adParamInput, 30, "")
' TW 13/03/2007 EP2_1934 End
        .Parameters.Append .CreateParameter("Initials", adVarChar, adParamInput, 3, "")
        .Parameters.Append .CreateParameter("Title", adInteger, adParamInput, 5, 99) 'Other
        .Parameters.Append .CreateParameter("DateOfBirth", adDate, adParamInput, , Null)
        .Parameters.Append .CreateParameter("UserId", adVarChar, adParamOutput, 24)
        .Parameters.Append .CreateParameter("UnitId", adVarChar, adParamOutput, 20)
        .Execute , , adExecuteNoRecords
        strUserID = .Parameters("UserId").Value
        strUnitID = .Parameters("UnitId").Value
    End With
    
    
' TW 05/04/2007 EP2_2292
' Create the Introducer record
    strIntroducerID = strUserID
    If cboListingStatus.ListIndex > -1 Then
        strListingStatus = cboListingStatus.GetExtra(cboListingStatus.ListIndex)
    Else
        strListingStatus = ""
    End If
    strEmailAddress = strIntroducerID & "@vertex.co.uk" ' invent an email address !!

    strSQL = "INSERT INTO INTRODUCER (INTRODUCERID, USERID, NATIONALINSURANCENUMBER, EMAILADDRESS, INTRODUCERTYPE," & _
                       "ARINDICATOR, LISTINGSTATUS, BUILDINGORHOUSENAME, BUILDINGORHOUSENUMBER, " & _
                       "FLATNUMBER, STREET, DISTRICT, TOWN, COUNTY, COUNTRY, POSTCODE, MARKETINGDATAOPTOUT, " & _
                       "LASTUPDATEDDATE, LASTUPDATEDBY) VALUES ("
    strSQL = strSQL & "'" & strIntroducerID & "', '" & strUserID & "', '', '" & strEmailAddress & "', 2, " & _
                      "0, " & strListingStatus & ", '" & Left$(txtAddressLine1.Text, 40) & "', '" & Left$(txtAddressLine2.Text, 10) & "', " & _
                      "'" & Left$(txtAddressLine3.Text, 10) & "', '" & txtAddressLine4.Text & "', '" & txtAddressLine5.Text & "', '" & txtAddressLine6.Text & "', '', 10, '" & txtPostCode.Text & "', 0, " & _
                      "GETDATE(), '" & strUserID & "')"
                      
    g_clsDataAccess.GetActiveConnection.Execute (strSQL)

' TW 05/04/2007 EP2_2292 End
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strUserID) = 0 Then
        MsgBox "Can't allocate a User ID for this record", vbCritical
    End If
    If Len(strUnitID) = 0 Then
        MsgBox "Can't allocate a Unit ID for this record", vbCritical
    Else
        txtUnitID.Text = strUnitID
    End If
    
' TW 21/12/2006 EP2_623
    If Len(strUserID) > 0 And Len(strUnitID) > 0 Then
        SaveChangeRequest strUserID, 4
        SaveChangeRequest strUnitID, 5
' TW 05/04/2007 EP2_2292
        SaveChangeRequest strIntroducerID, 6
' TW 05/04/2007 EP2_2292 End
    End If
' TW 21/12/2006 EP2_623 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function GetIntroducerFirmSeqNo() As String
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim strIntroducerFirmSeqNo As String
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "IntroducerFirmSeqNo")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strIntroducerFirmSeqNo = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    GetIntroducerFirmSeqNo = strIntroducerFirmSeqNo
    Exit Function
Failed:
End Function




Private Sub PopulateTradingNamesListView()
Dim m_clsFirmTradingNameTable As New FirmTradingNameTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strPackagerID
    
    m_clsFirmTradingNameTable.SetFindLinkedFirmTradingNameRecords strPackagerID
    lvTradingNames.LoadColumnDetails TypeName(m_clsFirmTradingNameTable)
    
    TableAccess(m_clsFirmTradingNameTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvTradingNames, m_clsFirmTradingNameTable
    
    cmdEditTradingName.Enabled = False
    cmdDeleteTradingName.Enabled = False
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateAvailableAssociationsListView()
Dim m_clsFirmClubNetAssocTable As New FirmClubNetAssocTable
Dim intPackagerIndicator As Integer
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strPackagerID

    If InStr(1, Me.Caption, "Packager") = 0 Then
        intPackagerIndicator = 0
    Else
        intPackagerIndicator = 1
    End If
    m_clsFirmClubNetAssocTable.SetFindAvailableFirmAssociationRecords strPackagerID, intPackagerIndicator
    lvAvailableAssociations.LoadColumnDetails TypeName(m_clsFirmClubNetAssocTable)
    
    TableAccess(m_clsFirmClubNetAssocTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvAvailableAssociations, m_clsFirmClubNetAssocTable
    
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateAssociationsListView()
Dim m_clsFirmClubNetAssocTable As New FirmClubNetAssocTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strPackagerID

    m_clsFirmClubNetAssocTable.SetFindLinkedFirmAssociationRecords strPackagerID
    lvCurrentAssociations.LoadColumnDetails TypeName(m_clsFirmClubNetAssocTable)
    
    TableAccess(m_clsFirmClubNetAssocTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvCurrentAssociations, m_clsFirmClubNetAssocTable
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupAvailableAssociationHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders

    lvHeaders.nWidth = 35
    lvHeaders.sName = "Firm Details"
    headers.Add lvHeaders

    lvHeaders.nWidth = 35
    lvHeaders.sName = "Address"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders

    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders

    lvAvailableAssociations.AddHeadings headers
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupAssociationHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 0
    lvHeaders.sName = "Sequence"
    headers.Add lvHeaders

    lvHeaders.nWidth = 35
    lvHeaders.sName = "Firm Details"
    headers.Add lvHeaders

    lvHeaders.nWidth = 35
    lvHeaders.sName = "Address"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Agreed Proc Fee Rate"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Vol Proc Fee Adjustment"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""
    headers.Add lvHeaders

    lvCurrentAssociations.AddHeadings headers
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

Private Sub cboFirmLegalStatus_Validate(Cancel As Boolean)
    Cancel = Not cboFirmLegalStatus.ValidateData()
End Sub


Private Sub cboFirmStatus_Validate(Cancel As Boolean)
    Cancel = Not cboFirmStatus.ValidateData()
End Sub


Private Sub cboFirmType_Validate(Cancel As Boolean)
    Cancel = Not cboFirmType.ValidateData()
End Sub


Private Sub cboListingStatus_Validate(Cancel As Boolean)
    Cancel = Not cboListingStatus.ValidateData()
End Sub



Private Sub cmdAddActivity_Click()
Dim X As Integer
Dim clsTableAccess As TableAccess
    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        For X = 1 To lvAvailableActivities.ListItems.Count
            If lvAvailableActivities.ListItems.Item(X).Selected Then
    
                strActivityID = lvAvailableActivities.ListItems.Item(X)
                Call GetFirmPermissionsSeqNo
                Set m_colKeys = New Collection
                       
                m_colKeys.Add strFirmPermissionsSeqNo
                
                'Create empty records.
                g_clsFormProcessing.CreateNewRecord m_clsFirmPermissionsTable
                Set clsTableAccess = m_clsFirmPermissionsTable
                m_clsFirmPermissionsTable.SetFirmPermissionSeqNo strFirmPermissionsSeqNo
                m_clsFirmPermissionsTable.SetPrincipalfirmID strPackagerID
                m_clsFirmPermissionsTable.SetARFirmID strARFirmID
                m_clsFirmPermissionsTable.SetActivityID strActivityID
' TW 16/05/2007 VR86
                m_clsFirmPermissionsTable.SetFRMPermissions "4"
' TW 16/05/2007 VR86 End
                
                ' Other fields ??
                m_clsFirmPermissionsTable.SetLastupdateddate Now
                  
                clsTableAccess.Update
                SaveChangeRequest strFirmPermissionsSeqNo, 2
                
            End If
        Next X
    
        Call PopulateAvailableActivitiesListView
        Call PopulateCurrentActivitiesListView
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub cmdAddAssociation_Click()
Dim X As Integer
Dim clsTableAccess As TableAccess

    On Error GoTo Failed:
    If CheckIfOKToSave() Then
        For X = 1 To lvAvailableAssociations.ListItems.Count
            If lvAvailableAssociations.ListItems.Item(X).Selected Then
    
                Call GetFirmClubSeqNo
                Set m_colKeys = New Collection
                       
                m_colKeys.Add strFirmClubSeqNo
                
                g_clsFormProcessing.CreateNewRecord m_clsFirmClubNetAssocTable
                Set clsTableAccess = m_clsFirmClubNetAssocTable
                m_clsFirmClubNetAssocTable.SetFirmClubSeqNo strFirmClubSeqNo
                m_clsFirmClubNetAssocTable.SetPrincipalfirmID strPackagerID
                m_clsFirmClubNetAssocTable.SetARFirmID strARFirmID
                m_clsFirmClubNetAssocTable.SetClubNetworkAssociationId lvAvailableAssociations.ListItems.Item(X).SubItems(5)
                  
                clsTableAccess.Update
                SaveChangeRequest strFirmClubSeqNo, 3
                
            End If
        Next X
    
    
        Call PopulateAvailableAssociationsListView
        Call PopulateAssociationsListView
    End If
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError

End Sub


Private Sub cmdBankWizard_Click()
    
Dim bRet As Boolean
Dim sSortCode As String
Dim sBankAccountNumber As String
Dim xdcBankWizardDoc As FreeThreadedDOMDocument
Dim strXMLData As String
Dim strError As String
Dim xndDescription As IXMLDOMNode
Dim iCount As Integer

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

Private Sub cmdFindExclusives_Click()
' TW 23/11/2006 EP2_172
Dim colKeys As New Collection
    On Error GoTo Failed:
    colKeys.Add strPackagerID
    colKeys.Add "Introducer"
    colKeys.Add frmMain.tvwDB.SelectedItem
    frmMaintainProductExclusivity.SetKeys colKeys
    frmMaintainProductExclusivity.SetIsEdit True
    frmMaintainProductExclusivity.Show 1
    
    Unload frmMaintainProductExclusivity
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
' TW 23/11/2006 EP2_172 End
End Sub

Private Sub cmdRemoveAssociation_Click()
Dim X As Integer
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection
Dim m_clsFirmClubNetAssocTable As New FirmClubNetAssocTable

    On Error GoTo Failed:
    Set clsTableAccess = m_clsFirmClubNetAssocTable
    For X = 1 To lvCurrentAssociations.ListItems.Count
        If lvCurrentAssociations.ListItems.Item(X).Selected Then
    
            Set colMatchValues = New Collection
            colMatchValues.Add lvCurrentAssociations.ListItems.Item(X)

            clsTableAccess.DeleteRow colMatchValues
            clsTableAccess.Update
            
' TW 06/12/2006 EP2_330
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsFirmClubNetAssocTable, , PromoteDelete
' TW 06/12/2006 EP2_330 End
        End If
    Next X

    Call PopulateAvailableAssociationsListView
    Call PopulateAssociationsListView
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvAvailableActivities_Click()
    cmdAddActivity.Enabled = True
    cmdRemoveActivity.Enabled = False
End Sub

Private Sub lvAvailableAssociations_Click()
    cmdAddAssociation.Enabled = True
    cmdRemoveAssociation.Enabled = False
End Sub

Private Sub lvCategories_Click()
    Call PopulateAvailableActivitiesListView

End Sub

Private Sub lvCurrentActivities_Click()
    cmdAddActivity.Enabled = False
    cmdRemoveActivity.Enabled = True
End Sub

Private Sub lvCurrentActivities_DeletePressed()
    Call cmdRemoveActivity_Click
End Sub


Private Sub cmdAddTradingName_Click()
    If CheckIfOKToSave() Then
        frmEditPackagerTradingName.SetIsEdit (False)
        frmEditPackagerTradingName.txtPrincipalFirmID.Text = strPackagerID
        frmEditPackagerTradingName.txtARFirmID.Text = ""
        frmEditPackagerTradingName.Show 1
        If frmEditPackagerTradingName.GetReturnCode() Then
            PopulateTradingNamesListView
        End If
        Unload frmEditPackagerTradingName
    End If
End Sub

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

Private Sub cmdEditTradingName_Click()
Dim colMatchValues As Collection
    
    Set colMatchValues = New Collection
    colMatchValues.Add lvTradingNames.getValueFromName("Sequence")
    
    frmEditPackagerTradingName.SetKeys colMatchValues
    frmEditPackagerTradingName.SetIsEdit (True)

    frmEditPackagerTradingName.Show 1
    If frmEditPackagerTradingName.GetReturnCode() Then
        PopulateTradingNamesListView
    End If
    Unload frmEditPackagerTradingName
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    strHeading = Replace(frmMain.tvwDB.SelectedItem, "s", "")
    Me.Caption = "Add/Edit " & strHeading
    
    Select Case strHeading
        Case "Packager"
' TW 16/05/2007 VR86
            SSTab1.TabVisible(3) = False
' TW 16/05/2007 VR86 End
' TW 01/02/2007 EP2_1036
            FSAReferencePanel.Visible = False
' TW 01/02/2007 EP2_1036 End
            optPackagerIndicator(0) = True
' TW 05/12/2006 EP2_258
'        Case "Principal Firm"
        Case "Principal Firm/Network"
' TW 05/12/2006 EP2_258 End
' TW 16/05/2007 VR86
            SSTab1.TabVisible(3) = True
' TW 16/05/2007 VR86 End
' TW 01/02/2007 EP2_1036
            FSAReferencePanel.Visible = True
' TW 01/02/2007 EP2_1036 End
            optPackagerIndicator(1) = True
            SSTab1.TabCaption(5) = "Mortgage Clubs"
            Frame5.Caption = "Linked Mortgage Clubs"
            Frame6.Caption = "Available Mortgage Clubs"
    End Select
    
' TW 25/04/2007 EP2_2568
    Label2.Caption = strHeading & " Id"
' TW 25/04/2007 EP2_2568 End
    
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New PrincipalFirmTable
    Set m_clsFirmPermissionsTable = New FirmPermissionsTable
    Set m_clsPrincipalFirmTable = New PrincipalFirmTable
    Set m_clsFirmClubNetAssocTable = New FirmClubNetAssocTable
' TW 21/12/2006 EP2_623
    Set m_clsUserTable = New OmigaUserTable
    Set m_clsUnitTable = New UnitTable
' TW 21/12/2006 EP2_623 End
' TW 05/04/2007 EP2_2292
    Set m_clsIntroducerTable = New IntroducerTable
' TW 05/04/2007 EP2_2292 End
    
    g_clsFormProcessing.PopulateCombo "FirmStatus", Me.cboFirmStatus
    g_clsFormProcessing.PopulateCombo "FirmLegalStatus", Me.cboFirmLegalStatus
    g_clsFormProcessing.PopulateCombo "ListingStatus", Me.cboListingStatus
    g_clsFormProcessing.PopulateCombo "FirmType", Me.cboFirmType
    
    SetupTradingNameHeaders
    SetupAssociationHeaders
    
    SetupCategoryHeaders
    PopulateCategories
    
    SetupAvailableAssociationHeaders
    
    SetupCurrentActivityHeaders
    
    SetupAvailableActivityHeaders
    
' TW 23/11/2006 EP2_172
    cmdFindExclusives.Visible = m_bIsEdit
' TW 23/11/2006 EP2_172 End
    
    If m_bIsEdit = True Then
        SetEditState
    Else
'        SSTab1.TabEnabled(3) = False
'        SSTab1.TabEnabled(4) = False
'        SSTab1.TabEnabled(5) = False
        SetAddState
    End If
    PopulateAvailableAssociationsListView
    
' TW 05/12/2006 EP2_258
'    cmdCreateAsOmigaUnit.Enabled = (Len(txtUnitID.Text) = 0
    cmdCreateAsOmigaUnit.Enabled = (Len(txtUnitID.Text) = 0 And optPackagerIndicator(1).Value = True)
' TW 05/12/2006 EP2_258 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupCategoryHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Category"
    headers.Add lvHeaders
    lvCategories.AddHeadings headers

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateCategories()
Dim m_clsFirmActivityTable As New FirmActivityTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    lvCategories.ListItems.Clear
    
    Set colMatchValues = New Collection
    
    m_clsFirmActivityTable.SetFindDistinctCategories
    lvCategories.LoadColumnDetails TypeName(m_clsFirmActivityTable)
    
    TableAccess(m_clsFirmActivityTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvCategories, m_clsFirmActivityTable
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetFirmClubSeqNo()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "FirmClubSeqNo")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strFirmClubSeqNo = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strFirmClubSeqNo) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub

Private Sub GetFirmPermissionsSeqNo()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "FirmPermissionsSeqNo")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strFirmPermissionsSeqNo = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strFirmPermissionsSeqNo) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub


Private Sub cmdRemoveActivity_Click()
    
Dim X As Integer
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection

    On Error GoTo Failed:
    Set clsTableAccess = m_clsFirmPermissionsTable
    For X = 1 To lvCurrentActivities.ListItems.Count
        If lvCurrentActivities.ListItems.Item(X).Selected Then
    
            Set colMatchValues = New Collection
            colMatchValues.Add lvCurrentActivities.ListItems.Item(X)

            clsTableAccess.DeleteRow colMatchValues
            clsTableAccess.Update

' TW 06/12/2006 EP2_330
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsFirmPermissionsTable, , PromoteDelete
' TW 06/12/2006 EP2_330 End
        End If
    Next X


    Call PopulateAvailableActivitiesListView
    Call PopulateCurrentActivitiesListView
'    SetReturnCode
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub PopulateAvailableActivitiesListView()
Dim m_clsActivityFSATable As New ActivityFSATable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    lvAvailableActivities.ListItems.Clear
    
    Set colMatchValues = New Collection
    colMatchValues.Add strPackagerID
    
    m_clsActivityFSATable.SetFindPackagerAvailableActivitiesByCategory strPackagerID, lvCategories.SelectedItem
    lvAvailableActivities.LoadColumnDetails TypeName(m_clsActivityFSATable)
    
    TableAccess(m_clsActivityFSATable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvAvailableActivities, m_clsActivityFSATable
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateCurrentActivitiesListView()
Dim m_clsFirmPermissionsTable As New FirmPermissionsTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    
    lvCurrentActivities.ListItems.Clear
    
    Set colMatchValues = New Collection
    colMatchValues.Add strPackagerID
    
    m_clsFirmPermissionsTable.SetFindFirmLinkedActivityFSARecords strPackagerID
    lvCurrentActivities.LoadColumnDetails TypeName(m_clsFirmPermissionsTable)
    
    TableAccess(m_clsFirmPermissionsTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditPackager.lvCurrentActivities, m_clsFirmPermissionsTable

 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub SetupCurrentActivityHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 0
    lvHeaders.sName = ""    'FirmPermissionsSeqNo
    headers.Add lvHeaders

    lvHeaders.nWidth = 0
    lvHeaders.sName = ""    'PrincipalFirmID
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = ""    'ARFirmID
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Activity ID"
    headers.Add lvHeaders

    lvHeaders.nWidth = 40
    lvHeaders.sName = "Category"
    headers.Add lvHeaders

    lvHeaders.nWidth = 40
    lvHeaders.sName = "Description"
    headers.Add lvHeaders

    lvCurrentActivities.AddHeadings headers

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetupAvailableActivityHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Activity ID"
    headers.Add lvHeaders

    lvHeaders.nWidth = 40
    lvHeaders.sName = "Description"
    headers.Add lvHeaders

    lvHeaders.nWidth = 0
    lvHeaders.sName = ""    'was Category
    headers.Add lvHeaders

    lvAvailableActivities.AddHeadings headers

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed

' TW 13/03/2007 EP2_1934
    If Len(txtPrincipalFirmName.Text) = 0 Then
        txtPrincipalFirmName.BackColor = vbRed
        g_clsErrorHandling.RaiseError errMandatoryFieldsRequired
        txtPrincipalFirmName.SetFocus
        Exit Function
    End If
' TW 13/03/2007 EP2_1934 End

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet Then
       bRet = g_clsValidation.ValidateSortCode(txtSortCode)
    End If
    If bRet = True Then
        SaveScreenData
        SaveChangeRequest strPackagerID, 1
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
            Set clsTableAccess = m_clsPrincipalFirmTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsPrincipalFirmTable
        Case 2
            Set clsTableAccess = m_clsFirmPermissionsTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsFirmPermissionsTable
        Case 3
            Set clsTableAccess = m_clsFirmClubNetAssocTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsFirmClubNetAssocTable
' TW 21/12/2006 EP2_623
        Case 4
            Set clsTableAccess = m_clsUserTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsUserTable
        Case 5
            Set clsTableAccess = m_clsUnitTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsUnitTable
' TW 21/12/2006 EP2_623 End
' TW 05/04/2007 EP2_2292
        Case 6
            Set clsTableAccess = m_clsIntroducerTable
            clsTableAccess.SetKeyMatchValues colMatchValues
            g_clsHandleUpdates.SaveChangeRequest m_clsIntroducerTable
' TW 05/04/2007 EP2_2292 End
    
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
Dim clsTableAccess As TableAccess
' TW 05/04/2007 EP2_2292
Dim strSQL As String
' TW 05/04/2007 EP2_2292 End
    
    On Error GoTo Failed
    
    ' TW 05/12/2006 EP2_258
    If Not m_bIsEdit And optPackagerIndicator(0).Value = True And Len(txtUnitID.Text) = 0 Then
        CreateNewUser
    End If
    ' TW 05/12/2006 EP2_258 End

    Set clsTableAccess = m_clsPrincipalFirmTable
    
    m_clsPrincipalFirmTable.SetPrincipalfirmID strPackagerID
    'Firm details tab
    m_clsPrincipalFirmTable.SetUnitID txtUnitID.Text
    m_clsPrincipalFirmTable.SetPrincipalFirmName txtPrincipalFirmName.Text
    If cboFirmStatus.ListIndex > -1 Then
        m_clsPrincipalFirmTable.SetFirmStatus cboFirmStatus.GetExtra(cboFirmStatus.ListIndex)
    End If
    If optPackagerIndicator(0).Value = True Then
        m_clsPrincipalFirmTable.SetPackagerIndicator 1
    Else
        m_clsPrincipalFirmTable.SetPackagerIndicator 0
    End If
    m_clsPrincipalFirmTable.SetFirmLegalStatusType cboFirmLegalStatus.GetExtra(cboFirmLegalStatus.ListIndex)
    If cboFirmType.ListIndex > -1 Then
        m_clsPrincipalFirmTable.SetFirmTypeCode cboFirmType.GetExtra(cboFirmType.ListIndex)
    End If
    If cboListingStatus.ListIndex > -1 Then
        m_clsPrincipalFirmTable.SetListingStatus cboListingStatus.GetExtra(cboListingStatus.ListIndex)
    End If
    If optIncludeLoading(0).Value = True Then
        m_clsPrincipalFirmTable.SetProcLoadingInd 1
    Else
        m_clsPrincipalFirmTable.SetProcLoadingInd 0
    End If
    If optIncludeOnlineSubmissionLoading(0).Value = True Then
        m_clsPrincipalFirmTable.SetProcFeeOnlineInd 1
    Else
        m_clsPrincipalFirmTable.SetProcFeeOnlineInd 0
    End If

' TW 29/01/2007 EP2_1034
    m_clsPrincipalFirmTable.SetAgreedProcFeeRate Val(Format$(txtAgreedProcFeeRate.Text))
' TW 29/01/2007 EP2_1034 End
    
    'Contact Details tab
    
    m_clsPrincipalFirmTable.SetAddressLine1 txtAddressLine1.Text
    m_clsPrincipalFirmTable.SetAddressLine2 txtAddressLine2.Text
    m_clsPrincipalFirmTable.SetAddressLine3 txtAddressLine3.Text
    m_clsPrincipalFirmTable.SetAddressLine4 txtAddressLine4.Text
    m_clsPrincipalFirmTable.SetAddressLine5 txtAddressLine5.Text
    m_clsPrincipalFirmTable.SetAddressLine6 txtAddressLine6.Text
    m_clsPrincipalFirmTable.SetPostcode txtPostCode.Text
    m_clsPrincipalFirmTable.SetTelephoneCountryCode txtCountryCode1.Text
    m_clsPrincipalFirmTable.SetTelephoneAreaCode txtAreaCode1.Text
    m_clsPrincipalFirmTable.SetTelephoneNumber txtTelephoneNumber1.Text
    m_clsPrincipalFirmTable.SetFaxCountryCode txtCountryCode2.Text
    m_clsPrincipalFirmTable.SetFaxAreaCode txtAreaCode2.Text
    m_clsPrincipalFirmTable.SetFaxNumber txtTelephoneNumber2.Text
      
    ' Bank Details tab
    m_clsPrincipalFirmTable.SetBankAccountName txtAccountName.Text
    m_clsPrincipalFirmTable.SetBankAccountNumber txtAccountNumber.Text
    m_clsPrincipalFirmTable.SetBankAccountBranchName txtBranchName.Text
    m_clsPrincipalFirmTable.SetBankSortCode txtSortCode.Text
    m_clsPrincipalFirmTable.SetBankWizardIndicator chkBankWizardIndicator.Value
    
    m_clsPrincipalFirmTable.SetLastupdateddate Now
    m_clsPrincipalFirmTable.SetLastUpdatedBy g_sSupervisorUser
      
    clsTableAccess.Update
    
' TW 05/04/2007 EP2_2292
    If Len(strIntroducerID) > 0 And optPackagerIndicator(0).Value = True Then
' Create the IntroducerFirm record
        strSQL = "INSERT INTO INTRODUCERFIRM (INTRODUCERFIRMSEQNO, INTRODUCERID, PRINCIPALFIRMID) VALUES ("
        strSQL = strSQL & "'" & GetIntroducerFirmSeqNo() & "', '" & strIntroducerID & "', '" & strPackagerID & "')"
        g_clsDataAccess.GetActiveConnection.Execute (strSQL)
    End If
    strIntroducerID = ""
' TW 05/04/2007 EP2_2292 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new PrincipalFirm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed
    
    Call GetNewPackagerID
' TW 25/04/2007 EP2_2568
    txtId.Text = strPackagerID
' TW 25/04/2007 EP2_2568 End
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add strPackagerID
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsPrincipalFirmTable
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetNewPackagerID()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
' TW 25/04/2007 EP2_2568
Dim strSequenceName As String
' TW 25/04/2007 EP2_2568 End

    On Error GoTo Failed

' TW 25/04/2007 EP2_2568
    Select Case strHeading
        Case "Packager"
            strSequenceName = "PackagerFirmID"
        Case "Principal Firm/Network"
            strSequenceName = "PrincipalFirmID"
    End Select
' TW 25/04/2007 EP2_2568 End

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
' TW 25/04/2007 EP2_2568
'        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "PrincipalFirmID")
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, strSequenceName)
' TW 25/04/2007 EP2_2568 End
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strPackagerID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strPackagerID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
' TW 25/04/2007 EP2_2568
    Else
        If strHeading = "Packager" Then
            strPackagerID = "P" & Format$(Val(strPackagerID), "0000")
        End If
' TW 25/04/2007 EP2_2568 End
    End If
Failed:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing PrincipalFirm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
Dim colDataSet As New Collection

    strPackagerID = m_colKeys(1)
' TW 25/04/2007 EP2_2568
    txtId.Text = strPackagerID
' TW 25/04/2007 EP2_2568 End

    PopulateAssociationsListView
    PopulateTradingNamesListView

    PopulateCurrentActivitiesListView

    colDataSet.Add m_colKeys.Item(1)
    TableAccess(m_clsPrincipalFirmTable).SetKeyMatchValues colDataSet
    TableAccess(m_clsPrincipalFirmTable).GetTableData
      
    'Firm details tab
' TW 16/01/2007 EP2_859
    txtFSARef.Text = m_clsPrincipalFirmTable.GetFSARef
' TW 16/01/2007 EP2_859 End
    txtUnitID.Text = m_clsPrincipalFirmTable.GetUnitId
    txtPrincipalFirmName.Text = m_clsPrincipalFirmTable.GetPrincipalFirmName()
    g_clsFormProcessing.HandleComboExtra cboFirmStatus, m_clsPrincipalFirmTable.GetFirmStatus, SET_CONTROL_VALUE
    
    If Val(m_clsPrincipalFirmTable.GetPackagerIndicator()) = 1 Then
        optPackagerIndicator(0) = True
    Else
        optPackagerIndicator(1) = True
    End If
    g_clsFormProcessing.HandleComboExtra cboFirmLegalStatus, m_clsPrincipalFirmTable.GetFirmLegalStatusType, SET_CONTROL_VALUE
    
    g_clsFormProcessing.HandleComboExtra cboFirmType, m_clsPrincipalFirmTable.GetFirmTypeCode, SET_CONTROL_VALUE
    
    g_clsFormProcessing.HandleComboExtra cboListingStatus, m_clsPrincipalFirmTable.GetListingStatus, SET_CONTROL_VALUE
    If Val(m_clsPrincipalFirmTable.GetProcLoadingInd) = 1 Then
        optIncludeLoading(0) = True
    Else
        optIncludeLoading(1) = True
    End If

' TW 29/01/2007 EP2_1034
    txtAgreedProcFeeRate.Text = m_clsPrincipalFirmTable.GetAgreedProcFeeRate
' TW 29/01/2007 EP2_1034 End

    If Val(m_clsPrincipalFirmTable.GetProcFeeOnlineInd) = 1 Then
        optIncludeOnlineSubmissionLoading(0) = True
    Else
        optIncludeOnlineSubmissionLoading(1) = True
    End If
      
    'Contact Details tab
    txtAddressLine1.Text = m_clsPrincipalFirmTable.GetAddressLine1()
    txtAddressLine2.Text = m_clsPrincipalFirmTable.GetAddressLine2()
    txtAddressLine3.Text = m_clsPrincipalFirmTable.GetAddressLine3()
    txtAddressLine4.Text = m_clsPrincipalFirmTable.GetAddressLine4()
    txtAddressLine5.Text = m_clsPrincipalFirmTable.GetAddressLine5()
    txtAddressLine6.Text = m_clsPrincipalFirmTable.GetAddressLine6()
    txtPostCode.Text = m_clsPrincipalFirmTable.GetPostcode()
    
    txtCountryCode1.Text = m_clsPrincipalFirmTable.GetTelephoneCountryCode()
    txtAreaCode1.Text = m_clsPrincipalFirmTable.GetTelephoneAreaCode()
    txtTelephoneNumber1.Text = m_clsPrincipalFirmTable.GetTelephoneNumber()
    txtCountryCode2.Text = m_clsPrincipalFirmTable.GetFaxCountryCode()
    txtAreaCode2.Text = m_clsPrincipalFirmTable.GetFaxAreaCode()
    txtTelephoneNumber2.Text = m_clsPrincipalFirmTable.GetFaxNumber
      
    ' Bank Details tab
    txtAccountName.Text = m_clsPrincipalFirmTable.GetBankAccountName()
    txtAccountNumber.Text = m_clsPrincipalFirmTable.GetBankAccountNumber()
    txtBranchName.Text = m_clsPrincipalFirmTable.GetBankAccountBranchName()
    txtSortCode.Text = m_clsPrincipalFirmTable.GetBankSortCode()
    chkBankWizardIndicator.Value = Val(m_clsPrincipalFirmTable.GetBankWizardIndicator())
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetPrincipalFirmTableClass(clsPrincipalFirmTable As PrincipalFirmTable)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetPrincipalFirmTableClass
' Description   : Sets the Processing class to be used by this form for all events - i.e., this
'                 form will call into m_clsPrincipalFirmTable for all business processing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo Failed
    
    If clsPrincipalFirmTable Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "PrincipalFirmTable table class emtpy"
    End If
    
    Set m_clsPrincipalFirmTable = clsPrincipalFirmTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, TypeName(Me) & ".SetPrincipalFirmTable " & Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
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
    Set m_clsPrincipalFirmTable = Nothing
End Sub


Private Sub lvCurrentAssociations_Click()
    cmdAddAssociation.Enabled = False
    cmdRemoveAssociation.Enabled = True
End Sub

Private Sub lvTradingNames_Click()
    ' AW 02/10/07  OMIGA00003231
    If frmEditPackager.lvTradingNames.ListItems.Count > 0 Then
        cmdEditTradingName.Enabled = True
        cmdDeleteTradingName.Enabled = True
    Else
        cmdEditTradingName.Enabled = False
        cmdDeleteTradingName.Enabled = False
    End If
End Sub

Private Sub lvTradingNames_DblClick()
    ' AW 02/10/07  OMIGA00003231
    If frmEditPackager.lvTradingNames.ListItems.Count > 0 Then
        Call cmdEditTradingName_Click
    End If
End Sub


Private Sub lvTradingNames_DeletePressed()
    Call cmdDeleteTradingName_Click
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



Private Sub txtBranchName_Validate(Cancel As Boolean)
    Cancel = Not txtBranchName.ValidateData()
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



' TW 13/03/2007 EP2_1934
Private Sub txtPrincipalFirmName_LostFocus()
    txtPrincipalFirmName.BackColor = vbWhite
End Sub
' TW 13/03/2007 EP2_1934 End


Private Sub txtSortCode_Validate(Cancel As Boolean)
    Cancel = Not txtSortCode.ValidateData()
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


