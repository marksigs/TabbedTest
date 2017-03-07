VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIndividualPackager 
   Caption         =   "Add/Edit Individual Packager"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form11"
   ScaleHeight     =   7200
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6600
      TabIndex        =   45
      Top             =   6720
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7980
      TabIndex        =   46
      Top             =   6720
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   47
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Introducer Details"
      TabPicture(0)   =   "frmEditIndividualPackager.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Principals"
      TabPicture(1)   =   "frmEditIndividualPackager.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "User Details"
      TabPicture(2)   =   "frmEditIndividualPackager.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "UserDetailsFrame"
      Tab(2).Control(1)=   "TelephonerNumberDetailsFrame"
      Tab(2).ControlCount=   2
      Begin VB.Frame TelephonerNumberDetailsFrame 
         Caption         =   "Telephone Numbers"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   77
         Top             =   3000
         Width           =   8655
         Begin MSGOCX.MSGComboBox cboTelephoneNumberType 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSGOCX.MSGEditBox txtCountryCode 
            Height          =   315
            Index           =   0
            Left            =   2280
            TabIndex        =   31
            Top             =   720
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
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
            MaxLength       =   3
         End
         Begin MSGOCX.MSGEditBox txtCountryCode 
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   36
            Top             =   1200
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
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
            MaxLength       =   3
         End
         Begin MSGOCX.MSGEditBox txtCountryCode 
            Height          =   315
            Index           =   2
            Left            =   2280
            TabIndex        =   41
            Top             =   1680
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
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
            MaxLength       =   3
         End
         Begin MSGOCX.MSGEditBox txtAreaCode 
            Height          =   315
            Index           =   0
            Left            =   3480
            TabIndex        =   32
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGEditBox txtAreaCode 
            Height          =   315
            Index           =   1
            Left            =   3480
            TabIndex        =   37
            Top             =   1200
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGEditBox txtAreaCode 
            Height          =   315
            Index           =   2
            Left            =   3480
            TabIndex        =   42
            Top             =   1680
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGEditBox txtTelephoneNumber 
            Height          =   315
            Index           =   0
            Left            =   4800
            TabIndex        =   33
            Top             =   720
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
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
            MaxLength       =   30
         End
         Begin MSGOCX.MSGEditBox txtTelephoneNumber 
            Height          =   315
            Index           =   1
            Left            =   4800
            TabIndex        =   38
            Top             =   1200
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
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
            MaxLength       =   30
         End
         Begin MSGOCX.MSGEditBox txtTelephoneNumber 
            Height          =   315
            Index           =   2
            Left            =   4800
            TabIndex        =   43
            Top             =   1680
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
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
            MaxLength       =   30
         End
         Begin MSGOCX.MSGEditBox txtExtensionNumber 
            Height          =   315
            Index           =   0
            Left            =   6840
            TabIndex        =   34
            Top             =   720
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtExtensionNumber 
            Height          =   315
            Index           =   1
            Left            =   6840
            TabIndex        =   39
            Top             =   1200
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtExtensionNumber 
            Height          =   315
            Index           =   2
            Left            =   6840
            TabIndex        =   44
            Top             =   1680
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGComboBox cboTelephoneNumberType 
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSGOCX.MSGComboBox cboTelephoneNumberType 
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   40
            Top             =   1680
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Country Code"
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   82
            Top             =   480
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Index           =   15
            Left            =   3480
            TabIndex        =   81
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Telephone Number"
            Height          =   195
            Index           =   16
            Left            =   4800
            TabIndex        =   80
            Top             =   480
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ext  No."
            Height          =   195
            Index           =   24
            Left            =   6840
            TabIndex        =   79
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Type"
            Height          =   195
            Index           =   27
            Left            =   240
            TabIndex        =   78
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Principals"
         Height          =   5775
         Left            =   -74760
         TabIndex        =   76
         Tag             =   "Tab3"
         Top             =   480
         Width           =   8655
         Begin VB.CommandButton cmdAddFirm 
            Caption         =   "&Modify"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Add or Delete Associated Principal Firms"
            Top             =   5160
            Width           =   1275
         End
         Begin MSGOCX.MSGListView lvFirms 
            Height          =   4695
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8281
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame UserDetailsFrame 
         Caption         =   "User Details"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   69
         Tag             =   "Tab2"
         Top             =   480
         Width           =   8655
         Begin MSGOCX.MSGComboBox cboUserTitle 
            Height          =   315
            Left            =   2280
            TabIndex        =   25
            Top             =   720
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
            Mandatory       =   -1  'True
            Text            =   ""
         End
         Begin MSGOCX.MSGEditBox txtUserInitials 
            Height          =   285
            Left            =   6360
            TabIndex        =   26
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSGOCX.MSGEditBox txtUserSurname 
            Height          =   285
            Left            =   2280
            TabIndex        =   28
            Top             =   1440
            Width           =   6015
            _ExtentX        =   10610
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
         Begin MSGOCX.MSGEditBox txtUserFirstForename 
            Height          =   285
            Left            =   2280
            TabIndex        =   27
            Top             =   1080
            Width           =   6015
            _ExtentX        =   10610
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
         Begin MSGOCX.MSGEditBox txtUserDateOfBirth 
            Height          =   285
            Left            =   2280
            TabIndex        =   29
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            Mask            =   "##/##/####"
            TextType        =   1
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
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
         Begin MSGOCX.MSGEditBox txtUserID 
            Height          =   285
            Left            =   2280
            TabIndex        =   24
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
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
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   75
            Top             =   1800
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "First Forename"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   74
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Title"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Initials"
            Height          =   195
            Index           =   20
            Left            =   5640
            TabIndex        =   72
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Surname"
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   71
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "User ID"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Introducer Address"
         Height          =   3015
         Left            =   240
         TabIndex        =   59
         Tag             =   "Tab1"
         Top             =   3240
         Width           =   8655
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   6360
            TabIndex        =   21
            Top             =   2520
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSGOCX.MSGEditBox txtPostCode 
            Height          =   285
            Left            =   2280
            TabIndex        =   20
            Top             =   2520
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   2280
            TabIndex        =   15
            Top             =   720
            Width           =   6015
            _ExtentX        =   10610
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
            Left            =   6360
            TabIndex        =   14
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   2280
            TabIndex        =   16
            Top             =   1080
            Width           =   6015
            _ExtentX        =   10610
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
            Left            =   2280
            TabIndex        =   17
            Top             =   1440
            Width           =   6015
            _ExtentX        =   10610
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
            Left            =   2280
            TabIndex        =   18
            Top             =   1800
            Width           =   6015
            _ExtentX        =   10610
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
            Left            =   2280
            TabIndex        =   19
            Top             =   2160
            Width           =   6015
            _ExtentX        =   10610
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
            Left            =   2280
            TabIndex        =   13
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Caption         =   "Building No."
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Postcode"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   67
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Building Name"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   66
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Street"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   65
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "District"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   64
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Town"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   63
            Top             =   1800
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "County"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   62
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Country"
            Height          =   195
            Index           =   8
            Left            =   4680
            TabIndex        =   61
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Flat No."
            Height          =   195
            Index           =   10
            Left            =   4680
            TabIndex        =   60
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Introducer Details"
         Height          =   2655
         Left            =   240
         TabIndex        =   48
         Tag             =   "Tab1"
         Top             =   480
         Width           =   8655
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   3495
            TabIndex        =   83
            Top             =   1800
            Width           =   3495
            Begin VB.OptionButton optNoCrossSellAgreement 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   11
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optNoCrossSellAgreement 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   10
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No Cross-Sell Agreement ?"
               Height          =   195
               Index           =   28
               Left            =   0
               TabIndex        =   84
               Top             =   0
               Width           =   1890
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   3495
            TabIndex        =   49
            Top             =   1080
            Width           =   3495
            Begin VB.OptionButton optARIndicator 
               Caption         =   "No"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   2880
               TabIndex        =   5
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
            Begin VB.OptionButton optARIndicator 
               Caption         =   "Yes"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   4
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Appointed Rep ?"
               Height          =   195
               Index           =   13
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Width           =   1200
            End
         End
         Begin VB.OptionButton optMarketingOptOut 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   8
            Top             =   1440
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optMarketingOptOut 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   7
            Top             =   1440
            Width           =   615
         End
         Begin MSGOCX.MSGEditBox txtEmailAddress 
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Top             =   2160
            Width           =   6015
            _ExtentX        =   10610
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
         Begin MSGOCX.MSGComboBox cboIntroducerType 
            Height          =   315
            Left            =   6360
            TabIndex        =   1
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
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
            Left            =   6360
            TabIndex        =   6
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
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
         Begin MSGOCX.MSGEditBox txtNationalInsuranceNumber 
            Height          =   285
            Left            =   2280
            TabIndex        =   2
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
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
         End
         Begin MSGOCX.MSGComboBox cboIntroducerStatus 
            Height          =   315
            Left            =   6360
            TabIndex        =   3
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
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
         Begin VB.Label lblFailedLoginAttempts 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6360
            TabIndex        =   9
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Failed Login Attempts"
            Height          =   195
            Index           =   35
            Left            =   4680
            TabIndex        =   58
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Marketing Opt Out ?"
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   57
            Top             =   1440
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Introducer Status"
            Height          =   195
            Index           =   14
            Left            =   4680
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "National Insurance No"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   55
            Top             =   720
            Width           =   1590
         End
         Begin VB.Label lblIntroducerID 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   0
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Introducer ID"
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Introducer Type"
            Height          =   195
            Index           =   23
            Left            =   4680
            TabIndex        =   53
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Listing Status"
            Height          =   195
            Index           =   25
            Left            =   4680
            TabIndex        =   52
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email Address"
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   51
            Top             =   2160
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmEditIndividualPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditIndividualPackager
' Description   : To Add and Edit DDD
'
' Change history
' Prog      Date        Description
' TW        17/10/2006  EP2_15 - Created
' TW        18/11/2006  EP2_132 ECR20/21 Added 'Updated by' and 'Date last updated'
' TW        04/12/2006  EP2_189 - Deal with mandatory fields.
' TW        05/12/2006  EP2_258 - Deal with mandatory fields.
' TW        06/12/2006  EP2_301 - Runtime error cancelling out of user
' TW        06/12/2006  EP2_305 - Wrong parameter passed to frmMaintainFSAAssociations
' TW        06/12/2006  EP2_296 - Black listing a client is not setting  user to not active
' TW        12/12/2006  EP2_438 - Meaningless error message if you add an email address that already exists in the database
' TW        14/12/2006  EP2_440 - Error adding new blacklisted firm
' TW        18/12/2006  EP2_568  Introducer Users need to have INTRODUCERUSER indicator set on OMIGAUSER table
' TW        09/01/2007  EP2_728 - Caching user numbers and Ids
' TW        10/01/2007  EP2_637 - Case not ingesting from the web onto omiga
' TW        16/01/2007  EP2_864 - Error amending individual broker record
' TW        25/01/2007  EP2_973 - Introducer ID should be greyed out
' TW        29/01/2007  EP2_863 - Add new col to IntroducerFirm - "Trading As"
' TW        05/02/2007  EP2_798 - Broker registration can be blocked as UserID corresponds to that created in Supervisor
' TW        09/02/2007  EP2_1296 - Error On the AR Firms Tab when selected the firm and clicked the Modify Button
' TW        14/05/2007  VR216 - Display up to 3 phone numbers for Introducers in Supervisor
' TW        22/05/2007  VR317 - Error editing Individual AR broker
' TW        15/06/2007  DBM405 - Add "No Cross Selling Agreement" processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Const mc_strTHIS_MODULE As String = "frmEditIndividualPackager"

' Private data
Private m_clsIntroducerTable As IntroducerTable
Private m_clsOrgUserTable As OrgUserTable

Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Private strIntroducerID As String
Private strARFirmID As String
' TW 14/12/2006 EP2_440
Private strCurrentListingStatus As String
' TW 14/12/2006 EP2_440 End
' TW 10/01/2007 EP2_637
Private strSearchType As String
' TW 10/01/2007 EP2_637 End
' TW 16/01/2007 EP2_864
Private strUserID As String
' TW 16/01/2007 EP2_864 End

' TW 14/05/2007 VR216
Private m_clsTelephoneTable As ContactDetailsTelephoneTable
Private m_clsUserContactTable As OrgUserContactTable
Private m_clsContact As ContactDetailsTable
' TW 14/05/2007 VR216 End
Private Sub SaveTelephoneDetails()
' TW 14/05/2007 VR216 - New routine
Dim clsGUID As New GuidAssist
Dim colMatchValues As Collection
Dim intMaxSequenceNumber As Integer
Dim intSequenceNumber As Integer
Dim vGUID As Variant
Dim X As Integer

    On Error GoTo Failed

    Set colMatchValues = New Collection
    colMatchValues.Add txtUserID.Text

    Set m_clsUserContactTable = New OrgUserContactTable
    Set m_clsTelephoneTable = New ContactDetailsTelephoneTable
    Set m_clsContact = New ContactDetailsTable
    TableAccess(m_clsUserContactTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsUserContactTable).GetTableData
    
    If TableAccess(m_clsUserContactTable).RecordCount = 0 Then
        
        vGUID = CVar(clsGUID.CreateGUID())
    
        If IsNull(vGUID) Then
            g_clsErrorHandling.RaiseError errGeneralError, "SaveTelephonDetails: ContactDetails GUID is empty"
        End If
        g_clsFormProcessing.CreateNewRecord m_clsContact

' Create the CONTACTDETAILS record (Empty except for GUID)
        m_clsContact.SetContactDetailsGUID vGUID
        TableAccess(m_clsContact).Update
        
' Create the ORGANISATIONUSERCONTACTDETAILS record
        g_clsFormProcessing.CreateNewRecord m_clsUserContactTable
        m_clsUserContactTable.SetUserID txtUserID.Text
        m_clsUserContactTable.SetContactDetailsGUID vGUID
        TableAccess(m_clsUserContactTable).Update
    Else
        vGUID = m_clsUserContactTable.GetContactDetailsGUID()
    End If
        
    Set colMatchValues = New Collection
    colMatchValues.Add vGUID
    
    TableAccess(m_clsTelephoneTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsTelephoneTable).GetTableData
    If TableAccess(m_clsTelephoneTable).RecordCount > 0 Then
' Clear contents of any existing records (this is necessary for promotions).
        TableAccess(m_clsTelephoneTable).MoveFirst
        For X = 1 To TableAccess(m_clsTelephoneTable).RecordCount
            intMaxSequenceNumber = TableAccess(m_clsTelephoneTable).GetRecordSet!TELEPHONESEQNUM
            m_clsTelephoneTable.SetType Null
        
            m_clsTelephoneTable.SetCOUNTRYCODE Null
            m_clsTelephoneTable.SetAREA_CODE Null
            m_clsTelephoneTable.SetTELEPHONE Null
            m_clsTelephoneTable.SetTELEPHONE_EXT Null
                
            TableAccess(m_clsTelephoneTable).Update
            TableAccess(m_clsTelephoneTable).MoveNext
        Next X
        TableAccess(m_clsTelephoneTable).MoveFirst
    End If
    For X = 1 To 3
        If cboTelephoneNumberType(X - 1).List(cboTelephoneNumberType(X - 1).ListIndex) <> COMBO_NONE And cboTelephoneNumberType(X - 1).ListIndex <> -1 Then
            intSequenceNumber = intSequenceNumber + 1
            If intSequenceNumber > intMaxSequenceNumber Then
' Create new Telephone record
                g_clsFormProcessing.CreateNewRecord m_clsTelephoneTable
                m_clsTelephoneTable.SetCONTACTDETAILSTELEPHONEGUID vGUID
                m_clsTelephoneTable.SetTELEPHONE_SEQ_NUMBER intSequenceNumber
            End If
            m_clsTelephoneTable.SetType cboTelephoneNumberType(X - 1).GetExtra(cboTelephoneNumberType(X - 1).ListIndex)
        
            m_clsTelephoneTable.SetCOUNTRYCODE txtCountryCode(X - 1).Text
            m_clsTelephoneTable.SetAREA_CODE txtAreaCode(X - 1).Text
            m_clsTelephoneTable.SetTELEPHONE txtTelephoneNumber(X - 1).Text
            m_clsTelephoneTable.SetTELEPHONE_EXT txtExtensionNumber(X - 1).Text
            TableAccess(m_clsTelephoneTable).Update
            If X < TableAccess(m_clsTelephoneTable).RecordCount Then
                TableAccess(m_clsTelephoneTable).MoveNext
            End If
        End If

    Next X
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub PopulateTelephoneDetails(Optional intIndex As Integer = -1)
' TW 14/05/2007 VR216 - New routine
' TW 22/05/2007 VR317 - Modified to deal with extraneous entries in CONTACTTELEPHONEDETAILS table
Dim colMatchValues As Collection

Dim vGUID As Variant
Dim X As Integer
Dim intCurrentEntry As Integer

Dim clsComboValueTable As ComboValueTable
Dim colComboMatchValues As Collection
Dim colComboMatchFields As Collection

    ' Combo Values
    Set clsComboValueTable = New ComboValueTable
    Set colComboMatchFields = New Collection
    colComboMatchFields.Add "GroupName"
    colComboMatchFields.Add "ValueId"
    TableAccess(clsComboValueTable).SetKeyMatchFields colComboMatchFields

    ' Contact Details
    Set colMatchValues = New Collection
    colMatchValues.Add txtUserID.Text

    Set m_clsUserContactTable = New OrgUserContactTable
    Set m_clsTelephoneTable = New ContactDetailsTelephoneTable
    
    TableAccess(m_clsUserContactTable).SetKeyMatchValues colMatchValues
    TableAccess(m_clsUserContactTable).GetTableData
    
    If TableAccess(m_clsUserContactTable).RecordCount > 0 Then
        vGUID = m_clsUserContactTable.GetContactDetailsGUID()
        
        If Len(vGUID) > 0 Then
            Set colMatchValues = New Collection
            colMatchValues.Add vGUID
            
            ' Now Telephone numbers.
            TableAccess(m_clsTelephoneTable).SetKeyMatchValues colMatchValues
            TableAccess(m_clsTelephoneTable).GetTableData
            If TableAccess(m_clsTelephoneTable).RecordCount > 0 Then
                TableAccess(m_clsTelephoneTable).MoveFirst
                
                For X = 1 To TableAccess(m_clsTelephoneTable).RecordCount
                    If Len(m_clsTelephoneTable.GetType()) > 0 Then
                        'Check to see if the Telephone Number record is a valid one, i.e there is an entry in the COMBOVALUE table for its type
                        Set colComboMatchValues = New Collection
                        colComboMatchValues.Add "ContactTelephoneUsage"
                        colComboMatchValues.Add m_clsTelephoneTable.GetType()
                        
                        TableAccess(clsComboValueTable).SetKeyMatchValues colComboMatchValues
                        TableAccess(clsComboValueTable).GetTableData
                        If TableAccess(clsComboValueTable).RecordCount > 0 Then
                            If intIndex = -1 Or intIndex = intCurrentEntry Then
                                txtCountryCode(intCurrentEntry).Text = m_clsTelephoneTable.GetCOUNTRYCODE()
                                txtAreaCode(intCurrentEntry).Text = m_clsTelephoneTable.GetAREA_CODE()
                                txtTelephoneNumber(intCurrentEntry).Text = m_clsTelephoneTable.GetTELEPHONE()
                                txtExtensionNumber(intCurrentEntry).Text = m_clsTelephoneTable.GetTELEPHONE_EXT()
                                
                                'Set the Number type combo selection.
                                g_clsFormProcessing.HandleComboExtra cboTelephoneNumberType(intCurrentEntry), m_clsTelephoneTable.GetType(), SET_CONTROL_VALUE
                            End If
                            intCurrentEntry = intCurrentEntry + 1
                            If intCurrentEntry > 2 Then
                                ' Only allow a maximum of 3 telephone numbers
                                Exit For
                            End If
                        End If
                    End If
                    TableAccess(m_clsTelephoneTable).MoveNext
                Next X
                For X = intCurrentEntry To 2
                    g_clsFormProcessing.HandleComboExtra cboTelephoneNumberType(X), COMBO_NONE, SET_CONTROL_VALUE
                Next X
            End If
            
        End If
    End If

End Sub


Private Sub cboTelephoneNumberType_Validate(Index As Integer, Cancel As Boolean)
' TW 14/05/2007 VR216 - New Routien
Dim blnChanged As Boolean
    blnChanged = ((Len(txtCountryCode(Index).Text) + Len(txtAreaCode(Index).Text) + Len(txtTelephoneNumber(Index).Text) + Len(txtExtensionNumber(Index).Text)) > 0)
    If cboTelephoneNumberType(Index).List(cboTelephoneNumberType(Index).ListIndex) = COMBO_NONE And blnChanged Then
        If MsgBox("Are you sure you want to delete this entry", vbQuestion + vbYesNo, "Delete Entry") = vbYes Then
            txtCountryCode(Index).Text = ""
            txtAreaCode(Index).Text = ""
            txtTelephoneNumber(Index).Text = ""
            txtExtensionNumber(Index).Text = ""
        Else
            Cancel = True
            PopulateTelephoneDetails Index
        End If
    End If
End Sub


Private Sub CreateNewUser(strUnitID As String)
' TW 10/01/2007 EP2_637 - New

Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim vTitle As Variant
Dim vDateOfBirth As Variant

    On Error GoTo Failed
    
    If cboUserTitle.ListIndex = -1 Then
        vTitle = 99 'Other
    Else
        g_clsFormProcessing.HandleComboExtra cboUserTitle, vTitle, GET_CONTROL_VALUE
    End If
    If IsDate(txtUserDateOfBirth.Text) Then
        vDateOfBirth = CStr(txtUserDateOfBirth.Text)
    Else
        vDateOfBirth = Null
    End If

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "usp_CreateIndividualIntroducerUser"
' TW 16/01/2007 EP2_864
'        .Parameters.Append .CreateParameter("UserId", adVarChar, adParamInput, 24, strIntroducerID)
        .Parameters.Append .CreateParameter("UserId", adVarChar, adParamInput, 24, strUserID)
' TW 16/01/2007 EP2_864 End
        .Parameters.Append .CreateParameter("UnitId", adVarChar, adParamInput, 20, strUnitID)
        .Parameters.Append .CreateParameter("UserSurname", adVarChar, adParamInput, 70, txtUserSurname.Text)
        .Parameters.Append .CreateParameter("UserForename", adVarChar, adParamInput, 60, txtUserFirstForename.Text)
        .Parameters.Append .CreateParameter("Initials", adVarChar, adParamInput, 3, txtUserInitials.Text)
        .Parameters.Append .CreateParameter("Title", adInteger, adParamInput, 5, vTitle)
        .Parameters.Append .CreateParameter("DateOfBirth", adDate, adParamInput, , vDateOfBirth)
        .Execute , , adExecuteNoRecords
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub UpdateUserFields()
' TW 10/01/2007 EP2_637 - New
Dim m_clsOrgUserTable As OrgUserTable
Dim m_clsOmigaUserTable As OmigaUserTable
Dim clsTableAccess As TableAccess
Dim colKeys As Collection
Dim vTitle As Variant
' Update OrganisationUser record

    If cboUserTitle.ListIndex = -1 Then
        vTitle = 99 'Other
    Else
        g_clsFormProcessing.HandleComboExtra cboUserTitle, vTitle, GET_CONTROL_VALUE
    End If

    Set m_clsOrgUserTable = New OrgUserTable
    Set colKeys = New Collection
' TW 16/01/2007 EP2_864
'    colKeys.Add strIntroducerID
    colKeys.Add strUserID
' TW 16/01/2007 EP2_864 End
    Set clsTableAccess = m_clsOrgUserTable
    clsTableAccess.SetKeyMatchValues colKeys
    clsTableAccess.GetTableData POPULATE_KEYS
    
    m_clsOrgUserTable.SetUserTitle CStr(vTitle)
    m_clsOrgUserTable.SetInitials txtUserInitials.Text
    m_clsOrgUserTable.SetForename txtUserFirstForename.Text
    m_clsOrgUserTable.SetUserSurname txtUserSurname.Text
    If IsDate(txtUserDateOfBirth.Text) Then
        m_clsOrgUserTable.SetDateOfBirth txtUserDateOfBirth.Text
    End If
    
    clsTableAccess.Update
    
' Update OmigaUser record
    Set m_clsOmigaUserTable = New OmigaUserTable
    Set colKeys = New Collection
' TW 16/01/2007 EP2_864
'    colKeys.Add strIntroducerID
    colKeys.Add strUserID
' TW 16/01/2007 EP2_864 End
    Set clsTableAccess = m_clsOmigaUserTable
    clsTableAccess.SetKeyMatchValues colKeys
    g_clsHandleUpdates.SaveChangeRequest m_clsOmigaUserTable
    
End Sub
Private Sub DealWithBlacklisting()
' TW 14/12/2006 EP2_440
Dim m_clsOmigaUserTable As OmigaUserTable
Dim clsTableAccess As TableAccess
Dim colKeys As Collection
    If cboListingStatus.ListIndex > -1 Then
        If cboListingStatus.GetExtra(cboListingStatus.ListIndex) = 1 And strCurrentListingStatus <> "1" Then
        ' Update OmigaUser record
            Set m_clsOmigaUserTable = New OmigaUserTable
            Set colKeys = New Collection
' TW 16/01/2007 EP2_864
        '    colKeys.Add strIntroducerID
            colKeys.Add strUserID
' TW 16/01/2007 EP2_864 End
            Set clsTableAccess = m_clsOmigaUserTable
            clsTableAccess.SetKeyMatchValues colKeys
            clsTableAccess.GetTableData POPULATE_KEYS
            m_clsOmigaUserTable.SetActiveTo Now
            clsTableAccess.Update
            g_clsHandleUpdates.SaveChangeRequest m_clsOmigaUserTable
        End If
    End If
' TW 14/12/2006 EP2_440 End
End Sub
Private Sub PopulateAssociatedFirmsListView()
Dim m_clsIntroducerPackIndTable As New IntroducerPackIndTable
Dim colMatchValues As Collection

    On Error GoTo Failed
    lvFirms.ListItems.Clear
    
    Set colMatchValues = New Collection
    colMatchValues.Add ""

    m_clsIntroducerPackIndTable.GetLinkedPrincipalFirmRecords strIntroducerID
    lvFirms.LoadColumnDetails TypeName(m_clsIntroducerPackIndTable)
    
    TableAccess(m_clsIntroducerPackIndTable).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmEditIndividualPackager.lvFirms, m_clsIntroducerPackIndTable
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
' TW 06/12/2006 EP2_301
        If Len(txtUserID.Text) = 0 Then
            MsgBox "An Omiga User must be created before this record can be saved", vbExclamation
            bRet = False
        Else
' TW 06/12/2006 EP2_301 End
            If lvFirms.ListItems.Count = 0 Then
                bRet = False
                MsgBox "There are no Principal(s) associated with this Introducer. Please make an association"
                SSTab1.Tab = 2
            Else
' TW 10/01/2007 EP2_637
                UpdateUserFields
' TW 10/01/2007 EP2_637 End
                DealWithBlacklisting
                SaveScreenData
' TW 14/05/2007 VR216
                SaveTelephoneDetails
' TW 14/05/2007 VR216 End
                SaveChangeRequest
            End If
' TW 06/12/2006 EP2_301
        End If
' TW 06/12/2006 EP2_301 End
    End If
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function
Private Sub PopulateUserDetails()

Dim colMatchValues As Collection

    Set m_clsOrgUserTable = New OrgUserTable
    Set colMatchValues = New Collection
    colMatchValues.Add txtUserID.Text
    
    TableAccess(m_clsOrgUserTable).SetKeyMatchValues colMatchValues
    
    TableAccess(m_clsOrgUserTable).GetTableData
    
    g_clsFormProcessing.HandleComboExtra cboUserTitle, m_clsOrgUserTable.GetTitle, SET_CONTROL_VALUE
    txtUserInitials.Text = m_clsOrgUserTable.GetInitials
    txtUserFirstForename.Text = m_clsOrgUserTable.GetForename
    txtUserSurname.Text = m_clsOrgUserTable.GetSurname
    txtUserDateOfBirth.Text = m_clsOrgUserTable.GetDateOfBirth
End Sub

Private Sub SaveChangeRequest()
Dim sProductNumber As String
Dim colMatchValues As Collection
Dim clsTableAccess As TableAccess
    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add strIntroducerID
    
    Set clsTableAccess = m_clsIntroducerTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsIntroducerTable
    
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
    
    On Error GoTo Failed
    Set clsTableAccess = m_clsIntroducerTable
    
    m_clsIntroducerTable.SetIntroducerID strIntroducerID
    m_clsIntroducerTable.SetUserID txtUserID.Text
    m_clsIntroducerTable.SetNationalInsuranceNumber txtNationalInsuranceNumber.Text
    m_clsIntroducerTable.SetARIndicator "0"
    If cboIntroducerType.ListIndex > -1 Then
        m_clsIntroducerTable.SetIntroducerType cboIntroducerType.GetExtra(cboIntroducerType.ListIndex)
    End If
    If cboIntroducerStatus.ListIndex > -1 Then
        m_clsIntroducerTable.SetIntroducerStatus cboIntroducerStatus.GetExtra(cboIntroducerStatus.ListIndex)
    End If
    If cboListingStatus.ListIndex > -1 Then
        m_clsIntroducerTable.SetListingStatus cboListingStatus.GetExtra(cboListingStatus.ListIndex)
    End If
    m_clsIntroducerTable.SetEmailAddress txtEmailAddress.Text
    m_clsIntroducerTable.SetBuildingOrHouseNumber txtBuildingNumber.Text
    m_clsIntroducerTable.SetFlatNumber txtFlatNumber.Text
    m_clsIntroducerTable.SetBuildingOrHouseName txtBuildingName.Text
    m_clsIntroducerTable.SetStreet txtStreet.Text
    m_clsIntroducerTable.SetDistrict txtDistrict.Text
    m_clsIntroducerTable.SetTown txtTown.Text
    m_clsIntroducerTable.SetCounty txtCounty.Text
    m_clsIntroducerTable.SetPostcode txtPostCode.Text
    If cboCountry.ListIndex > -1 Then
        m_clsIntroducerTable.SetCountry cboCountry.GetExtra(cboCountry.ListIndex)
    End If
' TW 18/11/2006 EP2_132
    If optARIndicator(0).Value = True Then
        m_clsIntroducerTable.SetARIndicator 1
    Else
        m_clsIntroducerTable.SetARIndicator 0
    End If
    If optMarketingOptOut(0).Value = True Then
        m_clsIntroducerTable.SetMarketingDataOptOut 1
    Else
        m_clsIntroducerTable.SetMarketingDataOptOut 0
    End If
    m_clsIntroducerTable.SetLastupdateddate Now
    m_clsIntroducerTable.SetLastUpdatedBy g_sSupervisorUser
' TW 18/11/2006 EP2_132 End
' TW 15/06/2007 DBM405
    If optNoCrossSellAgreement(0).Value = True Then
        m_clsIntroducerTable.SetNoCrossSellInd 1
    Else
        m_clsIntroducerTable.SetNoCrossSellInd 0
    End If
' TW 15/06/2007 DBM405 End
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


Private Sub UpdateLinks()
' TW 10/01/2007 EP2_637 - New
' This routine checks for changes to links made by frmMaintainFSAAssociations

Dim X As Long
Dim clsTableAccess As TableAccess

Dim colMatchValues As Collection
Dim colFields As Collection

Dim strIntroducerFirmSeqNo As String
Dim m_clsIntroducerFirmTable As IntroducerFirmTable

Dim blnFound As Boolean
Dim rs As ADODB.Recordset
Dim strId As String
Dim strSQL As String
' TW 29/01/2007 EP2_863
Dim strTradingAs As String
' TW 29/01/2007 EP2_863 End

    On Error GoTo Failed:
    
' Add Link from Introducer via Introducer Firm to ARFirm or PrincipalFirm

    Set m_clsIntroducerFirmTable = New IntroducerFirmTable
    Set clsTableAccess = m_clsIntroducerFirmTable
    For X = 1 To frmMaintainFSAAssociations.lvLinkedItems.ListItems.Count
            
' TW 29/01/2007 EP2_863
        strTradingAs = frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(X).SubItems(6)
        strId = frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(X).SubItems(1)
        Set m_colKeys = New Collection
' TW 29/01/2007 EP2_863 End

        If frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(X).SubItems(5) = 0 Then

            strIntroducerFirmSeqNo = GetIntroducerFirmSeqNo()
            If Len(strIntroducerFirmSeqNo) = 0 Then
                MsgBox "Can't allocate an identity number for this record", vbCritical
                Exit Sub
            End If

            Set m_colKeys = New Collection
                   
            m_colKeys.Add strIntroducerFirmSeqNo
            g_clsFormProcessing.CreateNewRecord m_clsIntroducerFirmTable
            m_clsIntroducerFirmTable.SetIntroducerFirmSequenceNumber strIntroducerFirmSeqNo
            m_clsIntroducerFirmTable.SetIntroducerID strIntroducerID
            Select Case strSearchType
                Case "Individual Packager", "DA Individual Broker"
                    m_clsIntroducerFirmTable.SetPrincipalfirmID strId
                Case "AR Individual Broker"
                    m_clsIntroducerFirmTable.SetARFirmID strId
            End Select
            
' TW 29/01/2007 EP2_863
            m_clsIntroducerFirmTable.SetTradingAs strTradingAs
' TW 29/01/2007 EP2_863 End
              
            clsTableAccess.Update
            clsTableAccess.SetKeyMatchValues m_colKeys
            g_clsHandleUpdates.SaveChangeRequest m_clsIntroducerFirmTable
' TW 29/01/2007 EP2_863
        Else
            m_colKeys.Add frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(X).SubItems(5)
            clsTableAccess.SetKeyMatchValues m_colKeys
            clsTableAccess.GetTableData
            If strTradingAs <> m_clsIntroducerFirmTable.GetTradingAs Then
                m_clsIntroducerFirmTable.SetTradingAs strTradingAs
                clsTableAccess.Update
                clsTableAccess.SetKeyMatchValues m_colKeys
                g_clsHandleUpdates.SaveChangeRequest m_clsIntroducerFirmTable
            End If
        End If
' TW 29/01/2007 EP2_863 End
    Next X

' Remove Link from Introducer via Introducer Firm to ARFirm or PrincipalFirm
    Set colFields = New Collection
    colFields.Add "INTRODUCERFIRMSEQNO"
    clsTableAccess.SetKeyMatchFields colFields
    
    strSQL = "SELECT * FROM INTRODUCERFIRM WHERE INTRODUCERID = '" & strIntroducerID & "' AND NOT PRINCIPALFIRMID IS NULL"
    Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSQL)
    Do While Not rs.EOF
        blnFound = False
        For X = 1 To frmMaintainFSAAssociations.lvLinkedItems.ListItems.Count
            strId = frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(X).SubItems(1)
            Select Case strSearchType
                Case "Individual Packager", "DA Individual Broker"
                    blnFound = (strId = rs!PRINCIPALFIRMID)
                Case "AR Individual Broker"
                    blnFound = (strId = rs!ARFIRMID)
            End Select
            If blnFound Then
                Exit For
            End If
        Next X
        If Not blnFound Then
            Set colMatchValues = New Collection
    
            colMatchValues.Add rs!INTRODUCERFIRMSEQNO
            clsTableAccess.SetKeyMatchValues colMatchValues

            clsTableAccess.DeleteRow colMatchValues
            clsTableAccess.Update
        
            g_clsHandleUpdates.SaveChangeRequest m_clsIntroducerFirmTable, , PromoteDelete
        End If
        rs.MoveNext
    Loop
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

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



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when this form is first loaded - called autmomatically by VB. Need to
'                 perform all initialisation processing here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCountry_Validate(Cancel As Boolean)
    Cancel = Not cboCountry.ValidateData()
End Sub


Private Sub cboIntroducerType_Validate(Cancel As Boolean)
    Cancel = Not cboIntroducerType.ValidateData()
End Sub


Private Sub cboListingStatus_Validate(Cancel As Boolean)
    Cancel = Not cboListingStatus.ValidateData()
End Sub


'Private Sub cboType1_Validate(Cancel As Boolean)
'    Cancel = Not cboType1.ValidateData()
'End Sub


'Private Sub cboType2_Validate(Cancel As Boolean)
'    Cancel = Not cboType2.ValidateData()
'End Sub


'Private Sub cboType3_Validate(Cancel As Boolean)
'    Cancel = Not cboType3.ValidateData()
'End Sub


Private Sub cboUserTitle_Validate(Cancel As Boolean)
    Cancel = Not cboUserTitle.ValidateData()
End Sub


Private Sub cmdAddFirm_Click()
' TW 10/01/2007 EP2_637 - Rewritten
Dim colKeys As New Collection

    On Error GoTo Failed:
    
' TW 09/02/2007 EP2_1296
    colKeys.Add strUserID
' TW 09/02/2007 EP2_1296 End
    colKeys.Add strIntroducerID

    frmMaintainFSAAssociations.SetKeys colKeys
    frmMaintainFSAAssociations.SetIsEdit m_bIsEdit
    frmMaintainFSAAssociations.txtSearchType = strSearchType
    frmMaintainFSAAssociations.Show 1
    If frmMaintainFSAAssociations.GetReturnCode() Then
        If frmMaintainFSAAssociations.lvLinkedItems.ListItems.Count > 0 Then
            If m_bIsEdit = False Then
                CreateNewUser frmMaintainFSAAssociations.lvLinkedItems.ListItems.Item(1)
                SaveScreenData
            End If
        End If
        UpdateLinks
        
        PopulateAssociatedFirmsListView
        
    End If
    
    Unload frmMaintainFSAAssociations
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub SetIntroducerUser()
Dim strSQL As String
    On Error GoTo Failed:
    
' TW 16/01/2007 EP2_864
'    strSQL = "UPDATE OMIGAUSER SET INTRODUCERUSER = 1 WHERE USERID = '" & strIntroducerID & "'"
    strSQL = "UPDATE OMIGAUSER SET INTRODUCERUSER = 1 WHERE USERID = '" & strUserID & "'"
' TW 16/01/2007 EP2_864 End
    g_clsDataAccess.GetActiveConnection.Execute strSQL
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
        
    Set m_clsIntroducerTable = New IntroducerTable
    g_clsFormProcessing.PopulateCombo "Country", Me.cboCountry
    g_clsFormProcessing.PopulateCombo "IntroducerType", Me.cboIntroducerType
    g_clsFormProcessing.PopulateCombo "IntroducerStatus", Me.cboIntroducerStatus
    g_clsFormProcessing.PopulateCombo "ListingStatus", Me.cboListingStatus
    g_clsFormProcessing.PopulateCombo "Title", Me.cboUserTitle
    
' TW 14/05/2007 VR216
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboTelephoneNumberType(0)
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboTelephoneNumberType(1)
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboTelephoneNumberType(2)
' TW 14/05/2007 VR216 End
    
    cboIntroducerType.Text = "Packager"
    cboIntroducerType.Enabled = False
    
    SetupAssociatedFirmsHeaders
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    PopulateAssociatedFirmsListView
    strSearchType = "Individual Packager"

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetEditState()
    
    Const c_strTHIS_FUNCTION As String = "SetEditState"
    Dim strErrorHelp As String
    On Error GoTo eh_SetEditState
    
    Dim objRS As ADODB.Recordset
    Dim objTableAccess As TableAccess
    Dim sGuid As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection
    
    ' First, the Introducer table
    Set objTableAccess = m_clsIntroducerTable
    objTableAccess.SetKeyMatchValues m_colKeys
    
    Set objRS = objTableAccess.GetTableData()
    
    If Not objRS Is Nothing Then
        If objRS.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Edit Packager - no records exist"
        End If
    End If
    '
CleanUp:
    ' destroy objects etc here
    Set objRS = Nothing
    Set objTableAccess = Nothing
    Exit Sub
    '
eh_SetEditState:
    '
    Err.Raise Err.Number, mc_strTHIS_MODULE & "[" & strErrorHelp & "]." & _
                            c_strTHIS_FUNCTION & "[" & Err.Source & "]", _
                            Err.DESCRIPTION
    Resume CleanUp
    '
End Sub
Private Function CheckEmailAddress() As Boolean
Dim blnValid As Boolean

Dim strEmailAddress As String
Dim strSearch As String

Dim rs As ADODB.Recordset

    On Error GoTo Failed
    
    blnValid = True
    strEmailAddress = Trim$(txtEmailAddress.Text)
    If Len(strEmailAddress) = 0 Then
        blnValid = False
    Else
        strSearch = "SELECT COUNT(1) FROM INTRODUCER WHERE EMAILADDRESS = '" & strEmailAddress & "'"
        Set rs = g_clsDataAccess.GetActiveConnection.Execute(strSearch)
        
        If m_bIsEdit Then
            If LCase$(m_clsIntroducerTable.GetEmailAddress) = LCase$(strEmailAddress) Then
                blnValid = True
            Else
                blnValid = (rs(0) = 0)
            End If
        Else
            blnValid = (rs(0) = 0)
        End If
    End If
    If Not blnValid Then
        MsgBox "A unique valid email address must be entered", vbInformation
    End If
    CheckEmailAddress = blnValid
    
    Set rs = Nothing
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    CheckEmailAddress = False
End Function

Private Sub SetAddState()
    
    On Error GoTo Failed
    
    Call GetNewIntroducerID
' TW 16/01/2007 EP2_864
' TW 10/01/2007 EP2_637
'    txtUserID.Text = strIntroducerID
    Call GetNewUserID
    txtUserID.Text = strUserID
' TW 10/01/2007 EP2_637 End
' TW 16/01/2007 EP2_864 End
    lblIntroducerID.Caption = strIntroducerID
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add strIntroducerID
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsIntroducerTable
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetNewIntroducerID()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "IntroducerId")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strIntroducerID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strIntroducerID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub


Private Sub PopulateScreenFields()
    
Const c_strTHIS_FUNCTION As String = "PopulateScreenFields"
Dim strErrorHelp As String

    On Error GoTo eh_PopulateScreenFields
    
    '
    ' First tab (User details)
    '
    strErrorHelp = "Populating first tab - Introducer Details"
    
' TW 14/12/2006 EP2_440
    strCurrentListingStatus = m_clsIntroducerTable.GetListingStatus
' TW 14/12/2006 EP2_440 End
    
    strIntroducerID = m_colKeys(1)
    lblIntroducerID.Caption = m_clsIntroducerTable.GetIntroducerID
    txtUserID.Text = m_clsIntroducerTable.GetUserID
' TW 16/01/2007 EP2_864
    strUserID = m_clsIntroducerTable.GetUserID
' TW 16/01/2007 EP2_864 End
    txtNationalInsuranceNumber.Text = m_clsIntroducerTable.GetNationalInsuranceNumber
    g_clsFormProcessing.HandleComboExtra cboIntroducerStatus, m_clsIntroducerTable.GetIntroducerStatus, SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra cboListingStatus, m_clsIntroducerTable.GetListingStatus, SET_CONTROL_VALUE
    txtEmailAddress.Text = m_clsIntroducerTable.GetEmailAddress
    lblFailedLoginAttempts.Caption = m_clsIntroducerTable.GetNumberOfLoginFailures
    
    txtBuildingNumber.Text = m_clsIntroducerTable.GetBuildingOrHouseNumber
    txtFlatNumber.Text = m_clsIntroducerTable.GetFlatNumber
    txtBuildingName.Text = m_clsIntroducerTable.GetBuildingOrHouseName
    txtStreet.Text = m_clsIntroducerTable.GetStreet
    txtDistrict.Text = m_clsIntroducerTable.GetDistrict
    txtTown.Text = m_clsIntroducerTable.GetTown
    txtCounty.Text = m_clsIntroducerTable.GetCounty
    txtPostCode.Text = m_clsIntroducerTable.GetPostcode
' TW 18/11/2006 EP2_132
    If m_clsIntroducerTable.GetMarketingDataOptOut = "1" Then
        optMarketingOptOut(0) = True
    Else
        optMarketingOptOut(1) = True
    End If
    If m_clsIntroducerTable.GetARIndicator = "1" Then
        optARIndicator(0) = True
    Else
        optARIndicator(1) = True
    End If
' TW 18/11/2006 EP2_132 End
' TW 15/06/2007 DBM405
    If m_clsIntroducerTable.GetNoCrossSellInd = "1" Then
        optNoCrossSellAgreement(0) = True
    Else
        optNoCrossSellAgreement(1) = True
    End If
' TW 15/06/2007 DBM405 End
    
    g_clsFormProcessing.HandleComboExtra cboCountry, m_clsIntroducerTable.GetCountry, SET_CONTROL_VALUE
    
    '
    ' Second tab (User Details)
    '
    PopulateUserDetails
    
' TW 14/05/2007 VR216
    PopulateTelephoneDetails
' TW 14/05/2007 VR216 End
    
CleanUp:
    Exit Sub
    '
eh_PopulateScreenFields:
    '
    Err.Raise Err.Number, mc_strTHIS_MODULE & "[" & strErrorHelp & "]." & _
                            c_strTHIS_FUNCTION & "[" & Err.Source & "]", _
                            Err.DESCRIPTION
    Resume CleanUp
    '
End Sub
Private Sub GetNewUserID()
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
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "UserId")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strUserID = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strUserID) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
' TW 05/02/2007 EP2_798
    Else
        strUserID = g_clsGlobalParameter.FindString("IntroducerUserPrefix") & strUserID
' TW 05/02/2007 EP2_798 End
    End If
Failed:
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
    Set m_clsIntroducerTable = Nothing
End Sub

Private Sub lvFirms_DblClick()
    Call cmdAddFirm_Click
End Sub

Private Sub txtBuildingName_Validate(Cancel As Boolean)
    Cancel = Not txtBuildingName.ValidateData()
End Sub


Private Sub txtBuildingNumber_Validate(Cancel As Boolean)
    Cancel = Not txtBuildingNumber.ValidateData()
End Sub


Private Sub txtCounty_Validate(Cancel As Boolean)
    Cancel = Not txtCounty.ValidateData()
End Sub


Private Sub txtDistrict_Validate(Cancel As Boolean)
    Cancel = Not txtDistrict.ValidateData()
End Sub


Private Sub txtEmailAddress_Validate(Cancel As Boolean)
    Cancel = Not txtEmailAddress.ValidateData()
' TW 12/12/2006 EP2_438
    If Not Cancel Then
        Cancel = Not CheckEmailAddress()
    End If
' TW 12/12/2006 EP2_438 End
End Sub


Private Sub txtFlatNumber_Validate(Cancel As Boolean)
    Cancel = Not txtFlatNumber.ValidateData()
End Sub


Private Sub txtNationalInsuranceNumber_Validate(Cancel As Boolean)
    Cancel = Not txtNationalInsuranceNumber.ValidateData()
End Sub


Private Sub txtPostCode_Validate(Cancel As Boolean)
    Cancel = Not txtPostCode.ValidateData()
End Sub


Private Sub txtStreet_Validate(Cancel As Boolean)
    Cancel = Not txtStreet.ValidateData()
End Sub


Private Sub txtTown_Validate(Cancel As Boolean)
    Cancel = Not txtTown.ValidateData()
End Sub


Private Sub txtUserDateOfBirth_Validate(Cancel As Boolean)
    Cancel = Not txtUserDateOfBirth.ValidateData()
End Sub


Private Sub txtUserFirstForename_Validate(Cancel As Boolean)
    Cancel = Not txtUserFirstForename.ValidateData()
End Sub


Private Sub txtUserID_Validate(Cancel As Boolean)
    Cancel = Not txtUserID.ValidateData()
End Sub


Private Sub txtUserInitials_Validate(Cancel As Boolean)
    Cancel = Not txtUserInitials.ValidateData()
End Sub


Private Sub txtUserSurname_Validate(Cancel As Boolean)
    Cancel = Not txtUserSurname.ValidateData()
End Sub


Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Sub SetIsEdit(blnEditStatus As Boolean)
    m_bIsEdit = blnEditStatus
End Sub
Public Function GetIsEdit() As Boolean
    GetIsEdit = m_bIsEdit
End Function

Private Sub SetupAssociatedFirmsHeaders()
Dim headers As New Collection
Dim lvHeaders As listViewAccess

    On Error GoTo Failed

    lvHeaders.nWidth = 10
    lvHeaders.sName = "Unit Id"
    headers.Add lvHeaders

    lvHeaders.nWidth = 10
    lvHeaders.sName = "FSA Ref"
    headers.Add lvHeaders

    lvHeaders.nWidth = 10
    lvHeaders.sName = "Active"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Role"
    headers.Add lvHeaders

' TW 29/01/2007 EP2_863
'    lvHeaders.nWidth = 55
    lvHeaders.nWidth = 40
' TW 29/01/2007 EP2_863 End
    lvHeaders.sName = "Firm Details"
    headers.Add lvHeaders

' TW 29/01/2007 EP2_863
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Trading As"
    headers.Add lvHeaders
' TW 29/01/2007 EP2_863 End

    lvFirms.AddHeadings headers
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


