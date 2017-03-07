VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditUser 
   Caption         =   "Add/Edit User"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "frmEditUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8595
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timerPasswordEnable 
      Enabled         =   0   'False
      Left            =   240
      Top             =   6720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Tag             =   "Tab3"
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "User Details"
      TabPicture(0)   =   "frmEditUser.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Qualifications"
      TabPicture(1)   =   "frmEditUser.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Competency History"
      TabPicture(2)   =   "frmEditUser.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "User Roles"
      TabPicture(3)   =   "frmEditUser.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdEditUserRole"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdAddUserRole"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdDeleteUserRole"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Supervisor Access"
      TabPicture(4)   =   "frmEditUser.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraFrame"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.CommandButton cmdDeleteUserRole 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72420
         TabIndex        =   53
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddUserRole 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditUserRole 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   -73560
         TabIndex        =   51
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Supervisor Items"
         Height          =   5055
         Left            =   -74760
         TabIndex        =   48
         Top             =   840
         Width           =   7695
         Begin MSGOCX.MSGHorizontalSwapList swpSecurityResources 
            Height          =   4455
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   7858
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5655
         Left            =   120
         TabIndex        =   8
         Tag             =   "Tab1"
         Top             =   720
         Width           =   7995
         Begin VB.Frame fraWorkgroup 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   4080
            TabIndex        =   57
            Top             =   120
            Width           =   3255
            Begin VB.OptionButton optWorkgroupUser 
               Caption         =   "No"
               Height          =   195
               Index           =   1
               Left            =   2280
               TabIndex        =   59
               Top             =   40
               Width           =   615
            End
            Begin VB.OptionButton optWorkgroupUser 
               Caption         =   "Yes"
               Height          =   195
               Index           =   0
               Left            =   1540
               TabIndex        =   58
               Top             =   40
               Width           =   615
            End
            Begin VB.Label lblWorkgroupUser 
               AutoSize        =   -1  'True
               Caption         =   "Workgroup User?"
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   40
               Width           =   1260
            End
         End
         Begin VB.OptionButton optCreditCheckAccess 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   56
            Top             =   2784
            Width           =   735
         End
         Begin VB.OptionButton optCreditCheckAccess 
            Caption         =   "Yes"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   55
            Top             =   2784
            Width           =   615
         End
         Begin MSGOCX.MSGPasswordEditBox txtPassword 
            Height          =   315
            Left            =   1920
            TabIndex        =   20
            Top             =   3468
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Mandatory       =   -1  'True
         End
         Begin MSGOCX.MSGTextMulti txtNotes 
            Height          =   555
            Left            =   1920
            TabIndex        =   23
            Top             =   4280
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   979
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
            MaxLength       =   500
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   9
            Top             =   120
            Width           =   1755
            _ExtentX        =   3096
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
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   1
            Left            =   5640
            TabIndex        =   11
            Top             =   491
            Width           =   1755
            _ExtentX        =   3096
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   2
            Left            =   1920
            TabIndex        =   12
            Top             =   864
            Width           =   5835
            _ExtentX        =   10292
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
            MaxLength       =   30
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   3
            Left            =   1920
            TabIndex        =   13
            Top             =   1236
            Width           =   5835
            _ExtentX        =   10292
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
            MaxLength       =   35
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   4
            Left            =   1920
            TabIndex        =   14
            Top             =   1608
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            Mask            =   "##/##/####"
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   5
            Left            =   1920
            TabIndex        =   17
            Top             =   2352
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            Mask            =   "##/##/####"
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   6
            Left            =   5640
            TabIndex        =   18
            Top             =   2352
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   8
            Left            =   1920
            TabIndex        =   19
            Top             =   3096
            Width           =   5835
            _ExtentX        =   10292
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
            MaxLength       =   35
         End
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   10
            Left            =   1920
            TabIndex        =   22
            Top             =   3840
            Width           =   1755
            _ExtentX        =   3096
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
         Begin MSGOCX.MSGComboBox cboAccessType 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   1980
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
            Mandatory       =   -1  'True
            Text            =   ""
         End
         Begin MSGOCX.MSGComboBox cboTitle 
            Height          =   315
            Left            =   1920
            TabIndex        =   10
            Top             =   492
            Width           =   1755
            _ExtentX        =   3096
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
         Begin MSGOCX.MSGPasswordEditBox txtPasswordConfirm 
            Height          =   315
            Left            =   5640
            TabIndex        =   21
            Top             =   3468
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            Mandatory       =   -1  'True
         End
         Begin MSGOCX.MSGComboBox cboWorkingHours 
            Height          =   315
            Left            =   5640
            TabIndex        =   16
            Top             =   1980
            Width           =   2115
            _ExtentX        =   3731
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   11
            Left            =   1920
            TabIndex        =   24
            Top             =   5200
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   12
            Left            =   3080
            TabIndex        =   25
            Top             =   5200
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   13
            Left            =   4360
            TabIndex        =   26
            Top             =   5200
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
         Begin MSGOCX.MSGEditBox txtUserDetails 
            Height          =   315
            Index           =   14
            Left            =   6360
            TabIndex        =   27
            Top             =   5200
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
         Begin VB.Label lblCreditCheck 
            AutoSize        =   -1  'True
            Caption         =   "Credit Check Access?"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   2784
            Width           =   1575
         End
         Begin VB.Label lblContact 
            AutoSize        =   -1  'True
            Caption         =   "Ext. No."
            Height          =   195
            Index           =   11
            Left            =   6360
            TabIndex        =   47
            Top             =   4940
            Width           =   570
         End
         Begin VB.Label lblContact 
            AutoSize        =   -1  'True
            Caption         =   "Country Code"
            Height          =   195
            Index           =   10
            Left            =   1920
            TabIndex        =   46
            Top             =   4940
            Width           =   960
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Area Code"
            Height          =   195
            Left            =   3080
            TabIndex        =   45
            Top             =   4940
            Width           =   750
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Telephone Number"
            Height          =   195
            Left            =   4360
            TabIndex        =   44
            Top             =   4940
            Width           =   1365
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Working Hours"
            Height          =   195
            Index           =   15
            Left            =   4200
            TabIndex        =   43
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "User Title"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   551
            Width           =   675
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "User ID"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   180
            Width           =   540
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "User Initials"
            Height          =   195
            Index           =   2
            Left            =   4200
            TabIndex        =   40
            Top             =   551
            Width           =   810
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "User First Forename"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   922
            Width           =   1410
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "User Surname"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   1296
            Width           =   1005
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   1668
            Width           =   885
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Access Type"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   2040
            Width           =   930
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Active To"
            Height          =   195
            Index           =   8
            Left            =   4200
            TabIndex        =   35
            Top             =   2412
            Width           =   690
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Active From"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   34
            Top             =   2412
            Width           =   840
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Mother's Maiden Name"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   33
            Top             =   3156
            Width           =   1635
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   32
            Top             =   3528
            Width           =   690
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Emergency Password"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   31
            Top             =   3900
            Width           =   1530
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Notes"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   30
            Top             =   4320
            Width           =   420
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Contact Telephone"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   5260
            Width           =   1365
         End
         Begin VB.Label lblPasswordConfirm 
            AutoSize        =   -1  'True
            Caption         =   "Confirm Password"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4200
            TabIndex        =   28
            Top             =   3528
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "User Roles"
         Height          =   4695
         Left            =   -74700
         TabIndex        =   7
         Tag             =   "Tab4"
         Top             =   840
         Width           =   7635
         Begin MSGOCX.MSGListView lvUserRoles 
            Height          =   4095
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   7223
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Competency Details"
         Height          =   4455
         Left            =   -74700
         TabIndex        =   5
         Top             =   840
         Width           =   7635
         Begin MSGOCX.MSGDataGrid dgCompetencies 
            Height          =   3375
            Left            =   360
            TabIndex        =   6
            Top             =   840
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5953
            Enabled         =   -1  'True
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
      Begin VB.Frame Frame1 
         Caption         =   "Qualification Details"
         Height          =   4455
         Left            =   -74700
         TabIndex        =   3
         Tag             =   "Tab2"
         Top             =   840
         Width           =   7635
         Begin MSGOCX.MSGDataGrid dgQualifications 
            Height          =   3375
            Left            =   360
            TabIndex        =   4
            Top             =   840
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5953
            Enabled         =   -1  'True
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
   End
End
Attribute VB_Name = "frmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : frmEditUser
' Description   : A data-entry form for Users.
' History       :
'
' Prog      Date        Description
' STB       28/11/2001  SYS2912 - Removed the frmMainUserDetails screen and
'                       ported its functionality to this screen.
' STB       21/01/02    SYS2957 Supervisor Security Enhancement.
' STB       24/04/02    SYS4430 Forename and Surname are now mandatory.
' STB       02/04/02    SYS4496 If a tab contains invalid data, bring that tab
'                       to the foreground.
' STB       15/05/02    SYS4906 Added the Mortgage Admin tab.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS
' AW        11/07/02    BMIDS00177  Removed Mortgage Admin tab and calls to external systems
' HMA       06/08/03    BMIDS623    Changed SetCurrentUnit function.
' HMA       04/02/04    BMIDS678    Added Credit Check Access radio buttons to frame.
' HMA       09/03/04    BMIDS678    Correction in cmdAddUserRole_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS
' LD        20/10/05    MAR242      NT Authentication
' GHun      21/10/2005  MAR264      Added Workgroup User radio buttons
' GHun      16/11/2005  MAR312      Extra Workgroup User requirements
' RF        18/01/2006  MAR1000     Error on updating user when authentication mode is
'                                   Windows Authentication - enable change password for
'                                   system user.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM
' TW        05/02/2007  EP2_798 - Broker registration can be blocked as UserID corresponds to that created in Supervisor
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The current record's user ID.
Public m_sUserID As String

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'A collection of tab-handler objects.
Private m_colForms As Collection

'A status indicator to the forms caller.
Private m_ReturnCode As MSGReturnCode

'Underlying table object for the user record.
Private m_clsTableAccess As TableAccess

'The currently active tab.
Private m_enumActiveTab As OmigaUserTabs

Private blnAddNewRecord As Boolean
Private m_bUserRoleMode  As Boolean
Private m_clsUserRoleTable As UserRoleTable
Private m_nCurrentIndex As Long

'MAR312 GHun
Private m_clsUserCompetencies As UserCompetency
Private m_blnLoading As Boolean
'MAR312 End

Private Const USER_EMERGENCY_PASSWORD As Long = 10

' RF 18/01/2006 MAR1000 Start
Public Enum SECURITY_CREDENTIALS_TYPE ' authentication mode
    WindowsAuthentication = 0
    Omiga = 1
End Enum
Private m_eSecurityCredentialsType As SECURITY_CREDENTIALS_TYPE

Public Function GetSecurityCredentialsType() As SECURITY_CREDENTIALS_TYPE
    GetSecurityCredentialsType = m_eSecurityCredentialsType
End Function

Private Sub SetSecurityCredentialsType()

    Dim strSecurityCredentialsType As String
    strSecurityCredentialsType = g_clsGlobalParameter.FindString("SecurityCredentialsType")
    
    If strSecurityCredentialsType = "WindowsAuthentication" Then
        m_eSecurityCredentialsType = WindowsAuthentication
    Else
        m_eSecurityCredentialsType = Omiga
    End If

End Sub

Private Sub cboAccessType_LostFocus()
    ' reset password state in case access type has changed
    SetPasswordState
End Sub

Private Sub SetPasswordState()
On Error GoTo Failed

    If m_eSecurityCredentialsType = WindowsAuthentication And _
        IsSystemUser() = False Then
        
        lblLenderDetails(12).Enabled = False
        txtPassword.Enabled = False
        txtPassword.Mandatory = False
        'If Len(txtPassword.Text) = 0 Then
            ' default the password, so that the value can be updated if Windows Authentication is
            ' switched off
            'txtPassword.Text = "omiga01"
        'End If
        
        txtPasswordConfirm.Enabled = False
        txtPasswordConfirm.Mandatory = False
        If Len(txtPasswordConfirm.Text) = 0 Then
            txtPasswordConfirm.Text = txtPassword.Text
        End If
            
    Else
        
        lblLenderDetails(12).Enabled = True
        txtPassword.Enabled = True
        txtPassword.Mandatory = True
        
        txtPasswordConfirm.Enabled = True
        txtPasswordConfirm.Mandatory = True
        If Len(txtPasswordConfirm.Text) = 0 Then
            txtPasswordConfirm.Text = txtPassword.Text
        End If
    
        If m_blnLoading = True Then
            'Disable the confirm password.
            EnableConfirmPassword False, False
        End If
        
    End If
    
    If m_eSecurityCredentialsType = WindowsAuthentication And _
        m_blnLoading Then
    
            'Disable the emergency password.
            lblLenderDetails(13).Enabled = False
            txtUserDetails(USER_EMERGENCY_PASSWORD).Enabled = False
            txtUserDetails(USER_EMERGENCY_PASSWORD).Mandatory = False
            
    End If

    Exit Sub

Failed:

    g_clsErrorHandling.RaiseError Err.Number, "Error in SetPasswordState: " & Err.DESCRIPTION
    
End Sub

Public Function IsSystemUser() As Boolean
On Error GoTo Failed
    ' check whether the user being edited has UserAccessType with validation id "SU"
    Dim strValidation As String
    
    g_clsFormProcessing.GetComboValidation cboAccessType, "UserAccessType", strValidation
    
    If strValidation = "SU" Then
        IsSystemUser = True
    Else
        IsSystemUser = False
    End If
    
    Exit Function
    
Failed:

    g_clsErrorHandling.RaiseError Err.Number, "Error in IsSystemUser: " & Err.DESCRIPTION

End Function
' RF 18/01/2006 MAR1000 End

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the current state of the form (add or edit).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Set the edit/add state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetActiveTab
' Description   : Stores the active tab and can select it on the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetActiveTab(ByVal enumTab As OmigaUserTabs, Optional ByVal bSetNow As Boolean = False)
    
    On Error GoTo Failed

    m_enumActiveTab = enumTab

    If bSetNow Then
        SSTab1.Tab = m_enumActiveTab - 1
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboAccessType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAccessType_Validate(Cancel As Boolean)
    Cancel = Not cboAccessType.ValidateData()
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboTitle_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboTitle_Validate(Cancel As Boolean)
    Cancel = Not cboTitle.ValidateData()
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboWorkingHours_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboWorkingHours_Validate(Cancel As Boolean)
    Cancel = Not cboWorkingHours.ValidateData()
End Sub

Private Sub cmdAddUserRole_Click()

    On Error GoTo Failed
        
    If frmEditUser.txtUserDetails(0).Text <> "" Then
        m_sUserID = frmEditUser.txtUserDetails(0).Text
        LoadUserRoles
    Else
        g_clsErrorHandling.DisplayError "Enter the User Details first"
        Me.SSTab1.Tab = OmigaUserTabs.UserDetails - 1   'MAR312 GHun
        frmEditUser.txtUserDetails(0).SetFocus
        Exit Sub
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub cmdDeleteUserRole_Click()
        
    On Error GoTo Failed
    Dim colValues As Collection
    
    If Not frmEditUser.lvUserRoles.SelectedItem Is Nothing Then
        If MsgBox("Delete selected UserRole Rate Set?", vbYesNo + vbQuestion) = vbYes Then
            GetUserRoleKeys colValues
            TableAccess(m_clsUserRoleTable).SetKeyMatchValues colValues
            TableAccess(m_clsUserRoleTable).DeleteRecords
            Dim clsUserRole As New UserRole
            Set clsUserRole = New UserRole
            clsUserRole.Initialise m_bIsEdit
            
            'PopulateListView
            'SetButtonsState
        End If
    End If
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub GetUserRoleKeys(colKeyValues As Collection)
    
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmEditUser.lvUserRoles.SelectedItem.Index
    Set clsPopulateDetails = frmEditUser.lvUserRoles.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colKeyValues = clsPopulateDetails.GetKeyMatchValues
    End If
    
    If colKeyValues Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "User Roles ,unable to obtain Keys"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdEditUserRole_Click()
    On Error GoTo Failed
    
    LoadUserRoles True
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgCompetencies_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgCompetencies_Validate(Cancel As Boolean)
    
    'Only validate if the Competency History tab is shown.
    If Me.SSTab1.Tab + 1 = CompetencyHistory Then
        Cancel = Not dgCompetencies.DoControlValidation()
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgQualifications_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgQualifications_Validate(Cancel As Boolean)

    'Only validate if the Qualifications tab is shown.
    If Me.SSTab1.Tab + 1 = Qualifications Then
        Cancel = Not dgQualifications.DoControlValidation()
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgQualifications_BeforeAdd
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgQualifications_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgQualifications.ValidateRow(nCurrentRow)
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgCompetencies_BeforeAdd
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgCompetencies_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgCompetencies.ValidateRow(nCurrentRow)
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Initialize
' Description   : Set the inital tab to display.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    m_enumActiveTab = UserDetails
    m_nCurrentIndex = -1
End Sub

Private Sub PopulateUserRoles()
    On Error GoTo Failed
    
    Dim sUserID As String
    
    If Len(frmEditUser.m_sUserID) = 0 Then
        g_clsErrorHandling.RaiseError errKeysEmpty, TypeName(Me) & " Populate User Roles"
    End If
    
    g_clsFormProcessing.PopulateFromRecordset frmEditUser.lvUserRoles, m_clsUserRoleTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create any underlying table objects and call-off to any intialisation
'                 routines.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo Failed

    m_blnLoading = True 'MAR312 GHun

    'The default return code is failure (until data is successfully saved).
    SetReturnCode MSGFailure
    
    'Create the underlying table object.
    Set m_clsTableAccess = New OmigaUserTable
    Set m_clsUserRoleTable = New UserRoleTable
    'Only do this is we're editing.
    If m_bIsEdit = True Then
        m_clsTableAccess.SetKeyMatchValues m_colKeys
    End If
    
    'Set the first tab to UserDetails.
    SSTab1_Click 0
    
    'Initialise the tab-handlers.
    InitialiseForms

    'Only do this if we're editing.
    If m_bIsEdit = True Then
        PopulateScreenFields
    End If
    
    ' RF 18/01/2006 MAR1000 Start
    SetSecurityCredentialsType
    
    SetPasswordState
    ' RF 18/01/2006 MAR1000 End
    
    SetUserRoleButtonState
    
    m_blnLoading = False    'MAR312 GHun
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Private Sub SetUserRoleButtonState(Optional bEnable As Boolean = True)
On Error GoTo Failed
    
    Dim bEnableEdit As Boolean
    
    cmdAddUserRole.Enabled = True
    
    Dim lstSelection As ListItem
    Set lstSelection = Me.lvUserRoles.SelectedItem
    
    If Not lstSelection Is Nothing Then
        bEnableEdit = True And bEnable
    Else
        bEnableEdit = False
    End If
    
    cmdEditUserRole.Enabled = bEnableEdit
    'cmdDeleteUserRole.Enabled = bEnableEdit
    cmdDeleteUserRole.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseForms
' Description   : Create and initialise all tab-handler classes for this form. This will
'                 involve loading data into table objects or creating blank records where
'                 appropriate. Also, setup any screen-control state which is common to both
'                 an add AND an edit state (state-specific is done elsewhere).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseForms()
        
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    Dim clsUserRole As New UserRole
    Dim clsOmigaUser As New OmigaUser
    Dim clsUserAccess As New UserAccess
    Dim clsUserQualification As New UserQualification
    'Dim clsUserAdmin As New UserMortgageAdmin
    
On Error GoTo Failed

    'Create a collection to hold our tab-handlers.
    Set m_colForms = New Collection
    Set m_clsUserCompetencies = New UserCompetency  'MAR312 GHun

    'Add each tab-handler into the collection.
    m_colForms.Add clsOmigaUser
    m_colForms.Add clsUserQualification
    m_colForms.Add m_clsUserCompetencies    'MAR312 GHun
    m_colForms.Add clsUserRole
    m_colForms.Add clsUserAccess
    
    'TODO: SYS4069 - Re-enable this tab when Optimus is up to date.
    'SSTab1.TabEnabled(5) = g_bODISystemPresent
    'SSTab1.TabEnabled(5) = False
    
    'First populate the User Details.
    clsOmigaUser.SetTableClass m_clsTableAccess
    
    ' The LenderDetails class contains the methods to write to the Lender Details tab.

    ' Need to populate the Lender Details tab first so we can get the keys from it for the
    ' rest of the tabs - the ones that need keys, that is.
    clsOmigaUser.Initialise m_bIsEdit
        
    'Store the UserID at module-level. If adding, this will be an empty string
    'until we're ready to save.
    m_sUserID = clsOmigaUser.GetUserID()
        
    If m_bIsEdit = True Then
        ' Now, User Qualifications
        clsUserQualification.SetUserID m_sUserID
    
        ' Now, User Competencies
        m_clsUserCompetencies.SetUserID m_sUserID
    
        ' Now, User Role
        clsUserRole.SetUserID m_sUserID
        
        ' Now, User Access
        clsUserAccess.SetUserID m_sUserID
        
        ' Now, User Admin.
        'clsUserAdmin.SetUserID m_sUserID
    End If
    
    clsUserQualification.Initialise m_bIsEdit
    m_clsUserCompetencies.Initialise m_bIsEdit
    clsUserRole.Initialise (m_bIsEdit)
    clsUserRole.SetTableAccess m_clsUserRoleTable
    clsUserAccess.Initialise m_bIsEdit
    
    'Show the current tab.
    SSTab1.Tab = m_enumActiveTab - 1
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetUserRoleTableAccess(clsTableAccess As TableAccess)
On Error GoTo Failed
    
    Dim sFunctionName As String

    sFunctionName = "SetTableAccess"

    If Not clsTableAccess Is Nothing Then
        Set m_clsUserRoleTable = clsTableAccess
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, " " & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form, the return status will indicate failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate data and attempt to save. If successful, close the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
On Error GoTo Failed

    'Ensure all mandatory fields have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    'If they have then proceed.
    If bRet = True Then
        'Validate all captured input.
        bRet = ValidateScreenData()
        
        'If the data was valid.
        If bRet = True Then
            'Save the data.
            SaveScreenData
            
            'Ensure the record is flagged for promotion.
            SaveChangeRequest
            
            'Set the return status to success.
            SetReturnCode
            
            'Hide the form so control returns to its opener.
            Hide
            
            'If the user has just edited their own rights, refresh the treeview.
            If g_sSupervisorUser = m_sUserID Then
                g_clsMainSupport.ShowSupervisorObjects
            End If
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Ensure the record is falgged for promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    
On Error GoTo Failed
    
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Iterates through each tab-handler and asks it to validate the controls which
'                 are relevant to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    Dim nThisTab As Integer
    Dim nTabCount As Integer

On Error GoTo Err_Handler

    'Initialise the loop variable(s).
    nTabCount = m_colForms.Count
    nThisTab = 1
    bRet = True
    
    'Whilst we haven't exceed the tab count and no invalid data has been found.
    While (nThisTab <= nTabCount) And (bRet = True)
        'Request that the tab-handler validates its data.
        bRet = m_colForms(nThisTab).ValidateScreenData()
    
        'Increment the counter if valid, otherwise select the tab which contains
        'invalid data.
        If bRet = True Then
            nThisTab = nThisTab + 1
        Else
            Me.SSTab1.Tab = nThisTab - 1
        End If
    Wend
    
    'Return true for valid data or false if invalid data has been found.
    ValidateScreenData = bRet

    Exit Function
    
Err_Handler:
    g_clsErrorHandling.SaveError
    
    'SYS4496 Bring the invalid tab to the front.
    Me.SSTab1.Tab = (nThisTab - 1)
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Ask each tab-handler to setup any specific controls for an add state. Common
'                 control state will be done in the .Initialse routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    
    Dim oForm As Object
    
    'Broker the call onto each tab-handler.
    For Each oForm In m_colForms
        oForm.SetAddState
    Next oForm
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Ask each tab-handler to setup any specific controls for an edit state.
'                 Common control state will be done in the .Initialse routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
    Dim oForm As Object
    
    'Broker the call onto each tab-handler.
    For Each oForm In m_colForms
        oForm.SetEditState
    Next oForm
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Ask each tab-handler to populate its relevant controls with any underlying
'                 data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    
    On Error GoTo Failed

    'Initialise the loop variable(s).
    nTabCount = m_colForms.Count
    
    'Ask each tab-handler to populate the controls.
    For nThisTab = 1 To nTabCount
        m_colForms(nThisTab).SetScreenFields
    Next

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Broker the call onto each tab-handler. Saving its data to the underlying
'                 table object(s).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    
On Error GoTo Failed

    'Initialise the loop variable(s).
    nTabCount = m_colForms.Count
    nThisTab = 2
    
    'Explicitly save the first tab so that we can retrieve a UserID.
    m_colForms(UserDetails).SaveScreenData
    m_colForms(UserDetails).DoUpdates
    m_sUserID = m_colForms(UserDetails).GetUserID
    
    'Ask each tab to save its data into the table object and then update any
    'relevant table object(s).
    While nThisTab <= nTabCount
        'Set the user id to the other tabs. If we're adding, this won't have
        'been done yet.
        m_colForms(nThisTab).SetUserID m_sUserID
        
        'Copy the screen controls into the table objects.
        m_colForms(nThisTab).SaveScreenData
        
        'Update the underlying table objects.
        m_colForms(nThisTab).DoUpdates
        
        nThisTab = nThisTab + 1
    Wend
    
    Exit Sub
    
Failed:
    Dim strMsg As String
    strMsg = "Error in SaveScreenData for tab " & CStr(nThisTab) & ": "
    g_clsErrorHandling.RaiseError Err.Number, strMsg & Err.DESCRIPTION
End Sub

Private Sub lvUserRoles_DblClick()
On Error GoTo Failed
    
    LoadUserRoles True
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub LoadUserRoles(Optional bIsEdit As Boolean = False)

    Dim sUnitID As String
    Dim enumReturn As MSGReturnCode
    Dim colExistingUnits As Collection
    'Dim clsUserRoleTable As New UserRoleTable
    
On Error GoTo Failed

    If bIsEdit Then
        SetCurrentUnit
        g_clsMainSupport.SetSelectedItem lvUserRoles, m_nCurrentIndex, GET_SELECTION
    End If

    frmUserRole.SetIsEdit bIsEdit
    frmUserRole.SetTableAccess m_clsUserRoleTable
    
    ' Get list of existing UNITS
    GetExistingUnits colExistingUnits, bIsEdit
    frmUserRole.SetExistingUnits colExistingUnits
    frmUserRole.SetUserID m_sUserID

    frmUserRole.Show vbModal, Me

    enumReturn = frmUserRole.GetReturnCode()

    If enumReturn = MSGSuccess Then
        
        If bIsEdit = False Then
            sUnitID = m_clsUserRoleTable.GetUnitId
            'm_clsUserRoleTable.CropDeletedRoles m_sUserID, sUnitID
        End If

        ' Populate the UserRole List view
         PopulateUserRoles
    Else

    End If

    Unload frmUserRole

    If m_bIsEdit Then
        g_clsMainSupport.SetSelectedItem lvUserRoles, m_nCurrentIndex, SET_SELECTION
    End If

    'SetButtonState

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetExistingUnits(colExistingUnits As Collection, Optional bDoNotIncludeCurrent As Boolean = True)
    
On Error GoTo Failed

    Dim rsUnits As ADODB.Recordset
    Dim rsUnitsClone As ADODB.Recordset
    Dim sUnitID As String

    Set colExistingUnits = New Collection

    If TableAccess(m_clsUserRoleTable).RecordCount() > 0 Then
        ' If we're editing, get the current task ID as we don't want to include it in the list
        If bDoNotIncludeCurrent Then
            Dim sCurrentUnitID  As String

            sCurrentUnitID = m_clsUserRoleTable.GetUnitId()

        End If

        Set rsUnits = TableAccess(m_clsUserRoleTable).GetRecordSet()
        Set rsUnitsClone = rsUnits.Clone()

        If Not rsUnitsClone Is Nothing Then
            TableAccess(m_clsUserRoleTable).SetRecordSet rsUnitsClone

            rsUnitsClone.MoveFirst
            Do While Not rsUnitsClone.EOF
                sUnitID = m_clsUserRoleTable.GetUnitId()

                If sCurrentUnitID <> sUnitID Then
                    colExistingUnits.Add sUnitID
                End If

                rsUnitsClone.MoveNext
            Loop

            rsUnitsClone.Close
            Set rsUnitsClone = Nothing
            TableAccess(m_clsUserRoleTable).SetRecordSet rsUnits

        End If

    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetCurrentUnit()
On Error GoTo Failed
    
'   BMIDS623 Start

'   RF 20/01/06 MAR1000 Removed commented code

'   BMIDS623
'   Old Code (above) always returns the FIRST entry in the UserRole table because the
'   key value is UserID and this is the same for all units.
'
'   New code (below):
'   Select the correct record set for the USERROLE table, depending on the unique
'   UnitID selected in the List View.

    Dim bFound As Boolean
    Dim lstItem As ListItem
    Dim rs As ADODB.Recordset
    Dim sValue As Variant
    Dim sUnit As String
 
    Set lstItem = lvUserRoles.SelectedItem
    sUnit = lstItem.Text

    If TableAccess(m_clsUserRoleTable).RecordCount() > 0 Then
        Set rs = TableAccess(m_clsUserRoleTable).GetRecordSet()

        rs.MoveFirst
        bFound = False

        Do While Not rs.EOF And Not bFound

            sValue = rs("UNITID").Value

            If Not IsNull(sValue) Then
                If sValue = sUnit Then
                    bFound = True
                Else
                    rs.MoveNext
                 End If
            End If

        Loop

        If Not bFound Then
           g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Unit record"
        End If
    End If

'   BMIDS623 End

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub lvUserRoles_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Failed
    
    SetUserRoleButtonState
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR312 GHun
Private Sub optWorkgroupUser_Click(Index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim colLine As Collection
    Dim rs As ADODB.Recordset

    If optWorkgroupUser(0).Value Then
        optCreditCheckAccess(1).Value = True
        optCreditCheckAccess(0).Enabled = False
        optCreditCheckAccess(1).Enabled = False
        lblCreditCheck.Enabled = False
        SSTab1.TabEnabled(OmigaUserTabs.Qualifications - 1) = False
        SSTab1.TabEnabled(OmigaUserTabs.SupervisorAccess - 1) = False
        'Move all the selected security items to the available list
        For intRow = 1 To swpSecurityResources.GetSecondCount
            Set colLine = swpSecurityResources.GetLineSecond(intRow)
            swpSecurityResources.AddLineFirst colLine
        Next
        swpSecurityResources.ClearSecond
        
        m_clsUserCompetencies.UpdateCompentencyCombo True
        'Clear the competency type combo selection
        If Not (m_blnLoading) Then
            Set rs = dgCompetencies.GetBoundRecordSet
            If Not (rs.BOF And rs.EOF) Then
                rs.MoveFirst
                rs.fields(1).Value = ""
                rs.fields(3).Value = ""
            End If
        End If
        
    Else
        optCreditCheckAccess(0).Enabled = True
        optCreditCheckAccess(1).Enabled = True
        lblCreditCheck.Enabled = True
        SSTab1.TabEnabled(OmigaUserTabs.Qualifications - 1) = True
        SSTab1.TabEnabled(OmigaUserTabs.SupervisorAccess - 1) = True
        m_clsUserCompetencies.UpdateCompentencyCombo False
    End If
End Sub
'MAR312 End

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SSTab1_Click
' Description   : Call into a global bugfix routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SSTab1_Click(PreviousTab As Integer)
    SetTabstops Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : timerPasswordEnable_Timer
' Description   : Disables the password input and sets the focus to the confirm input.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub timerPasswordEnable_Timer()
On Error Resume Next ' MAR1000
    timerPasswordEnable.Enabled = False
    txtPasswordConfirm.SetFocus
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtPassword_Validate
' Description   : Validate the password specified and enable the confirm input control if a
'                 valid password was entered.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_Validate(Cancel As Boolean)
    
    'Validate the input (and cancels if invalid).
    Cancel = Not txtPassword.ValidateData()
        
    If Cancel = False Then
        'If a password was specified.
        If Len(txtPassword.Text) > 0 Then
            'Enable the confirm input.
            EnableConfirmPassword True, False
        Else
            'Disable the confirm input and clear it.
            EnableConfirmPassword False, True
        End If
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableConfirmPassword
' Description   : Enables or disables and clears the password confirmation input control. If
'                 it's enabled, a timer is started which will disable the password input control
'                 after a VERY brief delay (not sure why a timer is used...).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableConfirmPassword(Optional ByVal bEnable As Boolean = True, Optional ByVal bClear As Boolean = True)
    
    'Enable/disable both the input and the corresponding label.
    lblPasswordConfirm.Enabled = bEnable
    txtPasswordConfirm.Enabled = bEnable
    
    'Clear the contents of the control if desired.
    If bClear Then
        txtPasswordConfirm.Text = ""
    End If
    
    'If the the confirmation control is visible and has been set to enabled, then
    'start a timer which will disable the password control.
    If txtPasswordConfirm.Visible = True And bEnable = True Then
        timerPasswordEnable.Interval = 1
        timerPasswordEnable.Enabled = True
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtUserDetails_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtUserDetails_Validate(Index As Integer, Cancel As Boolean)
' TW 05/02/2007 EP2_798
Dim strIntroducerUserPrefix As String
    Cancel = Not txtUserDetails(Index).ValidateData()
    
    If Not Cancel And Not m_bIsEdit Then
        strIntroducerUserPrefix = UCase$(g_clsGlobalParameter.FindString("IntroducerUserPrefix"))
        If UCase$(Left$(txtUserDetails(Index).Text, Len(strIntroducerUserPrefix))) = strIntroducerUserPrefix Then
            MsgBox "The prefix '" & strIntroducerUserPrefix & "' is not acceptable for manually created users", vbExclamation
            Cancel = True
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return status of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Returns the sucess/fail status to the form's caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


