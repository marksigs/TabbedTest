VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditThirdParty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Third Party Details Add/Edit"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   Icon            =   "frmEditThirdParty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   9390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7980
      TabIndex        =   12
      Top             =   7200
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   82
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Third Party Details"
      TabPicture(0)   =   "frmEditThirdParty.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Contact Details"
      TabPicture(1)   =   "frmEditThirdParty.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Legal Rep. Details"
      TabPicture(2)   =   "frmEditThirdParty.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Valuer Details"
      TabPicture(3)   =   "frmEditThirdParty.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Bank Accounts"
      TabPicture(4)   =   "frmEditThirdParty.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6375
         Left            =   -74760
         TabIndex        =   66
         Tag             =   "Tab2"
         Top             =   480
         Width           =   8595
         Begin VB.Frame Frame7 
            Caption         =   "Contact Details"
            Height          =   3315
            Left            =   360
            TabIndex        =   76
            Top             =   3000
            Width           =   7935
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   18
               Left            =   5580
               TabIndex        =   23
               Top             =   600
               Width           =   2055
               _ExtentX        =   3625
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
            End
            Begin MSGOCX.MSGComboBox cboContactType 
               Height          =   315
               Left            =   1560
               TabIndex        =   21
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
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
               ListIndex       =   -1
               Text            =   ""
            End
            Begin MSGOCX.MSGComboBox cboContactTitle 
               Height          =   315
               Left            =   1560
               TabIndex        =   22
               Top             =   600
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
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
               ListIndex       =   -1
               Text            =   ""
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   7
               Left            =   1560
               TabIndex        =   24
               Top             =   960
               Width           =   2715
               _ExtentX        =   4789
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   8
               Left            =   5580
               TabIndex        =   25
               Top             =   960
               Width           =   2055
               _ExtentX        =   3625
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGComboBox cboType2 
               Height          =   315
               Left            =   240
               TabIndex        =   89
               Top             =   2280
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
            Begin MSGOCX.MSGComboBox cboType1 
               Height          =   315
               Left            =   240
               TabIndex        =   84
               Top             =   1800
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
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   9
               Left            =   1680
               TabIndex        =   85
               Top             =   1800
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   10
               Left            =   2880
               TabIndex        =   86
               Top             =   1800
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   11
               Left            =   4200
               TabIndex        =   87
               Top             =   1800
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   12
               Left            =   6240
               TabIndex        =   88
               Top             =   1800
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
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   13
               Left            =   1680
               TabIndex        =   90
               Top             =   2280
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   14
               Left            =   2880
               TabIndex        =   91
               Top             =   2280
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   15
               Left            =   4200
               TabIndex        =   92
               Top             =   2280
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   16
               Left            =   6240
               TabIndex        =   93
               Top             =   2280
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
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   17
               Left            =   1680
               TabIndex        =   99
               Top             =   2760
               Width           =   5955
               _ExtentX        =   10504
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
            Begin VB.Label Label2 
               Caption         =   "OtherTitle"
               Height          =   255
               Left            =   4500
               TabIndex        =   109
               Top             =   660
               Width           =   855
            End
            Begin VB.Label lblContact 
               Caption         =   "Ext. No."
               Height          =   195
               Index           =   11
               Left            =   6240
               TabIndex        =   98
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblContact 
               Caption         =   "Country Code."
               Height          =   255
               Index           =   10
               Left            =   1680
               TabIndex        =   97
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "Area Code."
               Height          =   255
               Left            =   2880
               TabIndex        =   96
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Telephone Number."
               Height          =   255
               Left            =   4200
               TabIndex        =   95
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label21 
               Caption         =   "Type."
               Height          =   255
               Left            =   240
               TabIndex        =   94
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label lblDetails 
               Caption         =   "E-Mail Address"
               Height          =   195
               Index           =   13
               Left            =   180
               TabIndex        =   81
               Top             =   2820
               Width           =   1155
            End
            Begin VB.Label lblLenderDetails 
               Caption         =   "Contact Type"
               Height          =   255
               Index           =   5
               Left            =   180
               TabIndex        =   80
               Top             =   300
               Width           =   1395
            End
            Begin VB.Label lblLenderDetails 
               Caption         =   "Contact Title"
               Height          =   255
               Index           =   18
               Left            =   180
               TabIndex        =   79
               Top             =   660
               Width           =   1395
            End
            Begin VB.Label lblDetails 
               Caption         =   "Surname"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   78
               Top             =   1020
               Width           =   1155
            End
            Begin VB.Label lblDetails 
               Caption         =   "Forename"
               Height          =   195
               Index           =   1
               Left            =   4500
               TabIndex        =   77
               Top             =   1020
               Width           =   1035
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Address"
            Height          =   2895
            Left            =   360
            TabIndex        =   67
            Top             =   0
            Width           =   7935
            Begin MSGOCX.MSGComboBox cboCountry 
               Height          =   315
               Left            =   1575
               TabIndex        =   20
               Top             =   2400
               Width           =   2085
               _ExtentX        =   3678
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
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   0
               Left            =   1560
               TabIndex        =   13
               Top             =   240
               Width           =   1095
               _ExtentX        =   1931
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
               MaxLength       =   8
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   1
               Left            =   1560
               TabIndex        =   14
               Top             =   600
               Width           =   3795
               _ExtentX        =   6694
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   3
               Left            =   1560
               TabIndex        =   16
               Top             =   960
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   4
               Left            =   1560
               TabIndex        =   17
               Top             =   1320
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   5
               Left            =   1560
               TabIndex        =   18
               Top             =   1680
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   6
               Left            =   1560
               TabIndex        =   19
               Top             =   2040
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   2
               Left            =   6900
               TabIndex        =   15
               Top             =   600
               Width           =   795
               _ExtentX        =   1402
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
            Begin VB.Label lblDetails 
               Caption         =   "Postcode"
               Height          =   195
               Index           =   2
               Left            =   300
               TabIndex        =   75
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblDetails 
               Caption         =   "Building Name"
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   74
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblDetails 
               Caption         =   "No."
               Height          =   195
               Index           =   4
               Left            =   6240
               TabIndex        =   73
               Top             =   660
               Width           =   495
            End
            Begin VB.Label lblDetails 
               Caption         =   "Street"
               Height          =   195
               Index           =   5
               Left            =   300
               TabIndex        =   72
               Top             =   1080
               Width           =   915
            End
            Begin VB.Label lblDetails 
               Caption         =   "District"
               Height          =   195
               Index           =   6
               Left            =   300
               TabIndex        =   71
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label lblDetails 
               Caption         =   "Town"
               Height          =   195
               Index           =   7
               Left            =   300
               TabIndex        =   70
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblCounty 
               Caption         =   "County"
               Height          =   195
               Index           =   8
               Left            =   300
               TabIndex        =   69
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label lblDetails 
               Caption         =   "Country"
               Height          =   195
               Index           =   9
               Left            =   300
               TabIndex        =   68
               Top             =   2520
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5475
         Left            =   -74880
         TabIndex        =   64
         Tag             =   "Tab5"
         Top             =   420
         Width           =   8475
         Begin MSGOCX.MSGDataGrid dgBankAccounts 
            Height          =   4035
            Left            =   360
            TabIndex        =   40
            Top             =   600
            Width           =   8115
            _ExtentX        =   14314
            _ExtentY        =   7117
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
         Begin VB.Label Label14 
            Caption         =   "Bank Account Details"
            Height          =   255
            Index           =   5
            Left            =   400
            TabIndex        =   65
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5955
         Left            =   120
         TabIndex        =   43
         Tag             =   "Tab4"
         Top             =   420
         Width           =   8475
         Begin MSGOCX.MSGDataGrid dgValuationType 
            Height          =   2715
            Left            =   1740
            TabIndex        =   39
            Top             =   3240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   4789
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
         Begin MSGOCX.MSGEditBox txtValuer 
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   34
            Top             =   480
            Width           =   1635
            _ExtentX        =   2884
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
         Begin MSGOCX.MSGComboBox cboValuerPaymentMethod 
            Height          =   315
            Left            =   1920
            TabIndex        =   36
            Top             =   840
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
         Begin MSGOCX.MSGEditBox txtValuer 
            Height          =   315
            Index           =   1
            Left            =   5760
            TabIndex        =   35
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGTextMulti txtQualifications 
            Height          =   1395
            Left            =   1920
            TabIndex        =   38
            Top             =   1740
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   2461
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
         Begin MSGOCX.MSGComboBox cboValuerType 
            Height          =   315
            Left            =   1920
            TabIndex        =   37
            Top             =   1260
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
         Begin VB.Label lblValuerType 
            Caption         =   "Valuer Type"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label14 
            Caption         =   "Valuation Type"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   63
            Top             =   3540
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Qualifications"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   62
            Top             =   1740
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Associated Panel ID"
            Height          =   255
            Left            =   4080
            TabIndex        =   61
            Top             =   540
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   900
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Panel ID"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   540
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5115
         Left            =   -74880
         TabIndex        =   42
         Tag             =   "Tab3"
         Top             =   600
         Width           =   8835
         Begin VB.CommandButton cmdAllocatePanelId 
            Caption         =   "&Allocate Panel Id"
            Height          =   375
            Left            =   4260
            TabIndex        =   102
            Top             =   300
            Width           =   1515
         End
         Begin VB.CheckBox chkTempAppointment 
            Alignment       =   1  'Right Justify
            Caption         =   "Temporary Appointment Indicator"
            Height          =   435
            Left            =   420
            TabIndex        =   33
            Top             =   4320
            Width           =   2055
         End
         Begin MSGOCX.MSGTextMulti txtSeniorPartnerDetails 
            Height          =   1395
            Left            =   2100
            TabIndex        =   29
            Top             =   1860
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2461
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
         Begin MSGOCX.MSGEditBox txtLegalRep 
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   26
            Top             =   300
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            Enabled         =   0   'False
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
         Begin MSGOCX.MSGComboBox cboLegalRepType 
            Height          =   315
            Left            =   6060
            TabIndex        =   27
            Top             =   1140
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
         Begin MSGOCX.MSGEditBox txtLegalRep 
            Height          =   315
            Index           =   2
            Left            =   2100
            TabIndex        =   28
            Top             =   1140
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            TextType        =   6
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
         Begin MSGOCX.MSGEditBox txtLegalRep 
            Height          =   315
            Index           =   3
            Left            =   2100
            TabIndex        =   30
            Top             =   3420
            Width           =   1035
            _ExtentX        =   1826
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
         Begin MSGOCX.MSGComboBox cboCountryOfPractice 
            Height          =   315
            Left            =   2100
            TabIndex        =   31
            Top             =   3900
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
         Begin MSGOCX.MSGComboBox cboCountryOfOrigin 
            Height          =   315
            Left            =   6060
            TabIndex        =   32
            Top             =   3420
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
         Begin MSGOCX.MSGComboBox cboLegalRepStatus 
            Height          =   315
            Left            =   6060
            TabIndex        =   101
            Top             =   3900
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
         Begin MSGOCX.MSGComboBox cboLegalPaymentMethod 
            Height          =   315
            Left            =   2100
            TabIndex        =   103
            Top             =   780
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
         Begin MSGOCX.MSGEditBox txtLegalRep 
            Height          =   315
            Index           =   1
            Left            =   6060
            TabIndex        =   104
            Top             =   780
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            TextType        =   6
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
         Begin VB.Label Label13 
            Caption         =   "Legal Rep Type"
            Height          =   255
            Left            =   4380
            TabIndex        =   108
            Top             =   1200
            Width           =   1515
         End
         Begin VB.Label Label1 
            Caption         =   "Legal Rep Status"
            Height          =   255
            Left            =   4380
            TabIndex        =   107
            Top             =   3960
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Associated Panel ID"
            Height          =   255
            Left            =   4380
            TabIndex        =   106
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Country of Origin"
            Height          =   255
            Left            =   4380
            TabIndex        =   105
            Top             =   3480
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Country of Practice"
            Height          =   255
            Left            =   480
            TabIndex        =   58
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Indemnity Insurance Expiry Date"
            Height          =   435
            Left            =   480
            TabIndex        =   57
            Top             =   3420
            Width           =   1500
         End
         Begin VB.Label Label14 
            Caption         =   "Senior Partner Details"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   56
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Number of Partners"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   55
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   420
            TabIndex        =   54
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Panel ID"
            Height          =   255
            Left            =   420
            TabIndex        =   53
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6195
         Left            =   -74820
         TabIndex        =   41
         Tag             =   "Tab1"
         Top             =   360
         Width           =   8475
         Begin VB.CheckBox chkHeadOffice 
            Alignment       =   1  'Right Justify
            Caption         =   "Head Office Indicator"
            Height          =   255
            Left            =   420
            TabIndex        =   2
            Top             =   1080
            Width           =   2115
         End
         Begin MSGOCX.MSGTextMulti txtNotes 
            Height          =   1455
            Left            =   2340
            TabIndex        =   10
            Top             =   4320
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
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
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   3
            Left            =   2340
            TabIndex        =   6
            Top             =   2520
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            TextType        =   4
            PromptInclude   =   0   'False
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MaxValue        =   "999999999"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            MaxLength       =   9
         End
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   4
            Top             =   1800
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            Mask            =   "##/##/####"
            TextType        =   1
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
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   0
            Left            =   2340
            TabIndex        =   1
            Top             =   660
            Width           =   5775
            _ExtentX        =   10186
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
         Begin MSGOCX.MSGComboBox cboAddressType 
            Height          =   315
            Left            =   2340
            TabIndex        =   0
            Top             =   300
            Width           =   4275
            _ExtentX        =   7541
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
         Begin MSGOCX.MSGComboBox cboOrgType 
            Height          =   315
            Left            =   2340
            TabIndex        =   3
            Top             =   1440
            Width           =   5775
            _ExtentX        =   10186
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
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   2
            Left            =   2340
            TabIndex        =   5
            Top             =   2160
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
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   4
            Left            =   2340
            TabIndex        =   7
            Top             =   2880
            Width           =   2655
            _ExtentX        =   4683
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
            MaxLength       =   20
         End
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   5
            Left            =   2340
            TabIndex        =   9
            Top             =   3600
            Width           =   2655
            _ExtentX        =   4683
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
            MaxLength       =   8
         End
         Begin MSGOCX.MSGEditBox txtThirdParty 
            Height          =   315
            Index           =   7
            Left            =   2340
            TabIndex        =   8
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
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
            MaxLength       =   50
         End
         Begin VB.Label lblLabel 
            Caption         =   "Branch Name"
            Height          =   255
            Index           =   7
            Left            =   420
            TabIndex        =   100
            Top             =   3300
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Notes"
            Height          =   255
            Index           =   10
            Left            =   420
            TabIndex        =   52
            Top             =   4380
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Bank Sort Code"
            Height          =   255
            Index           =   8
            Left            =   420
            TabIndex        =   51
            Top             =   3660
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "DX Location"
            Height          =   255
            Index           =   6
            Left            =   420
            TabIndex        =   50
            Top             =   2940
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "DX Id"
            Height          =   255
            Index           =   5
            Left            =   420
            TabIndex        =   49
            Top             =   2580
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Name && Address Type"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   48
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Company Name"
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   47
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Organisation Type"
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   46
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Active From"
            Height          =   255
            Index           =   3
            Left            =   420
            TabIndex        =   45
            Top             =   1860
            Width           =   1815
         End
         Begin VB.Label lblLabel 
            Caption         =   "Active To"
            Height          =   255
            Index           =   4
            Left            =   420
            TabIndex        =   44
            Top             =   2220
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmEditThirdParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditThirdParty
' Description   :
'
' Change history
' Prog      Date        Description
' AA        01/03/01    AQR - SYS1821. Add ValuerType Combo
' STB       08/11/01    Added telephone table functionality.
' STB       22/11/01    Amended contact details key's collection to use a constant
'                       and the address details key's collection (SYS2912).
' STB       29/11/01    SYS2912 - Removed the frmMainThirdParty screen and
'                       ported its functionality to this screen.
' AW        08/01/02    Added OtherSystemNumber
' STB       13/02/02    SYS4054 - Added Branchname field.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS
' AW        11/07/02    BMIDS00177   Removed redundant controls
' DB        15/11/2002  BMIDS00901/909 Removed Bank details tab and set dxID to max val of 99999
' SA        18/11/2002  BMIDS00987  Client Specific version of Contact Details should be used to save county
' DB        27/11/2002  BMIDS01096   Changed dxID to allow for max value of 999999999
' DJP       21/02/2002  BM0318 Set return code default
' MV        05/03/2003  BM0370  Changed the DXID property to allow alphanumeric
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MARS specific changes:
'
'GHun       29/07/2005  MARS27  Re-enable bank account tab, but only allow 1 bank account to be created
'MF         02/08/2005  MAR20   Added combo for Legal Rep Status
'TK         30/11/2005  MAR81   ThirdPartydetails
'TK         01/12/2005  MAR761  Amended Form_Load() and cmdOK_Click()
'TK         01/12/2005  MAR763  Amended cmdAllocatePanelId_Click and cmdOK_Click()
'TK         01/12/2005  MAR775  Amended Form_Load()
'JD         05/01/2006  MAR990  only check bankaccount number if a legalrep type.
'--------------------------------------------------------------------------------------------
'EPSOM specific changes:
'
'SAB        10/04/2006  EP363   Updated to make a local call to Bank Wizard
'PB         25/04/2006  EP448   Bank details no longer mandatory
'PB         10/05/2006  EP511   Merge in change from MAR1684
'PB         15/05/2006  EP375   Allow all Supervisor administrators to set the legal rep status
'PB         15/05/2006  EP545   Refresh main list after edit
'PB         16/05/2006  EP548   Problem with the Activate button and editing a Legal Rep
'PB         01/06/2006  EP642   Error when editing valuer after activating
'TW         05/01/2007  EP2_678 Adding new legal rep causes error if you do not add a panel id
'TW         26/01/2007  EP2_1004 - Edit a Valuer throws alert/error - 'A Panel Id must be allocated' - even when the Panel Id is present
'--------------------------------------------------------------------------------------------

'TODO: Each loop in this module should use m_colForms as the base (and not the
'SSTab.Tabs count). The number of tabs will always be 4 (although some may be
'hidden) whilst the number of tab-handlers, could alter...

'Input control indexes (also defined in tab-handlers?).
Private Const POST_CODE       As Long = 0
Private Const COMPANY_NAME    As Long = 0
Private Const BANK_BRANCHNAME As Long = 7

Private Const BANK_VALUEID As Long = 3
'TK 29/11/2005 MAR81
Private Const LEGALREPPANELID As Long = 0
Private m_sLegalRepUserID As String
Private m_sLegalRepStatus As String
Private m_sLastActivatedDate As String
Private m_bBankDetailsValidated As Boolean

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

''The state of Legal Rep Status. True - Activate mode, False - Inactivate mode.
Private m_bIsActivate As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'A collection of tab-handler objects.
Private m_colForms As Collection

'A status indicator to the forms caller.
Private m_uReturnCode As MSGReturnCode

'The third party type (Local, Legal Rep, or Valuer).
Private m_uThirdPartyType As ThirdPartyType

'The currently active tab.
Private m_uActiveTab As ThirdPartyDetailTabs

'TODO: (eventually) these should be removed...
Private m_vDirectoryGUID As Variant
Private m_clsTableAccess As TableAccess
Private m_clsThirdParty As ThirdPartyDetails
Private m_clsContactDetails As ContactDetails


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
Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

'TK 29/11/2005 MAR81
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsActivate
' Description   : Set the Activate state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsActivate(Optional bActivate As Boolean = True)
    m_bIsActivate = bActivate
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLegalRepStatus
' Description   : Set the Legal Rep Status.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLegalRepStatus(Optional sLegalRepStatus As String)
    m_sLegalRepStatus = sLegalRepStatus
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLegalRepUserID
' Description   : Set the Legal Rep UserID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLegalRepUserID(Optional sLegalRepUserID As String)
    m_sLegalRepUserID = sLegalRepUserID
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLegalRepUserID
' Description   : Returns Legal Rep UserID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLegalRepUserID() As String
    GetLegalRepUserID = m_sLegalRepUserID
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLastActivatedDate
' Description   : Set the LastActivatedDate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLastActivatedDate(Optional sLastActivatedDate As String)
    m_sLastActivatedDate = sLastActivatedDate
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboAddressType_Click
' Description   : Updates the visible tabs, based on the selected value and
'                 set the branchname mandatory state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAddressType_Click()

    Dim lValueID As Long
            
    'Get the selected addresses valueid.
    lValueID = cboAddressType.GetExtra(cboAddressType.ListIndex)
    
    'The Branch name field is only mandatory if the selected item is a bank/building society.
    txtThirdParty(BANK_BRANCHNAME).Mandatory = (lValueID = BANK_VALUEID)

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboAddressType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAddressType_Validate(Cancel As Boolean)
            
    On Error GoTo Failed
            
    'Validate the address selection.
    Cancel = Not cboAddressType.ValidateData()
    
    'Update the organisation combo values and state based upon this selection.
    SetOrganisationComboState

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


Friend Sub SetOrganisationComboState()
    
    Dim vTmp As Variant
    Dim bEnable As Boolean
    Dim uThirdParty As ThirdPartyCombo

    'If the address type is bank/building society, enable organisation type and make
    'it mandatory. Otherwise, disable it.
    g_clsFormProcessing.HandleComboExtra cboAddressType, vTmp, GET_CONTROL_VALUE
    
    If Len(vTmp) = 0 Then
        uThirdParty = ThirdPartyInvalid
    Else
        uThirdParty = CInt(vTmp)
    End If
        
    'Convert this type into the private equivalent. uThirdParty will be either -1
    '(for no selection), or any value from the database (1-12, 99). We need to
    'set m_uThirdPartyType to be equivalent.
    Select Case uThirdParty
        Case ThirdPartyValuer
            m_uThirdPartyType = ThirdPartyValuersType
        
        Case ThirdPartyLegalRep
            m_uThirdPartyType = ThirdPartyLegalRepType
        
        Case -1
            m_uThirdPartyType = ThirdPartyInvalid
            
        Case Else
            m_uThirdPartyType = ThirdPartyLocalType
    End Select
        
    'Enable the organisation combo if the addresstype is...
    bEnable = (uThirdParty = ThirdPartyLender)
    
    cboOrgType.Mandatory = bEnable
    cboOrgType.Enabled = bEnable

    'Reset the tabs and tab-handlers based upon the ThirdParty type selected.
    SetTabs
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboContactTitle_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboContactTitle_Validate(Cancel As Boolean)
    Cancel = Not cboContactTitle.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboContactType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboContactType_Validate(Cancel As Boolean)
    Cancel = Not cboContactType.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboCountry_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCountry_Validate(Cancel As Boolean)
    Cancel = Not cboCountry.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboCountryOfOrigin_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCountryOfOrigin_Validate(Cancel As Boolean)
    Cancel = Not cboCountryOfOrigin.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboCountryOfPractice_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCountryOfPractice_Validate(Cancel As Boolean)
    Cancel = Not cboCountryOfPractice.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboLegalPaymentMethod_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLegalPaymentMethod_Validate(Cancel As Boolean)
    Cancel = Not cboLegalPaymentMethod.ValidateData()
End Sub



Private Sub cboLegalRepStatus_Validate(Cancel As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Me.cboLegalRepStatus.SetListTextFromCollection (inactive)

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboLegalRepType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLegalRepType_Validate(Cancel As Boolean)
    Cancel = Not cboLegalRepType.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboOrgType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboOrgType_Validate(Cancel As Boolean)
    Cancel = Not cboOrgType.ValidateData()
End Sub

' PJO 28/11/2005 coded - MAR81
Private Sub cmdAllocatePanelId_Click()
On Error GoTo Failed
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command

    Set conn = New ADODB.Connection ' TK 01/12/2005 MAR763
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETNEXTSEQUENCENUMBER"
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "LegalRepPanelID")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        txtLegalRep(LEGALREPPANELID).Text = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If txtLegalRep(LEGALREPPANELID).Text <> "" Then
        cmdAllocatePanelId.Enabled = False
    End If
Failed:

End Sub

'MAR27 GHun Enable Add button after a delete
Private Sub dgBankAccounts_BeforeDelete(Cancel As Integer)
    dgBankAccounts.AllowAdd = True
End Sub
'MAR27 End

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgBankAccounts_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgBankAccounts_Validate(Cancel As Boolean)
    If Me.SSTab1.Tab = TabBankAccounts Then
        Cancel = Not dgBankAccounts.ValidateRows()
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgBankAccounts_BeforeAdd
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgBankAccounts_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgBankAccounts.ValidateRow(nCurrentRow)
    End If
    
    'MAR27 GHun Disable add button after adding
    dgBankAccounts.AllowAdd = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgValuationType_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgValuationType_Validate(Cancel As Boolean)

    If Me.SSTab1.Tab = TabValuerDetails Then
        Cancel = Not dgValuationType.ValidateRows()
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : dgValuationType_BeforeAdd
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgValuationType_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgValuationType.ValidateRow(nCurrentRow)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Initialize
' Description   : Sets the default active tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    m_uActiveTab = TabThirdParty
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Initialise the tabs/tab-handlers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    Dim vval As Variant
    Dim iCount As Integer
    Dim nSPLegalRepRole As Integer
    Dim g_sUserAuthority As Integer ' TK 01/12/2005 MAR761
    Dim clsUserRoleTable As UserRoleTable

    On Error GoTo Failed
    
    ' DJP BM0318
    SetReturnCode MSGFailure
    
    'Create a name/address directory table object.
    Set m_clsTableAccess = New NameAndAddressDirTable
    
    'If we're editing then specify the colkeys.
    If m_bIsEdit Then
        m_clsTableAccess.SetKeyMatchValues m_colKeys
        ' PB 15/05/2006 EP375 - commented out next line
        'cboLegalRepStatus.Enabled = False ' TK 29/11/2005 MAR81
        ' EP375 End
    End If
       
    'Create and initialise our tab-handlers.
    InitialiseForms
    
    'The tabs which are actually displayed are governed by third-party type.
    SetTabs
            
    If m_bIsEdit Then
        'Only attempt to populate the screen if we're editing.
        PopulateScreenFields
        ' TK 01/12/2005 MAR761
        Set clsUserRoleTable = New UserRoleTable

        clsUserRoleTable.GetUserRolesForUserID g_sSupervisorUser
        g_sUserAuthority = clsUserRoleTable.GetRole

        If m_bIsEdit Then
            'Legal Rep Status
            If m_uThirdPartyType = ThirdPartyLegalRepType Then
                ' EP375 - info
                ' Previously, if details were changed for an active legal rep, kit became inactive.
                ' This is still ok, but we should be allowed to reactivate if required.
                ' EP375 End or info
                If frmMain.lvListView.ListItems.Item(frmMain.lvListView.SelectedItem.Index).SubItems(5) = "Active" Then
                    'TK 01/12/2005 MAR775
                    If Format$(m_sLastActivatedDate, "yyyymmdd") = Format$(Now, "yyyymmdd") Then
                        nSPLegalRepRole = g_clsGlobalParameter.FindAmount("SPLegalRepRole")
                        If g_sUserAuthority >= nSPLegalRepRole Then
                            setBankDetailsValidated True
                            ' PB 15/05/2006 EP375
                            'frmEditThirdParty.cboLegalRepStatus.ListIndex = 1
                            If frmEditThirdParty.cboLegalRepStatus.ListIndex = -1 Then
                                frmEditThirdParty.cboLegalRepStatus.ListIndex = 1
                            End If
                            ' EP375 End
                            cmdAllocatePanelId.Enabled = False
                            m_bIsActivate = False ' TK 01/12/2005 MAR775
                        Else
                            MsgBox "You have insufficient Authority to edit this Active Record"
                        End If
                    Else
                        setBankDetailsValidated True
                        frmEditThirdParty.cboLegalRepStatus.ListIndex = 1
                        ' PB 15/05/2006 EP375 - commented out next line
                        'frmEditThirdParty.cboLegalRepStatus.Enabled = False
                        ' EP375 End
                        cmdAllocatePanelId.Enabled = False
                        m_bIsActivate = False
                    End If
                Else 'Legal rep status inactive
                    setBankDetailsValidated True
                    ' PB 15/05/2006 EP375 - Commented out next line
                    'frmEditThirdParty.cboLegalRepStatus.Enabled = False
                    ' EP375 End
                    cmdAllocatePanelId.Enabled = False
                End If
            End If
        End If
        
        If m_bIsActivate Then  'TK 29/11/2005 MAR81
            frmEditThirdParty.Frame1.Enabled = False
            frmEditThirdParty.Frame2.Enabled = False
            'frmEditThirdParty.Frame3.Enabled = False
            frmEditThirdParty.Frame4.Enabled = False
            frmEditThirdParty.Frame5.Enabled = False
            frmEditThirdParty.Frame6.Enabled = False
            frmEditThirdParty.Frame7.Enabled = False
            For iCount = 0 To frmEditThirdParty.txtContactDetails.Count - 1
                frmEditThirdParty.txtContactDetails.Item(iCount).Enabled = False
            Next iCount
            For iCount = 0 To frmEditThirdParty.txtLegalRep.Count - 1
                frmEditThirdParty.txtLegalRep.Item(iCount).Enabled = False
            Next iCount
            
            frmEditThirdParty.txtNotes.Enabled = False
            frmEditThirdParty.txtQualifications.Enabled = False
            frmEditThirdParty.txtSeniorPartnerDetails.Enabled = False
            For iCount = 0 To frmEditThirdParty.txtThirdParty.Count - 1
                If iCount <> 6 Then
                    frmEditThirdParty.txtThirdParty.Item(iCount).Enabled = False
                End If
            Next iCount
            For iCount = 0 To frmEditThirdParty.txtValuer.Count - 1
                frmEditThirdParty.txtValuer.Item(iCount).Enabled = False
            Next iCount
            frmEditThirdParty.cmdAllocatePanelId.Enabled = False
            frmEditThirdParty.cboCountryOfOrigin.Enabled = False
            frmEditThirdParty.cboLegalRepType.Enabled = False
            frmEditThirdParty.cboLegalPaymentMethod.Enabled = False
            frmEditThirdParty.cboCountryOfPractice.Enabled = False
            frmEditThirdParty.chkTempAppointment.Enabled = False
            frmEditThirdParty.cboLegalRepStatus.Enabled = True ' TK 29/11/2005 MAR81 End
            If frmEditThirdParty.cboLegalRepStatus.ListCount > 0 Then ' PB EP642
                frmEditThirdParty.cboLegalRepStatus.ListIndex = 0 ' PB EP375
            End If ' PB EP642
        End If
    Else
        'Although if we're adding, select the organisation type passed into this form.
        vval = MapThirdPartyComboToType()
        
        'Select the address type chosen from frmMain.
        g_clsFormProcessing.HandleComboExtra cboAddressType, vval, SET_CONTROL_VALUE
    End If
    
    'Select desired tab.
    SSTab1.Tab = m_uActiveTab

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : MapThirdPartyComboType
' Description   : Converts the module-level m_uThirdPartyType enum into a combo equivalent.
'                 If its equal to ThirdPartyLocalType, the <SELECT> will be chosen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MapThirdPartyComboToType() As ThirdPartyCombo

    Dim uReturn As ThirdPartyCombo

    Select Case m_uThirdPartyType
        Case ThirdPartyValuersType
            uReturn = ThirdPartyValuer
        
        Case ThirdPartyLegalRepType
            uReturn = ThirdPartyLegalRep
        
        Case Else
            uReturn = ThirdPartyLender
    End Select

    MapThirdPartyComboToType = uReturn
    
End Function


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
    Dim vDirectoryGUID As Variant
    Dim clsDirectoryTable As NameAndAddressDirTable
    
    On Error GoTo Failed
    
    Set m_colForms = New Collection

    'First, Third Party Details which is always present.
    Set m_clsThirdParty = New ThirdPartyDetails
    m_colForms.Add m_clsThirdParty
        
    
    'Associate the name and address directory table with the thirdparty details tab-handler.
    m_clsThirdParty.SetTableClass m_clsTableAccess
    
    'Inform the tab-handler what type the intermediary is (valuer, legal rep. or other).
    m_clsThirdParty.SetThirdPartyType m_uThirdPartyType
    
    'Initialise the first tab-handler.
    m_clsThirdParty.Initialise m_bIsEdit

    'Second, Contact Details are always present too.
    'BMIDS00987 Instantiate as client specific version
    Set m_clsContactDetails = New ContactDetailsCS
    m_colForms.Add m_clsContactDetails
    m_clsContactDetails.SetContactDetailsForm Me
        
    'Set the contact details address guid to match the directory one.
    SetContactDetailKeys
    
    'Set the address details address guid to match the directory one.
    SetAddressKeys
    
    m_clsContactDetails.Initialise m_bIsEdit
        
    Set clsDirectoryTable = m_clsTableAccess
    m_vDirectoryGUID = clsDirectoryTable.GetDirectoryGUID()
    
    
    
    ' DJP SQL Server port
    If Not IsEmpty(m_vDirectoryGUID) And Not IsNull(m_vDirectoryGUID) Then
        If Len(CStr(m_vDirectoryGUID)) = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "ThirdParty: DirectoryGUID is empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "ThirdParty: DirectoryGUID is empty"
    End If
    
    'The rest of the forms are done in SetTabs.
    
    'Dim nLegalRepRole As Integer
    'nLegalRepRole = g_clsGlobalParameter.FindAmount("SPLegalRepRole")
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddressKeys
' Description   : Sets the address GUID in the contact details table to be the
'                 same as used in the name/address directory record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddressKeys()
    
    Dim vAddressGUID As Variant
    Dim colValues As New Collection

    On Error GoTo Failed
    
    Const NUMBER_OF_KEYS = 1

    vAddressGUID = m_clsThirdParty.GetAddressGUID()
    
    'Ensure the address directory GUID is used in the contact details record.
    If Not IsNull(vAddressGUID) Then
        If Len(vAddressGUID) > 0 Then
            colValues.Add vAddressGUID, ADDRESS_KEY
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Address GUID is empty"
    End If
    
    If colValues.Count = NUMBER_OF_KEYS Then
        m_clsContactDetails.SetAddressKeyValues colValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetContactDetails: AddressGUID must be valid"
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetContactDetailKeys
' Description   : Ensures the contact details GUID is set to that used by the name and address
'                 directory record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetContactDetailKeys()
    
    Dim colValues As New Collection
    Dim vContactDetailsGUID As Variant
        
    On Error GoTo Failed

    Const NUMBER_OF_KEYS = 1

    ' To populate the Contant Details tab we need the Contact Details GUID
    vContactDetailsGUID = m_clsThirdParty.GetContactDetailsGUID()
        
    If Not IsNull(vContactDetailsGUID) Then
        If Len(vContactDetailsGUID) > 0 Then
            colValues.Add vContactDetailsGUID, CONTACT_DETAILS_KEY
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to set contact detail keys - ContactDetailsGUID is empty"
    End If
    
    If colValues.Count = NUMBER_OF_KEYS Then
        m_clsContactDetails.SetContactKeyValues colValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetContactDetails: ContactDetailsGUID must be valid"
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTabs
' Description   : Ensure the only tabs visible are those relevant to the third-party type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetTabs()
    
    On Error GoTo Failed
        
    ' Always visible
    SSTab1.TabVisible(TabThirdParty) = True
    SSTab1.TabVisible(TabContactDetails) = True
    SSTab1.TabVisible(TabBankAccounts) = True

    'Add/remove tabs and tab-handlers appropriately.
    Select Case m_uThirdPartyType
        Case ThirdPartyLocalType
            EnableTab TabLegalRepDetails, False
            EnableTab TabValuerDetails, False
            EnableTab TabBankAccounts, False
                    
        Case ThirdPartyLegalRepType
            EnableTab TabLegalRepDetails, True
            EnableTab TabValuerDetails, False
            EnableTab TabBankAccounts, True
            
        Case ThirdPartyValuersType
            EnableTab TabLegalRepDetails, False
            EnableTab TabValuerDetails, True
            EnableTab TabBankAccounts, True

        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "ThirdParty: Not a valid third party type " & m_uThirdPartyType
    End Select
        
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableTab
' Description   : Hide or shows the specified tab and will add or remove the required tab-
'                 handler class dynamically.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableTab(ByVal uTabType As ThirdPartyDetailTabs, ByVal bVisible As Boolean)

    Dim oTab As Object
    Dim lIndex As Long
    Dim oCurrent As Object
    
    'Set the visibility of the flag.
    SSTab1.TabVisible(uTabType) = bVisible
    
    'Create a dummy object of the same type of the tab-handler we're interested in.
    Select Case uTabType
        Case TabValuerDetails
            Set oTab = New ValuerDetails
        
        Case TabLegalRepDetails
            Set oTab = New LegalRepDetails

        Case TabBankAccounts
            Set oTab = New PanelBankAccountDetails
    End Select
    
    'Always attempt to remove the tab specified.
    'Initialise our loop variable(s).
    lIndex = 1
    
    'If we're disabling the tab, scan the tab-handler collection for an existing
    'instance of the tab-handler. If found, remove it.
    For Each oCurrent In m_colForms
        'If the datatypes are the same, we've found it.
        If TypeName(oCurrent) = TypeName(oTab) Then
            'Remove the tab-handler.
            m_colForms.Remove lIndex
            
            'Escape the loop we're done.
            Exit For
        End If
        
        'Increment our index counter.
        lIndex = lIndex + 1
    Next oCurrent
    
    'If we're adding a tab...
    If bVisible = True Then
        'Set the directory GUID.
        oTab.SetDirectoryGUID m_vDirectoryGUID
        
        'Add the tab-handler into the collection.
        m_colForms.Add oTab
        
        'Initialise the tab-handler.
        oTab.Initialise m_bIsEdit
    End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetActiveTab
' Description   : Stores the active tab and can select it on the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetActiveTab(enumTab As ThirdPartyDetailTabs, Optional bActivateNow As Boolean = False)
    
    On Error GoTo Failed
    
    m_uActiveTab = enumTab
    
    If bActivateNow Then
        SSTab1.Tab = m_uActiveTab
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetThirdPartyType
' Description   : Establishes the type of third party. This is set prior to the Form_Load
'                 event.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetThirdPartyType(ByVal uType As ThirdPartyType)
    m_uThirdPartyType = uType
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : setBankDetailsValidated
' Description   : Set BankDetailsValidated true or false
'                 event.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub setBankDetailsValidated(Optional bBankDetails As Boolean = True)
    m_bBankDetailsValidated = bBankDetails
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form, the return status will indicate failure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    ' PB 16/05/2006 EP548
    m_bIsActivate = False
    m_bIsEdit = False
    ' EP548 End
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate data and attempt to save. If successful, close the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    Dim sSortCode As String
    Dim sBankAccountNumber As String
    Dim xdcBankWizardDoc As FreeThreadedDOMDocument
    Dim xelBankDetails As IXMLDOMElement
    Dim strXMLData As String
    Dim strError As String
    Dim xndDescription As IXMLDOMNode
    Dim iCount As Integer ' EP511

' TW 26/01/2007 EP2_1004
    Dim strPanelId As String
' TW 26/01/2007 EP2_1004 End

    Const strFunctionName As String = "cmdOK_Click"
    
    On Error GoTo Failed

    'Ensure all mandatory fields have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    'If they have then proceed.
    If bRet = True Then
        'Validate all captured input.
        bRet = ValidateScreenData()
        
        If bRet = True Then
            'JD MAR990 only check bank accounts if tab is visible
            If m_uThirdPartyType <> ThirdPartyLocalType Then
' TW 05/01/2007 EP2_678
' TW 26/01/2007 EP2_1004
                Select Case m_uThirdPartyType
                    Case ThirdPartyLegalRepType
                        strPanelId = txtLegalRep(LEGALREPPANELID).Text
                    Case ThirdPartyValuersType
                        strPanelId = txtValuer(0).Text
                End Select
'                If txtLegalRep(LEGALREPPANELID).Text = "" Then
                If strPanelId = "" Then
' TW 26/01/2007 EP2_1004 End
                    MsgBox "A Panel Id must be allocated"
                    bRet = False
' TW 26/01/2007 EP2_1004
'                    frmEditThirdParty.SetActiveTab TabLegalRepDetails, True
                    Select Case m_uThirdPartyType
                        Case ThirdPartyLegalRepType
                            frmEditThirdParty.SetActiveTab TabLegalRepDetails, True
                        Case ThirdPartyValuersType
                            frmEditThirdParty.SetActiveTab TabValuerDetails, True
                    End Select
' TW 26/01/2007 EP2_1004 End
                Else
' TW 05/01/2007 EP2_678 End
                    If frmEditThirdParty.dgBankAccounts.Rows = 0 Then
                        'PB EP448 - Commented out as bank details no longer mandatory
                        'MsgBox "A Bank Sort Code and Bank Account Number must be entered"
                        'bRet = False
                        'frmEditThirdParty.SetActiveTab TabBankAccounts, True
                        'PB EP448 END
                    Else
                        sSortCode = frmEditThirdParty.dgBankAccounts.Columns.Item(2).Text
                        sBankAccountNumber = frmEditThirdParty.dgBankAccounts.Columns.Item(5).Text
                        If sSortCode = "" Or sBankAccountNumber = "" Then
                            MsgBox "A Bank Sort Code and Bank Account Number must be entered"
                            frmEditThirdParty.SetActiveTab TabBankAccounts, True
                        Else
                            'Call bank wizard with BankSortCode and BankAccountNumber
                            strXMLData = ValidateBankDetails(sSortCode, sBankAccountNumber)
                        
                            Set xdcBankWizardDoc = New FreeThreadedDOMDocument
                            xdcBankWizardDoc.async = False         ' wait until XML is fully loaded
                            'xdcBankWizardDoc.setProperty "NewParser", True
                            xdcBankWizardDoc.validateOnParse = False
                            xdcBankWizardDoc.loadXML strXMLData
                        
                            If xdcBankWizardDoc.selectSingleNode("RESPONSE/@TYPE").Text = "BANKWIZARDERROR" Then
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
                                setBankDetailsValidated False
                                bRet = False
                            ElseIf xdcBankWizardDoc.selectSingleNode("RESPONSE/@TYPE").Text = "APPERR" Then
                                ' EP363 - Change the error handling
                                Set xndDescription = xdcBankWizardDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION")
                                If xndDescription Is Nothing Then
                                    MsgBox "Bank Wizard failed to complete the search without returning an error"
                                Else
                                    MsgBox "Bank Wizard error - " + xndDescription.Text
                                End If
                                bRet = False
                            Else
                                'Get BANK details from the new response
                                frmEditThirdParty.dgBankAccounts.Columns(BANK_VALUEID) = xdcBankWizardDoc.selectSingleNode(".//ORGANISATIONNAME/@FULLNAME").Text
                                bRet = True
                            End If
                        End If
                    End If
' TW 05/01/2007 EP2_678
                End If
' TW 05/01/2007 EP2_678 End
            End If
        End If
    
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
            
            ' PB 15/05/2006 EP545 - Refresh list
            frmMain.PopulateListView
            '
        Else
            ' PB EP511
            frmEditThirdParty.Frame5.Enabled = True

            For iCount = 0 To frmEditThirdParty.txtThirdParty.Count - 1
                If iCount <> 6 Then
                    frmEditThirdParty.txtThirdParty.Item(iCount).Enabled = True
                End If
            Next iCount
            ' EP511 End
        End If
    End If
    
    Set xdcBankWizardDoc = Nothing
    Set xelBankDetails = Nothing
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

    On Error GoTo Failed

    Dim xmlDoc As New MSXML2.DOMDocument
    Dim clsOmiga4 As New Omiga4Support
    Dim sResponse As String

    ' 10/04/2006 EP363 Updated to make a local call to Bank Wizard
    xmlDoc.async = False
    xmlDoc.loadXML ("<REQUEST><BANKDETAILS SORTCODE=""" + sSortCode + """ ACCOUNTNUMBER=""" + sBankAccountNumber + """/></REQUEST>")
    
    sResponse = clsOmiga4.RunASP(xmlDoc.xml, "GetBankDetails.asp")

    ValidateBankDetails = sResponse
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Iterates through each tab-handler and asks it to validate the controls which
'                 are relevant to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    
    On Error GoTo Failed

    'Initialise the loop variable(s).
    nTabCount = m_colForms.Count
    nThisTab = 0
    bRet = True
    
    'Whilst we haven't exceed the tab count and no invalid data has been found.
    While nThisTab < nTabCount And bRet = True
        'Only validate visisble tabs.
        If SSTab1.TabVisible(nThisTab) = True Then
            bRet = m_colForms(nThisTab + 1).ValidateScreenData()
        End If
        
        'Increment the counter if valid, otherwise select the tab which contains
        'invalid data.
        If bRet = True Then
            nThisTab = nThisTab + 1
        Else
            Me.SSTab1.Tab = nThisTab
        End If
    Wend

    'Return true for valid data or false if invalid data has been found.
    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.SaveError
    SSTab1.Tab = nThisTab
    g_clsErrorHandling.DisplayError
    ValidateScreenData = False
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
    Next nThisTab

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
    nThisTab = 1
    
    'This form can be used to add records, the ContactDetails records
    'must exist BEFORE the name and address directory record.
    m_colForms(TabContactDetails + 1).SaveScreenData
    m_colForms(TabContactDetails + 1).DoUpdates
    
    'Ask each tab to save its data into the table object and then update any
    'relevant table object(s).
    While nThisTab <= nTabCount
        'Skip the contact details tab.
        If nThisTab <> (TabContactDetails + 1) Then
            m_colForms(nThisTab).SaveScreenData
            m_colForms(nThisTab).DoUpdates
        End If
        
        nThisTab = nThisTab + 1
    Wend
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SSTab1_Click
' Description   : Call into a global bugfix routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SSTab1_Click(PreviousTab As Integer)
    SetTabstops Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtContactDetails_KeyPress
' Description   : Ensure postcodes are always in upper-case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContactDetails_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = POST_CODE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtContactDetails_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContactDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtContactDetails(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtLegalRep_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtLegalRep_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtLegalRep(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtNotes_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtNotes_Validate(Cancel As Boolean)
    Cancel = Not txtNotes.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtQualifications_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtQualifications_Validate(Cancel As Boolean)
    Cancel = Not txtQualifications.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtSeniorPartnerDetails_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtSeniorPartnerDetails_Validate(Cancel As Boolean)
    Cancel = Not txtSeniorPartnerDetails.ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtThirdParty_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtThirdParty_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtThirdParty(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtValuer_Validate
' Description   : Validation routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtValuer_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtValuer(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return status of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Returns the sucess/fail status to the form's caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : Ensure the record is falgged for promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    
    Dim sType As String
    Dim col As New Collection
    Dim sCompanyName As String
    Dim sDescription As String
    
    sCompanyName = txtThirdParty(COMPANY_NAME).Text
    sType = Me.cboAddressType.SelText

    sDescription = sCompanyName + ", " + sType
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess, sDescription

End Sub
