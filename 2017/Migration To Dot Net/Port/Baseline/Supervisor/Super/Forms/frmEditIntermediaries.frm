VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIntermediaries 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Intermediary"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmEditIntermediaries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   38
      Top             =   7200
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   37
      Top             =   7200
      Width           =   1230
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6915
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   12197
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Intermediary Details"
      TabPicture(0)   =   "frmEditIntermediaries.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraMain"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOrganisation"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraIndividual"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCorrespondence"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCrossSelling"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdTarget"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdProcFee"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Contact Details"
      TabPicture(1)   =   "frmEditIntermediaries.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Bank Details"
      TabPicture(2)   =   "frmEditIntermediaries.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBankDetails"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraBankDetails 
         Caption         =   "Bank Account"
         Height          =   2175
         Left            =   -74700
         TabIndex        =   73
         Tag             =   "Tab3"
         Top             =   600
         Width           =   9015
         Begin MSGOCX.MSGEditBox txtBankDetails 
            Height          =   315
            Index           =   0
            Left            =   1740
            TabIndex        =   33
            Top             =   360
            Width           =   3075
            _ExtentX        =   5424
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
         Begin MSGOCX.MSGEditBox txtBankDetails 
            Height          =   315
            Index           =   1
            Left            =   1740
            TabIndex        =   34
            Top             =   780
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSGOCX.MSGEditBox txtBankDetails 
            Height          =   315
            Index           =   2
            Left            =   1740
            TabIndex        =   36
            Top             =   1620
            Width           =   7035
            _ExtentX        =   12409
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
            MaxLength       =   55
         End
         Begin MSGOCX.MSGEditBox txtBankDetails 
            Height          =   315
            Index           =   3
            Left            =   1740
            TabIndex        =   35
            Top             =   1200
            Width           =   6195
            _ExtentX        =   10927
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
            MaxLength       =   45
         End
         Begin VB.Label lblBankAccount 
            Caption         =   "Account Name"
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   77
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblBankAccount 
            Caption         =   "Bank Name"
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   76
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label lblBankAccount 
            Caption         =   "Sort Code"
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   75
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblBankAccount 
            Caption         =   "Account Number"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   74
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contact Details"
         Height          =   3255
         Left            =   300
         TabIndex        =   67
         Tag             =   "Tab2"
         Top             =   3480
         Width           =   7935
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   18
            Left            =   5580
            TabIndex        =   30
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   31
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
            TabIndex        =   32
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
            TabIndex        =   83
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
            TabIndex        =   78
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
            TabIndex        =   79
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
            TabIndex        =   80
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
            TabIndex        =   81
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
            TabIndex        =   82
            Top             =   1800
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   13
            Left            =   1680
            TabIndex        =   84
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
            TabIndex        =   85
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
            TabIndex        =   86
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
            TabIndex        =   87
            Top             =   2280
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   17
            Left            =   1680
            TabIndex        =   93
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
         Begin VB.Label lblOtherTitle 
            Caption         =   "Other Title"
            Height          =   255
            Left            =   4500
            TabIndex        =   97
            Top             =   660
            Width           =   975
         End
         Begin VB.Label lblContact 
            Caption         =   "Ext. No."
            Height          =   195
            Index           =   11
            Left            =   6240
            TabIndex        =   92
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblContact 
            Caption         =   "Country Code."
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   91
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Area Code."
            Height          =   255
            Left            =   2880
            TabIndex        =   90
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Telephone Number."
            Height          =   255
            Left            =   4200
            TabIndex        =   89
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label21 
            Caption         =   "Type."
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblContact 
            Caption         =   "E-Mail Address"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   72
            Top             =   2820
            Width           =   1155
         End
         Begin VB.Label lblContact 
            Caption         =   "Contact Type"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   71
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblContact 
            Caption         =   "Contact Title"
            Height          =   255
            Index           =   18
            Left            =   180
            TabIndex        =   70
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lblContact 
            Caption         =   "Surname"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   69
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label lblContact 
            Caption         =   "Forename"
            Height          =   195
            Index           =   1
            Left            =   4500
            TabIndex        =   68
            Top             =   1020
            Width           =   1035
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Address"
         Height          =   2895
         Left            =   300
         TabIndex        =   58
         Tag             =   "Tab2"
         Top             =   480
         Width           =   7935
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   1575
            TabIndex        =   27
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
            TabIndex        =   20
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MaxLength       =   8
         End
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   1
            Left            =   1560
            TabIndex        =   21
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
            TabIndex        =   23
            Top             =   960
            Width           =   6135
            _ExtentX        =   10821
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
            MaxLength       =   40
         End
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   4
            Left            =   1560
            TabIndex        =   24
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
            TabIndex        =   25
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
            TabIndex        =   26
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
            Left            =   6420
            TabIndex        =   22
            Top             =   600
            Width           =   1275
            _ExtentX        =   2249
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
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblDetails 
            Caption         =   "Building Name"
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   65
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label lblDetails 
            Caption         =   "No."
            Height          =   195
            Index           =   4
            Left            =   5940
            TabIndex        =   64
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblDetails 
            Caption         =   "Street"
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   63
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label lblDetails 
            Caption         =   "District"
            Height          =   195
            Index           =   6
            Left            =   300
            TabIndex        =   62
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblDetails 
            Caption         =   "Town"
            Height          =   195
            Index           =   7
            Left            =   300
            TabIndex        =   61
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblCounty 
            Caption         =   "County"
            Height          =   195
            Index           =   8
            Left            =   300
            TabIndex        =   60
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label lblDetails 
            Caption         =   "Country"
            Height          =   195
            Index           =   9
            Left            =   300
            TabIndex        =   59
            Top             =   2520
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdProcFee 
         Caption         =   "&Procuration Fee"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70140
         TabIndex        =   19
         Top             =   6300
         Width           =   1395
      End
      Begin VB.CommandButton cmdTarget 
         Caption         =   "&Target"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71700
         TabIndex        =   18
         Top             =   6300
         Width           =   1395
      End
      Begin VB.CommandButton cmdCrossSelling 
         Caption         =   "C&ross Selling"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73200
         TabIndex        =   17
         Top             =   6300
         Width           =   1395
      End
      Begin VB.CommandButton cmdCorrespondence 
         Caption         =   "Corr&espondence"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74640
         TabIndex        =   16
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Frame fraIndividual 
         Caption         =   "Individual"
         Height          =   2535
         Left            =   -74700
         TabIndex        =   41
         Tag             =   "Tab1"
         Top             =   3360
         Width           =   4755
         Begin MSGOCX.MSGComboBox cboTitle 
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   1140
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
         Begin MSGOCX.MSGEditBox txtForename 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Top             =   300
            Width           =   3375
            _ExtentX        =   5953
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
            MaxLength       =   40
         End
         Begin MSGOCX.MSGEditBox txtSurname 
            Height          =   315
            Left            =   1200
            TabIndex        =   12
            Top             =   720
            Width           =   3375
            _ExtentX        =   5953
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
            MaxLength       =   40
         End
         Begin MSGOCX.MSGComboBox cboKeyDevArea 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   1560
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
            Text            =   ""
         End
         Begin MSGOCX.MSGComboBox cboIntermediaryStatus 
            Height          =   315
            Left            =   1200
            TabIndex        =   15
            Top             =   1980
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
            Text            =   ""
         End
         Begin VB.Label lblIndividual 
            Caption         =   "Status Code"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   96
            Top             =   2030
            Width           =   975
         End
         Begin VB.Label lblIndividual 
            Caption         =   "Development Area"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblIndividual 
            Caption         =   "Title"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label lblIndividual 
            Caption         =   "Surname"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblIndividual 
            Caption         =   "Forename"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraOrganisation 
         Caption         =   "Organisation"
         Height          =   1335
         Left            =   -69780
         TabIndex        =   40
         Tag             =   "Tab1"
         Top             =   3360
         Width           =   4275
         Begin MSGOCX.MSGEditBox txtOrgName 
            Height          =   315
            Left            =   1020
            TabIndex        =   9
            Top             =   300
            Width           =   3075
            _ExtentX        =   5424
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
            MaxLength       =   45
         End
         Begin MSGOCX.MSGEditBox txtCommisionNo 
            Height          =   315
            Left            =   1020
            TabIndex        =   10
            Top             =   780
            Width           =   2295
            _ExtentX        =   4048
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
            MaxLength       =   10
         End
         Begin VB.Label lblOrganisation 
            Caption         =   "Commission Number"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lblOrganisation 
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.Frame fraMain 
         Caption         =   "General Details"
         Height          =   2775
         Left            =   -74700
         TabIndex        =   48
         Tag             =   "Tab1"
         Top             =   480
         Width           =   9195
         Begin MSGOCX.MSGDataCombo cboParentIntermediary 
            Height          =   315
            Left            =   1920
            TabIndex        =   4
            Top             =   1140
            Width           =   2355
            _ExtentX        =   4154
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
         End
         Begin MSGOCX.MSGEditBox txtLTVLimit 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            TextType        =   7
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGComboBox cboReportFreq 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   2220
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
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
         Begin MSGOCX.MSGComboBox cboProcFeeMethod 
            Height          =   315
            Left            =   6540
            TabIndex        =   7
            Top             =   1680
            Width           =   2355
            _ExtentX        =   4154
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
         Begin MSGOCX.MSGComboBox cboCommunicationMethod 
            Height          =   315
            Left            =   6540
            TabIndex        =   5
            Top             =   1140
            Width           =   2355
            _ExtentX        =   4154
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
         Begin MSGOCX.MSGEditBox txtActiveTo 
            Height          =   315
            Left            =   6540
            TabIndex        =   3
            Top             =   720
            Width           =   1395
            _ExtentX        =   2461
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
         Begin MSGOCX.MSGEditBox txtActiveFrom 
            Height          =   315
            Left            =   1920
            TabIndex        =   2
            Top             =   720
            Width           =   1395
            _ExtentX        =   2461
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
         Begin MSGOCX.MSGEditBox txtPanelID 
            Height          =   315
            Left            =   6540
            TabIndex        =   1
            Top             =   300
            Width           =   2355
            _ExtentX        =   4154
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
         Begin MSGOCX.MSGComboBox cboType 
            Height          =   315
            Left            =   1920
            TabIndex        =   0
            Top             =   300
            Width           =   2355
            _ExtentX        =   4154
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
         Begin MSGOCX.MSGEditBox txtOtherSystemNo 
            Height          =   315
            Left            =   6540
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   2160
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   556
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
         Begin VB.Label lblOtherSysNo 
            Caption         =   "Optimus Customer No"
            Height          =   255
            Left            =   4740
            TabIndex        =   94
            Top             =   2205
            Width           =   1575
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Communication Method"
            Height          =   195
            Index           =   7
            Left            =   4740
            TabIndex        =   57
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Parent Intermediary"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   56
            Top             =   1200
            Width           =   1635
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Intermediary Type"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   55
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Intermediary Panel ID"
            Height          =   195
            Index           =   1
            Left            =   4740
            TabIndex        =   54
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Active From"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   53
            Top             =   780
            Width           =   1635
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Active To"
            Height          =   195
            Index           =   3
            Left            =   4740
            TabIndex        =   52
            Top             =   780
            Width           =   1635
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Report Frequency"
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   51
            Top             =   2280
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblLTVLimit 
            Caption         =   "Procuration Fee LTV Limit"
            Height          =   375
            Left            =   180
            TabIndex        =   50
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblContactDetails 
            Caption         =   "Procuration Fee Payment Method"
            Height          =   495
            Index           =   5
            Left            =   4740
            TabIndex        =   49
            Top             =   1620
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmEditIntermediaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditIntermediaries
' Description   :   Form which allows the adding/editing of intermediaries.
' Change history
' Prog      Date        Description
' AA        10/04/01    Created
' DJP       27/06/01    SQL Server por
' STB       08/11/01    Added telephone table functionality.
' STB       15/11/01    SYS2550 SQL Server support.
' STB       26/11/01    SYS2912 SQL Server locking issue.
' AW        08/01/02    SYS3560: Added 'Optimus Customer No'
' SDS       25/01/02    SYS3770: Field validation for Postcode, Street , Contact telephone
' SDS       25/01/02    SYS3766: Enabled 'Optimus Customer No' field
' STB       20/02/02    SYS4117: Screen locking fixed - Enabled was set to False.
' STB       20/02/02    SYS4118: 'Optimus Customer No' field made back into a read-
'                       only, non-mandatory field.
' STB       20/02/02    SYS3768: LTV Limit label also enabled/disabled.
' STB       22/04/02    SYS4397: Default state for proc fees button is enabled.
' STB       02/04/02    SYS4496 If a tab contains invalid data, bring that tab
'                       to the foreground.
' CL        24/05/02    SYS4699 Client specific contact details class hasn't been utilised by Intermedaries
' CL        28/05/02    SYS4766 Merge MSMS & CORE
' DB        06/01/03    SYS5457 Modify promotion of Intermediaries
' DB        07/01/03    SYS5550 Removed mandatory flag for the second telephone no: on the contact details
'
' EPSOM CHANGES
' PB        25/04/06    EP374   Removed mandatory flag from Street field
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Control indexes - also defined in tab handler...
Private Const POST_CODE     As Long = 0
Private Const COUNTRY_CODE1 As Long = 9
Private Const AREA_CODE1    As Long = 10
Private Const TELEPHONE1    As Long = 11
Private Const EXTENSION1    As Long = 12
Private Const COUNTRY_CODE2 As Long = 13
Private Const AREA_CODE2    As Long = 14
Private Const TELEPHONE2    As Long = 15
Private Const EXTENSION2    As Long = 16

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of tab-handler objects.
Private m_colForms As Collection

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'Underlying table object for the record.
Private m_clsIntermediaryTable As IntermediaryTable

'The parent Intermediaries GUID - set when adding.
Private m_vParentGUID As Variant

'The type of intermediary (Individual, Company, etc.)
Private m_uIntermediaryType As IntermediaryTypeEnum

'These are table-objects who don't have a tab-handler. They are modified by
'popup screens (which are loaded from this screen) whilst their lifetime
'(creation, deletion and data loading/adding) is controlled within this form.

'Tables for the correspondence popup.
Private m_clsCorrespondenceTable As IntCorrespondenceTable
Private m_clsReportTable As IntermediaryReportDaysTable

'Table for the cross selling popup.
Private m_clsCrossSellingTable As IntCrossSellingTable

'Table for the target details popup.
Private m_clsTargetTable As IntermediaryTargetTable

'Tables for the procuration fee popup(s).
Private m_clsProcFeeSplitTable As IntProcFeeSplitTable
Private m_clsProcFeeTypeTable As IntermediaryProcFeeTable
Private m_clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetReturnCode
' Description   :   Sets the return code for the form for any calling method to check. Defaults
'                   to MSG_SUCCESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal uReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = uReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboType1_Click
' Description   : Enable Contact telephone extension text box if type is work.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboType1_Click()
    
    If cboType1.SelText = "Work" Then
        txtContactDetails(EXTENSION1).Enabled = True
    Else
        txtContactDetails(EXTENSION1).Enabled = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboType2_Click
' Description   : Enable Contact telephone extension text box if type is work.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboType2_Click()
    
    If cboType2.SelText = "Work" Then
        txtContactDetails(EXTENSION2).Enabled = True
    Else
        txtContactDetails(EXTENSION2).Enabled = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCorrespondence_Click
' Description   : Show the correspondence form modally.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCorrespondence_Click()
    
    Dim frmCorrespondence As frmEditIntReportDetails
    
    On Error GoTo Failed
    
    'Create a correspondance popup.
    Set frmCorrespondence = New frmEditIntReportDetails
    
    'Set the keys collection on the form.
    frmCorrespondence.SetKeys m_colKeys
    
    'Set the add/edit state on the form.
    frmCorrespondence.SetIsEdit m_bIsEdit
    
    'Associate the table(s) with the screen.
    frmCorrespondence.SetDailyReport m_clsReportTable
    frmCorrespondence.SetIntermediary m_clsIntermediaryTable
    frmCorrespondence.SetCorrespondence m_clsCorrespondenceTable
    
    'Show the form.
    frmCorrespondence.Show vbModal
    
    'Unload it.
    Unload frmCorrespondence
    Set frmCorrespondence = Nothing
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCrossSelling_Click
' Description   : Show the cross-selling form modally.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCrossSelling_Click()
    
    Dim frmCrossSelling As frmEditIntCrossSelling
    
    On Error GoTo Failed
    
    'Create a cross-selling popup.
    Set frmCrossSelling = New frmEditIntCrossSelling
    
    'Associate the table with the form.
    frmCrossSelling.SetCrossSelling m_clsCrossSellingTable
    
    'Set the keys collection against the form.
    frmCrossSelling.SetKeys m_colKeys
    
    'Set the add/edit state against the form.
    frmCrossSelling.SetIsEdit m_bIsEdit
    
    'Show the form.
    frmCrossSelling.Show vbModal
    
    'Unload it.
    Unload frmCrossSelling
    Set frmCrossSelling = Nothing
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the record, closing the form if everything is okay.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bValid As Boolean
    Dim sNodeText As String
    Dim clsIntermediary As Intermediary
     
    On Error GoTo Failed

    'Create an Intermediary Helper-class.
    Set clsIntermediary = New Intermediary
                
    'Set the caption to be displayed in the tree node.
    If m_uIntermediaryType = IndividualType Then
        sNodeText = txtForename.Text & " " & txtSurname.Text
    Else
        sNodeText = txtOrgName.Text
    End If
    
    'If either the country code, area code, or telephone number have been specified,
    'Then make the number-type combo mandatory.
    If Trim(txtContactDetails(COUNTRY_CODE1).Text) <> "" Or Trim(txtContactDetails(AREA_CODE1).Text) <> "" Or Trim(txtContactDetails(TELEPHONE1).Text) <> "" Then
        cboType1.Mandatory = True
    Else
        cboType1.Mandatory = False
        cboType1.Text = COMBO_NONE
    End If
        
    'If either the country code, area code, or telephone number have been specified,
    'Then make the number-type combo mandatory.
    If Trim(txtContactDetails(COUNTRY_CODE2).Text) <> "" Or Trim(txtContactDetails(AREA_CODE2).Text) <> "" Or Trim(txtContactDetails(TELEPHONE2).Text) <> "" Then
        cboType2.Mandatory = True
    Else
        cboType2.Mandatory = False
        cboType2.Text = COMBO_NONE
    End If
    
    'Ensure all mandatory fields have been populated.
    bValid = g_clsFormProcessing.DoMandatoryProcessing(Me)

    'If they have then proceed.
    If bValid = True Then
        'Validate all captured input.
        bValid = ValidateScreenData()
        
        'If the data was valid.
        If bValid = True Then
            'Save the data.
            SaveScreenData
            
            'Ensure the record is flagged for promotion.
            SaveChangeRequest
                        
            'Ensure any changes are reflected in the main treeview (frmMain).
            clsIntermediary.UpdateNode m_colKeys(INTERMEDIARY_KEY), m_vParentGUID, sNodeText, m_uIntermediaryType, m_bIsEdit
            
            'Set the return status to success.
            SetReturnCode
            
            'Hide the form so control returns to its opener.
            Hide
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdProcFee_Click
' Description   : Show the Procuration Fees dialog (modally).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdProcFee_Click()

    Dim frmProcFee As frmEditProcurationFees

    On Error GoTo Failed

    'Create a correspondance popup.
    Set frmProcFee = New frmEditProcurationFees

    'Set the keys collection on the form.
    frmProcFee.SetKeys m_colKeys

    'Set the add/edit state on the form.
    frmProcFee.SetIsEdit m_bIsEdit

    'Associate the table(s) with the screen.
    frmProcFee.SetProcFeeType m_clsProcFeeTypeTable
    frmProcFee.SetProcFeeSplit m_clsProcFeeSplitTable
    frmProcFee.SetProcFeeSplitByInt m_clsProcFeeSplitByIntTable

    'Show the form.
    frmProcFee.Show vbModal

    'Unload it.
    Unload frmProcFee
    Set frmProcFee = Nothing

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdTarget_Click
' Description   : Show the Targets form (modally).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTarget_Click()

    Dim frmTarget As frmEditIntTargetDetails

    On Error GoTo Failed
    
    'Create a Target details form.
    Set frmTarget = New frmEditIntTargetDetails
    
    'Associate the Target table with the Target details screen.
    frmTarget.SetTarget m_clsTargetTable
    
    'Pass the add/edit state to the form.
    frmTarget.SetIsEdit m_bIsEdit
    
    'Pass the keys collection to the form.
    frmTarget.SetKeys m_colKeys
    
    'Show the form (modally).
    frmTarget.Show vbModal
    
    'Unload it from memory.
    Unload frmTarget
    Set frmTarget = Nothing
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create the underlying table objects, build the tabs/tab-handlers and
'                 setup the controls.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom changes
'
' PB        09/06/06    EP521   Changed lblOrganisation(1) caption 'Commision' to 'Commission'
'
Private Sub Form_Load()
    
    On Error GoTo Failed

    'Create all module-level table objects which are handled by this form.
    'These tables are handled by the form rather than the tab-handlers.
    Set m_clsTargetTable = New IntermediaryTargetTable
    Set m_clsReportTable = New IntermediaryReportDaysTable
    Set m_clsCrossSellingTable = New IntCrossSellingTable
    Set m_clsIntermediaryTable = New IntermediaryTable
    Set m_clsCorrespondenceTable = New IntCorrespondenceTable
    Set m_clsProcFeeTypeTable = New IntermediaryProcFeeTable
    Set m_clsProcFeeSplitTable = New IntProcFeeSplitTable
    Set m_clsProcFeeSplitByIntTable = New IntProcFeeSplitForIntTable
        
    'Set the keys collection against all local tables.
    TableAccess(m_clsIntermediaryTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsCorrespondenceTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsCrossSellingTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsReportTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsTargetTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsProcFeeTypeTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsProcFeeSplitTable).SetKeyMatchValues m_colKeys
    TableAccess(m_clsProcFeeSplitByIntTable).SetKeyMatchValues m_colKeys
        
    If m_bIsEdit Then
        'Populate local tables with data.
        SetEditState
    Else
        'Create GUIDs and blank records where required.
        SetAddState
    End If
    
    'Build the tabs and tab-handlers.
    InitialiseTabs
    
    'Update the controls on the form with the data in the tables.
    PopulateScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Create empty records and generate GUIDs where they are required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
        
    Dim vGUID As Variant
    Dim clsGUID As GuidAssist
        
    On Error GoTo Failed
    
    'We must create a new Intermediary GUID.
    Set clsGUID = New GuidAssist
    vGUID = g_clsSQLAssistSP.GuidStringToByteArray(clsGUID.CreateGUID)
    
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add vGUID
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsIntermediaryTable
    
    'Populate these child tables as empty.
    TableAccess(m_clsCorrespondenceTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsReportTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsCrossSellingTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsTargetTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsProcFeeTypeTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsProcFeeSplitTable).GetTableData POPULATE_EMPTY
    TableAccess(m_clsProcFeeSplitByIntTable).GetTableData POPULATE_EMPTY
        
    'Set the GUID field on the underlying table object.
    m_clsIntermediaryTable.SetIntermediaryGuid vGUID
    
    'Set the Parent GUID.
    m_clsIntermediaryTable.SetParentGUID m_vParentGUID
    
    'Set the Intermediary Type against the table.
    m_clsIntermediaryTable.SetIntermediaryType m_uIntermediaryType
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Load data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    On Error GoTo Failed
    
    'Intermediaries.
    TableAccess(m_clsIntermediaryTable).GetTableData
    
    'Correspondence.
    TableAccess(m_clsCorrespondenceTable).GetTableData
    
    'Dailey Report.
    TableAccess(m_clsReportTable).GetTableData
    
    'Cross-Selling.
    TableAccess(m_clsCrossSellingTable).GetTableData
    
    'Targets.
    TableAccess(m_clsTargetTable).GetTableData
    
    'Procuration Fees.
    TableAccess(m_clsProcFeeTypeTable).GetTableData
    TableAccess(m_clsProcFeeSplitTable).GetTableData
    TableAccess(m_clsProcFeeSplitByIntTable).GetTableData
         
    'Store the parent GUID and IntermediaryType at module-level.
    m_vParentGUID = m_clsIntermediaryTable.GetParentIntermediaryGUID
    m_uIntermediaryType = m_clsIntermediaryTable.GetIntermediaryType
            
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Each tab-handler should now populate the screen from its underlying data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    
    Dim iTab As Integer
    
    On Error GoTo Failed
        
    For iTab = 1 To m_colForms.Count
        m_colForms(iTab).SetScreenFields
    Next
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseTabs
' Description   : Adds all screen clasess to the form collection. A GUID ID will have been
'                 created at this point and a keys collection generated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseTabs()
    
    Dim clsContactDetails As ContactDetails
    
    On Error GoTo Failed
        
    'Create the tab-handler collection.
    Set m_colForms = New Collection
    Set clsContactDetails = New ContactDetailsCS
    
    'Add new instances of each tab-handler into the collection.
    m_colForms.Add New IntermediaryScreenDetails
    m_colForms.Add clsContactDetails
    m_colForms.Add New IntermediaryBankDetails
    
    m_colForms(IntDetails).SetKeys m_colKeys
    m_colForms(IntBankDetails).SetKeys m_colKeys
        
    'The Contact and Address GUIDs will be in the base Intermediary record.
    'Obtain them and set correct keys collections against the contact details
    'tables (contact and address).
    SetContactKeys
    
    'Set the form for each tab class.
    m_colForms(IntDetails).SetForm Me
    m_colForms(IntBankDetails).SetForm Me
       
    m_colForms(IntContactDetails).SetContactDetailsForm Me
    
    'Associate the base table object with the first tab-handler.
    m_colForms(IntDetails).SetIntermediaryTable m_clsIntermediaryTable
    
    'Ensure the first tab-handler knows what type it's dealing with before
    'initialising.
    m_colForms(IntDetails).SetIntermediaryType m_uIntermediaryType
    
    'Inititialise tab-handlers.
    m_colForms(IntDetails).Initialise m_bIsEdit
    m_colForms(IntBankDetails).Initialise m_bIsEdit
    
    m_colForms(IntContactDetails).Initialise m_bIsEdit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the success code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetContactKeys
' Description   : When we're editing, we need to get the Contact and Address GUIDs from the
'                 base Intermediary record and populate the keys collections for the contact
'                 details tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetContactKeys()
    
    On Error GoTo Failed
    
    Dim clsContactDetails As ContactDetails
    Dim colKey As Collection
    Dim vGUID As Variant
        
    'Get the Contact's GUID from the Intermediary record.
    vGUID = m_clsIntermediaryTable.GetContactDetailsGUID
            
    'Assuming we have one, add it into a keys collection and set against the
    'contact tab-handler.
    If Not IsNull(vGUID) Then
        If Len(vGUID) > 0 Then
            Set colKey = New Collection
            colKey.Add vGUID, CONTACT_DETAILS_KEY
            Set clsContactDetails = m_colForms(ContactDetails)
            clsContactDetails.SetContactKeyValues colKey
        End If
    End If
    
    'Get the Address's GUID from the Intermediary record.
    vGUID = m_clsIntermediaryTable.GetAddressGUID
    
    'Assuming we have one, add it into a keys collection and set against the
    'contact tab-handler.
    If Not IsNull(vGUID) Then
        If Len(vGUID) > 0 Then
            Set colKey = New Collection
            colKey.Add vGUID, ADDRESS_KEY
            clsContactDetails.SetAddressKeyValues colKey
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetParentKey
' Description   : Stores the Parent intermediaries GUID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetParentKey(ByVal vParentKey As Variant)
    m_vParentGUID = vParentKey
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
    m_colForms(IntContactDetails).SaveScreenData
    m_colForms(IntContactDetails).DoUpdates
    
    'Now that the contact information has been saved, we'll have GUIDs to assign
    'to the base record intermediary.
    If m_bIsEdit = False Then
        m_clsIntermediaryTable.SetAddressGUID m_colForms(IntContactDetails).GetAddressGUID
        m_clsIntermediaryTable.SetContactGUID m_colForms(IntContactDetails).GetContactDetailsGUID
    End If
    
    'Ask each tab to save its data into the table object and then update any
    'relevant table object(s).
    While nThisTab <= nTabCount
        'Skip the contact details tab.
        If nThisTab <> (IntContactDetails) Then
            m_colForms(nThisTab).SaveScreenData
            m_colForms(nThisTab).DoUpdates
        End If
        
        nThisTab = nThisTab + 1
    Wend
    
    'Ensure all the local tables perform an update.
    TableAccess(m_clsCorrespondenceTable).Update
    TableAccess(m_clsReportTable).Update
    TableAccess(m_clsCrossSellingTable).Update
    TableAccess(m_clsTargetTable).Update
    
    TableAccess(m_clsProcFeeTypeTable).Update
    TableAccess(m_clsProcFeeSplitTable).Update
    TableAccess(m_clsProcFeeSplitByIntTable).Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIntermediaryType
' Description   : Sets the listindex of the type combo to that of the main (calling) screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIntermediaryType(ByVal uIntermediaryType As IntermediaryTypeEnum)
    m_uIntermediaryType = uIntermediaryType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Clear-up object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    
    'Release these object references.
    Set m_colKeys = Nothing
    Set m_clsReportTable = Nothing
    Set m_clsTargetTable = Nothing
    Set m_clsIntermediaryTable = Nothing
    Set m_clsCorrespondenceTable = Nothing
    Set m_clsProcFeeTypeTable = Nothing
    Set m_clsProcFeeSplitTable = Nothing
    Set m_clsProcFeeSplitByIntTable = Nothing

    'These tab-handlers need terminating to clean-up cyclic references.
    m_colForms(IntDetails).Terminate
    m_colForms(IntBankDetails).Terminate
    
    'Contact Details (tab-handler) still holds a reference to this form.
     m_colForms(IntContactDetails).SetContactDetailsForm Nothing
    
    'Release the forms collection reference.
    Set m_colForms = Nothing

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SSTab1_Click
' Description   : Broker the call into a global bugfix routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SSTab1_Click(PreviousTab As Integer)
    SetTabstops Me
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtContactDetails_KeyPress
' Description   : Ensure the postcodes are in upper-case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContactDetails_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = POST_CODE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Iterates through each tab-handler and asks it to validate the
'                 controls which are relevant to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    Dim bValid As Boolean
    Dim iThisTab As Integer
    Dim iTabCount As Integer

    On Error GoTo Err_Handler

    'Initialise the loop variable(s).
    iTabCount = m_colForms.Count
    iThisTab = 1
    bValid = True
    
    'Whilst we haven't exceed the tab count and no invalid data has been found.
    While (iThisTab <= iTabCount) And (bValid = True)
        'Request that the tab-handler validates its data.
        bValid = m_colForms(iThisTab).ValidateScreenData()
    
        'Increment the counter if valid, otherwise select the tab which contains
        'invalid data.
        If bValid = True Then
            iThisTab = iThisTab + 1
        Else
            Me.SSTab1.Tab = iThisTab - 1
        End If
    Wend
    
    'Return true for valid data or false if invalid data has been found.
    ValidateScreenData = bValid
    
    Exit Function
    
Err_Handler:
    g_clsErrorHandling.SaveError
    
    'SYS4496 - Bring the invalid tab to the foreground.
    Me.SSTab1.Tab = (iThisTab - 1)
    Err.Clear
    
    g_clsErrorHandling.RaiseError
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Common Function used to setup a promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()

    Dim colValues As Collection
    Dim clsTableAccess As TableAccess
    Dim sDesc As String
    Dim vGUID As Variant

    On Error GoTo Failed

    vGUID = m_colKeys(INTERMEDIARY_KEY)
    Set clsTableAccess = m_clsIntermediaryTable

    'sDesc = Org name if organisation record or Surname and forename if individual
    If Len(txtOrgName.Text) > 0 Then
        sDesc = txtOrgName.Text
    Else
        sDesc = txtForename.Text & " " & txtSurname.Text
    End If

    Set colValues = New Collection

    colValues.Add vGUID
    clsTableAccess.SetKeyMatchValues colValues
    'DB SYS5457 06/01/03 Don't save the Intermdiary Type
    'g_clsHandleUpdates.SaveChangeRequest clsTableAccess, sDesc & ", " & GetIntermediaryType
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess, sDesc
    'DB End

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryType
' Description   : Return the combo valuename from the given enum/valueid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIntermediaryType() As String
            
    Dim sText As String
    Dim clsCombo As ComboValueTable
    
    On Error GoTo Failed
    
    Set clsCombo = New ComboValueTable
    
    sText = clsCombo.FindComboValue("IntermediaryType", CStr(m_uIntermediaryType))
    GetIntermediaryType = sText
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

