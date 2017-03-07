VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#6.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditUnit 
   Caption         =   "Add/Edit Unit"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmEditUnit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   60
      TabIndex        =   28
      Top             =   120
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Unit Details"
      TabPicture(0)   =   "frmEditUnit.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Address Details"
      TabPicture(1)   =   "frmEditUnit.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contact Details"
      TabPicture(2)   =   "frmEditUnit.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Task Owners"
      TabPicture(3)   =   "frmEditUnit.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dgTaskOwner"
      Tab(3).ControlCount=   1
      Begin MSGOCX.MSGDataGrid dgTaskOwner 
         Height          =   3255
         Left            =   -74520
         TabIndex        =   56
         Top             =   1080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5741
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3195
         Left            =   60
         TabIndex        =   47
         Tag             =   "Tab3"
         Top             =   360
         Width           =   8595
         Begin MSGOCX.MSGComboBox cboContactType 
            Height          =   315
            Left            =   2100
            TabIndex        =   20
            Top             =   540
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
            Index           =   0
            Left            =   5460
            TabIndex        =   21
            Top             =   540
            Width           =   2535
            _ExtentX        =   4471
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
            Index           =   1
            Left            =   2100
            TabIndex        =   23
            Top             =   1260
            Width           =   5895
            _ExtentX        =   10398
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
            Left            =   2100
            TabIndex        =   24
            Top             =   1620
            Width           =   4035
            _ExtentX        =   7117
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
         Begin MSGOCX.MSGEditBox txtContactDetails 
            Height          =   315
            Index           =   3
            Left            =   6960
            TabIndex        =   25
            Top             =   1620
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
            Index           =   4
            Left            =   2100
            TabIndex        =   26
            Top             =   1980
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSGOCX.MSGComboBox cboContactTitle 
            Height          =   315
            Left            =   2100
            TabIndex        =   22
            Top             =   900
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
         Begin Supervisor.MSGTextMulti txtEmailAddress 
            Height          =   315
            Left            =   2100
            TabIndex        =   27
            Top             =   2340
            Width           =   5895
            _ExtentX        =   10398
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
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Type"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   55
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Forename"
            Height          =   255
            Index           =   15
            Left            =   4620
            TabIndex        =   54
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Surname"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   53
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Telephone"
            Height          =   255
            Index           =   17
            Left            =   480
            TabIndex        =   52
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Fax Number"
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   51
            Top             =   2040
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Ext."
            Height          =   255
            Index           =   19
            Left            =   6360
            TabIndex        =   50
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Email Address"
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   49
            Top             =   2400
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Title"
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   48
            Top             =   960
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3195
         Left            =   -74940
         TabIndex        =   38
         Tag             =   "Tab2"
         Top             =   360
         Width           =   8595
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   12
            Top             =   420
            Width           =   1155
            _ExtentX        =   2037
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
            MaxLength       =   8
         End
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   13
            Top             =   780
            Width           =   4335
            _ExtentX        =   7646
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   2
            Left            =   6900
            TabIndex        =   14
            Top             =   780
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   15
            Top             =   1140
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   4
            Left            =   2040
            TabIndex        =   16
            Top             =   1500
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   5
            Left            =   2040
            TabIndex        =   17
            Top             =   1860
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   6
            Left            =   2040
            TabIndex        =   18
            Top             =   2220
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   2040
            TabIndex        =   19
            Top             =   2580
            Width           =   2475
            _ExtentX        =   4366
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
         Begin VB.Label lblLenderDetails 
            Caption         =   "Postcode"
            Height          =   255
            Index           =   5
            Left            =   420
            TabIndex        =   46
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Building Name"
            Height          =   255
            Index           =   6
            Left            =   420
            TabIndex        =   45
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "No."
            Height          =   255
            Index           =   7
            Left            =   6540
            TabIndex        =   44
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Street"
            Height          =   255
            Index           =   8
            Left            =   420
            TabIndex        =   43
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "District"
            Height          =   255
            Index           =   9
            Left            =   420
            TabIndex        =   42
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Town"
            Height          =   255
            Index           =   10
            Left            =   420
            TabIndex        =   41
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "County"
            Height          =   255
            Index           =   11
            Left            =   420
            TabIndex        =   40
            Top             =   2280
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Country"
            Height          =   255
            Index           =   12
            Left            =   420
            TabIndex        =   39
            Top             =   2640
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5115
         Left            =   -74880
         TabIndex        =   29
         Tag             =   "Tab1"
         Top             =   360
         Width           =   8475
         Begin VB.CheckBox chkProcessingUnit 
            Caption         =   "Processing Unit?"
            Height          =   255
            Left            =   5040
            TabIndex        =   9
            Top             =   4620
            Width           =   1935
         End
         Begin VB.CheckBox chkOrganisationUnit 
            Caption         =   "Organisation Unit?"
            Height          =   255
            Left            =   2220
            TabIndex        =   8
            Top             =   4620
            Width           =   1935
         End
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   0
            Left            =   2220
            TabIndex        =   0
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   1
            Left            =   2220
            TabIndex        =   1
            Top             =   720
            Width           =   4515
            _ExtentX        =   7964
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
         Begin MSGOCX.MSGDataCombo cboDepartment 
            Height          =   315
            Left            =   2220
            TabIndex        =   2
            Top             =   1080
            Width           =   2115
            _ExtentX        =   3731
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
            Mandatory       =   -1  'True
         End
         Begin MSGOCX.MSGDataGrid dgRegions 
            Height          =   2115
            Left            =   2040
            TabIndex        =   3
            Top             =   1500
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   3731
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   2
            Left            =   2220
            TabIndex        =   4
            Top             =   3900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Mask            =   "##/##/####"
            Format          =   "c"
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   3
            Left            =   5040
            TabIndex        =   5
            Top             =   3900
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            Mask            =   "##/##/####"
            Format          =   "c"
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   4
            Left            =   2220
            TabIndex        =   6
            Top             =   4260
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            TextType        =   6
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   5
            Left            =   5040
            TabIndex        =   7
            Top             =   4260
            Width           =   1695
            _ExtentX        =   2990
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
         Begin VB.Label lblLenderDetails 
            Caption         =   "DX Location"
            Height          =   195
            Index           =   23
            Left            =   3840
            TabIndex        =   37
            Top             =   4320
            Width           =   1035
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "DX ID"
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   36
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Department"
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   35
            Top             =   1140
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Unit ID"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   34
            Top             =   420
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Unit Name"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   33
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Regions"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   32
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Active From"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   31
            Top             =   3960
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Active To"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   30
            Top             =   3960
            Width           =   1035
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7590
      TabIndex        =   11
      Top             =   5835
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   6270
      TabIndex        =   10
      Top             =   5835
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditUnit
' Description   : Form which allows the user to edit and add unit/task details
'
' Change history
' Prog      Date        Description
' AA        15/01/00    Added TaskOwners Tab and LoadTaskowner function.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Contains the tab index's for each frmEditUnit Tab
Private tabsUnit As UnitTabs

' Unit Tab
Private Const UNIT_ID = 0
Private Const UNIT_NAME = 1
Private Const UNIT_ACTIVE_FROM = 2
Private Const UNIT_ACTIVE_TO = 3
Private Const UNIT_DX_ID = 4
Private Const UNIT_DX_LOCATION = 5

' Address Tab
Private Const ADDRESS_POSTCODE = 0
Private Const ADDRESS_BUILDING_NAME = 1
Private Const ADDRESS_BUILDING_NO = 2
Private Const ADDRESS_STREET = 3
Private Const ADDRESS_DISTRICT = 4
Private Const ADDRESS_TOWN = 5
Private Const ADDRESS_COUNTY = 6
Private Const ADDRESS_COUNTRY = 7

' Contact Tab
Private Const CONTACT_FORNAME = 0
Private Const CONTACT_SURNAME = 1
Private Const CONTACT_TELEPHONE = 2
Private Const CONTACT_EXT = 3
Private Const CONTACT_FAX_NO = 4
'Private Const CONTACT_EMAIL_ADDRESS = 5


Private m_bIsEdit As Boolean
Private m_clsAddress As New AddressTable
Private m_clsContact As New ContactDetailsTable
Private m_clsUnitContacts As New UnitContactDetailsTable
Private m_clsUnitRegion As New UnitRegionTable
Private m_clsUnit As New UnitTable
Private m_clsTaskOwner As TaskOwner
Private m_ReturnCode As MSGReturnCode
Private m_colComboValues As Collection
Private m_colComboIDS As Collection
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsUnit = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing()
    If bRet = True Then
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Failed
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        bRet = ValidateScreenData()
        
        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
        End If
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing()
    
    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Function ValidateScreenData() As Boolean
    Dim nCount As Integer
    Dim bRet As Boolean
    On Error GoTo Failed
        
    nCount = 0
    bRet = True
    
    bRet = dgRegions.ValidateRows()
    
    If bRet = True And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    
        If bRet = False Then
            MsgBox "Unit already exists", vbCritical
            txtUnit(UNIT_ID).SetFocus
            Me.SSTab1.Tab = UnitDetails
        End If
    End If
    
    If bRet = True Then
        ' On the Address screen, at least on of House Number or house name must be entered
        If Len(txtAddress(ADDRESS_POSTCODE).Text) > 0 And Len(txtAddress(ADDRESS_BUILDING_NAME).Text) = 0 And Len(txtAddress(ADDRESS_BUILDING_NO).Text) = 0 Then
            MsgBox "At least one of Building Name or Building Number must be entered"
            SSTab1.Tab = UnitAddress
            txtAddress(ADDRESS_BUILDING_NAME).SetFocus
            bRet = False
        End If
    End If
    
    If bRet = True Then
        bRet = g_clsValidation.ValidateActiveFromTo(Me.txtUnit(UNIT_ACTIVE_FROM), txtUnit(UNIT_ACTIVE_TO))
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsUnitRegion)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Region must be unique"
        End If
    
    End If
    
    If bRet Then
        bRet = g_clsValidation.ValidatePostCode(Me.txtAddress(ADDRESS_POSTCODE).Text)
        If Not bRet Then
            SSTab1.Tab = UnitAddress
            Me.txtAddress(ADDRESS_POSTCODE).SetFocus
            g_clsErrorHandling.RaiseError errPostCodeInvalid
        End If
    End If
    
    'Validate taskwoner
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rs = dgTaskOwner.DataSource
    
    m_clsTaskOwner.SetForm Me
    m_clsTaskOwner.SetRecordSet rs
    bRet = m_clsTaskOwner.ValidateTaskOwner()
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sUnit As String
    Dim clsTableAccess As TableAccess
    Dim col As New Collection
    
    bRet = False
    
    sUnit = Me.txtUnit(UNIT_ID).Text
    
    If Len(sUnit) > 0 Then
        col.Add sUnit
        Set clsTableAccess = m_clsUnit
        bRet = clsTableAccess.DoesRecordExist(col)
    End If
    
    DoesRecordExist = bRet
End Function
Private Sub dgRegions_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    bCancel = Not dgRegions.ValidateRow(nCurrentRow)
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    Set m_clsUnit = New UnitTable
    Set m_clsAddress = Nothing
    Set m_clsContact = Nothing
    Set m_clsUnitContacts = Nothing
    Set m_clsUnitRegion = Nothing
    Set m_clsTaskOwner = New TaskOwner
    
    SSTab1_Click 0
    GetRegions
    PopulateDepartment
    
    g_clsFormProcessing.PopulateCombo "Country", cboCountry
    g_clsFormProcessing.PopulateCombo "ContactType", cboContactType
    g_clsFormProcessing.PopulateCombo "Title", cboContactTitle

    If m_bIsEdit = True Then
        SetEditState
        PopulateScreen
        SetDataGridState
    Else
        SetAddState
    End If
    
    SSTab1.Tab = 0
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub PopulateDepartment()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsDepartment As New DepartmentTable
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim sDepartmentIDField As String
    
    Set clsTableAccess = clsDepartment
    
    Set rs = clsTableAccess.GetTableData(POPULATE_ALL)

    clsTableAccess.ValidateData
    Set cboDepartment.RowSource = rs
    sDepartmentIDField = clsDepartment.GetDepartmentField()
    cboDepartment.ListField = sDepartmentIDField '
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateRegion(sUnit As String)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    Dim rs As ADODB.Recordset
    
    Set clsTableAccess = m_clsUnitRegion
    colValues.Add sUnit
    
    clsTableAccess.SetKeyMatchValues colValues
    
    Set rs = clsTableAccess.GetTableData()

    clsTableAccess.ValidateData
    Set Me.dgRegions.DataSource = rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetAddState()
    PopulateRegion "null"
    SetGridFields
End Sub
Public Sub SetEditState()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGUID As Variant
    Dim sUnitID As String
    Dim colValues As New Collection
    
    ' First, the department table
    Set clsTableAccess = m_clsUnit
    clsTableAccess.SetKeyMatchValues m_colKeys
    Set rs = clsTableAccess.GetTableData()

    txtUnit(UNIT_ID).Enabled = False
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            ' Now the Address table - need the address guid
            sGUID = m_clsUnit.GetAddressGUID()
            sUnitID = m_clsUnit.GetUnitID()
            
            If Len(sGUID) > 0 Then
                Set clsTableAccess = m_clsAddress
                
                colValues.Add sGUID
                clsTableAccess.SetKeyMatchValues colValues
                            
                Set rs = clsTableAccess.GetTableData()
            End If
        
            ' Unit Region
            If Len(sUnitID) > 0 Then
                PopulateRegion sUnitID
            End If
            
            ' Now Contact Details
            If Len(sUnitID) > 0 Then
                Set clsTableAccess = m_clsUnitContacts
                Set colValues = New Collection
            
                colValues.Add sUnitID
                clsTableAccess.SetKeyMatchValues colValues
                Set rs = clsTableAccess.GetTableData()
                
                If Not rs Is Nothing Then
                    If rs.RecordCount > 0 Then
                        sGUID = m_clsUnitContacts.GetContactDetailsGUID()
                        
                        If Len(sGUID) > 0 Then
                            Set colValues = New Collection
                            Set clsTableAccess = m_clsContact
                            
                            colValues.Add sGUID
                            clsTableAccess.SetKeyMatchValues colValues
                            Set rs = clsTableAccess.GetTableData()
                        End If
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
Failed:
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateScreen()
'    Dim rs As adodb.Recordset
    Dim bRet As Boolean
    Dim colTables As New Collection
    On Error GoTo Failed
    
    PopulateScreenFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim clsFixedParameters  As FixedParametersTable
    
    LoadUnit
    LoadAddress
    LoadContact
    LoadTaskOwners
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveUnit()
    On Error GoTo Failed
    Dim vVal As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsUnit
    End If

    ' Save the department ID
    m_clsUnit.SetUnitID txtUnit(UNIT_ID).Text
    
    ' Name
    m_clsUnit.SetUnitName txtUnit(UNIT_NAME).Text
    
    ' Active From
    g_clsFormProcessing.HandleDate txtUnit(UNIT_ACTIVE_FROM), vVal, GET_CONTROL_VALUE
    m_clsUnit.SetActiveFrom vVal

    ' Active To
    'm_clsUnit.SetActiveTo txtUnit(UNIT_ACTIVE_TO).Text
    g_clsFormProcessing.HandleDate txtUnit(UNIT_ACTIVE_TO), vVal, GET_CONTROL_VALUE
    m_clsUnit.SetActiveTo vVal

    ' Department Combo
    g_clsFormProcessing.HandleComboText Me.cboDepartment, vVal, GET_CONTROL_VALUE
    m_clsUnit.SetDepartmentID CStr(vVal)

    ' Organisation Unit
    g_clsFormProcessing.HandleCheckBox chkOrganisationUnit, vVal, GET_CONTROL_VALUE
    m_clsUnit.SetOrganisationIndicator CStr(vVal)
    
    ' Processing Unit
    g_clsFormProcessing.HandleCheckBox chkProcessingUnit, vVal, GET_CONTROL_VALUE
    m_clsUnit.SetProcessingUnitIndicator CStr(vVal)

    ' DX ID
    m_clsUnit.SetDXID txtUnit(UNIT_DX_ID).Text
    
    ' DX Location
    m_clsUnit.SetDXLocation txtUnit(UNIT_DX_LOCATION).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadUnit()
    Dim vVal As Variant
    On Error GoTo Failed
    
    ' Save the department ID
    txtUnit(UNIT_ID).Text = m_clsUnit.GetUnitID()
    ' Name
    txtUnit(UNIT_NAME).Text = m_clsUnit.GetUnitName()
    ' Active From
    txtUnit(UNIT_ACTIVE_FROM).Text = m_clsUnit.GetActiveFrom()
    ' Active To
    txtUnit(UNIT_ACTIVE_TO).Text = m_clsUnit.GetActiveTo()
    ' DX ID
    txtUnit(UNIT_DX_ID).Text = m_clsUnit.GetDXID()
    ' DX Location
    txtUnit(UNIT_DX_LOCATION).Text = m_clsUnit.GetDXLocation()
    
    ' Organisation Unit
    g_clsFormProcessing.HandleCheckBox chkOrganisationUnit, m_clsUnit.GetOrganisationIndicator(), SET_CONTROL_VALUE
    
    ' Processing Unit
    g_clsFormProcessing.HandleCheckBox chkProcessingUnit, m_clsUnit.GetProcessingUnitIndicator, SET_CONTROL_VALUE
    
    ' Department Combo
    g_clsFormProcessing.HandleComboText Me.cboDepartment, m_clsUnit.GetDepartmentID(), SET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub LoadContact()
    Dim vTmp As Variant
    On Error GoTo Failed
    
    txtContactDetails(CONTACT_FORNAME).Text = m_clsContact.GetContactForname()
    txtContactDetails(CONTACT_SURNAME).Text = m_clsContact.GetContactSurname()
    txtContactDetails(CONTACT_TELEPHONE).Text = m_clsContact.GetTelNo()
    txtContactDetails(CONTACT_EXT).Text = m_clsContact.GetTelExtNo()
    txtContactDetails(CONTACT_FAX_NO).Text = m_clsContact.GetFaxNumber()
    txtEmailAddress.Text = m_clsContact.GetEmailAddress()

    vTmp = m_clsContact.GetContactType
    g_clsFormProcessing.HandleComboExtra cboContactType, vTmp, SET_CONTROL_VALUE
    
    vTmp = m_clsContact.GetContactTitle
    g_clsFormProcessing.HandleComboExtra cboContactTitle, vTmp, SET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadAddress()
    Dim vTmp As Variant
    
    On Error GoTo Failed
    
    txtAddress(ADDRESS_POSTCODE).Text = m_clsAddress.GetPostCode()
    txtAddress(ADDRESS_BUILDING_NAME).Text = m_clsAddress.GetBuildingOrHouseName()
    txtAddress(ADDRESS_BUILDING_NO).Text = m_clsAddress.GetBuildingOrHouseNumber()
    txtAddress(ADDRESS_STREET).Text = m_clsAddress.GetStreet()
    txtAddress(ADDRESS_DISTRICT).Text = m_clsAddress.GetDistrict()
    txtAddress(ADDRESS_TOWN).Text = m_clsAddress.GetTown()
    txtAddress(ADDRESS_COUNTY).Text = m_clsAddress.GetCounty()
    
    vTmp = m_clsAddress.GetCountry()
    g_clsFormProcessing.HandleComboExtra cboCountry, vTmp, SET_CONTROL_VALUE

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveAddress()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim vGUID As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsAddress
        
        vGUID = m_clsAddress.SetAddressGUID()
        
        If Not IsNull(vGUID) Then
            m_clsUnit.SetAddressGUID CStr(vGUID)
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveAddress: Address GUID is empty"
        End If
    End If

    g_clsFormProcessing.HandleComboExtra cboCountry, vTmp, GET_CONTROL_VALUE
    m_clsAddress.SetCountry vTmp
        
    m_clsAddress.SetPostCode txtAddress(ADDRESS_POSTCODE).Text
    m_clsAddress.SetBuildingOrHouseName txtAddress(ADDRESS_BUILDING_NAME).Text
    m_clsAddress.SetBuildingOrHouseNumber txtAddress(ADDRESS_BUILDING_NO).Text
    m_clsAddress.SetStreet txtAddress(ADDRESS_STREET).Text
    m_clsAddress.SetDistrict txtAddress(ADDRESS_DISTRICT).Text
    m_clsAddress.SetTown txtAddress(ADDRESS_TOWN).Text
    m_clsAddress.SetCounty txtAddress(ADDRESS_COUNTY).Text
    
    g_clsFormProcessing.HandleComboExtra cboCountry, vTmp, GET_CONTROL_VALUE
    m_clsAddress.SetCountry vTmp
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveContact()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    Dim vGUID As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsContact
    
        vGUID = m_clsContact.SetContactDetailsGUID()
    
        If Not IsNull(vGUID) Then
            g_clsFormProcessing.CreateNewRecord m_clsUnitContacts
        
            m_clsUnitContacts.SetContactDetailsGUID CStr(vGUID)
            m_clsUnitContacts.SetUnitID txtUnit(UNIT_ID).Text
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveContact: ContactDetails GUID is empty"
        End If
    End If
    
    m_clsContact.SetContactForname txtContactDetails(CONTACT_FORNAME).Text
    m_clsContact.SetContactSurname txtContactDetails(CONTACT_SURNAME).Text
    m_clsContact.SetTelNo txtContactDetails(CONTACT_TELEPHONE).Text
    m_clsContact.SetTelExtNo txtContactDetails(CONTACT_EXT).Text
    m_clsContact.SetFaxNumber txtContactDetails(CONTACT_FAX_NO).Text
    m_clsContact.SetEmailAddress txtEmailAddress.Text
    
    g_clsFormProcessing.HandleComboExtra cboContactType, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactType vTmp

    g_clsFormProcessing.HandleComboExtra cboContactTitle, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactTitle vTmp
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess

    SaveUnit
    SaveAddress
    SaveContact
    
    ' First, Address
    Set clsTableAccess = m_clsAddress
    clsTableAccess.Update
        
    ' Contact details
    Set clsTableAccess = m_clsContact
    clsTableAccess.Update
        
    ' Unit
    Set clsTableAccess = m_clsUnit
    clsTableAccess.Update
        
    ' Department Contact details
    Set clsTableAccess = m_clsUnitContacts
    clsTableAccess.Update
    
    ' UnitRegion
    Set clsTableAccess = m_clsUnitRegion
    clsTableAccess.Update

    'Task Owner
    m_clsTaskOwner.SetUnitID txtUnit(UNIT_ID).Text
    m_clsTaskOwner.SaveTaskOwner
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    SetTabstops Me
End Sub

Private Sub txtAddress_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = ADDRESS_POSTCODE Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtAddress_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtAddress(Index).ValidateData()
End Sub
Private Sub txtContactDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtContactDetails(Index).ValidateData()
End Sub
Private Sub txtUnit_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo Failed
    Cancel = Not txtUnit(Index).ValidateData()

    If Cancel = False Then
        SetDataGridState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    Cancel = True
End Sub
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsCombo As New ComboUtils
    Dim sCombo As String
    bRet = True

    ' UnitID
    fields.sField = "UnitID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtUnit(UNIT_ID).Text
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' RegionID
    fields.sField = "RegionID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Region ID Text
    fields.sField = "RegionNameText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Region ID must be selected"
    fields.sTitle = "Region Name"
    fields.sOtherField = "RegionID"
    
    If m_colComboValues.Count > 0 And m_colComboIDS.Count > 0 Then
        Set fields.colComboValues = m_colComboValues
        Set fields.colComboIDS = m_colComboIDS
        colFields.Add fields
    Else
        MsgBox "Unable to locate combo group" + sCombo
    End If

    Me.dgRegions.SetColumns colFields, "EditUnitRegion", "Regions"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'Private Function GetRegions(colComboValues As Collection, colComboIDS As Collection) As Boolean
Private Sub GetRegions()
    On Error GoTo Failed
    ' Need to get the region name and id's
    Dim clsRegion As New RegionTable
    
    Set m_colComboValues = New Collection
    Set m_colComboIDS = New Collection
    
    clsRegion.GetRegionsAsCollection m_colComboValues, m_colComboIDS

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetDataGridState()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False

    Dim bEnabled As Boolean

    ' Unit ID has to be entered
    
    If Len(Me.txtUnit(UNIT_ID).Text) > 0 Then
        bEnabled = dgRegions.Enabled
        dgRegions.Enabled = True
    
        SetGridFields
    
        If m_bIsEdit = False Then
            If bEnabled = False Then
                'dgRegions.AddRow
                'dgRegions.SetFocus
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sUnitID As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sUnitID = Me.txtUnit(UNIT_ID).Text
    colMatchValues.Add sUnitID
    
    Set clsTableAccess = m_clsUnit
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsUnit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LoadTaskOwners
' Description   : Retrieves the Task Owner data and calls the taskowner class to set the datagrid datasource etc
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadTaskOwners()
    
    Dim clsTableAccess As TableAccess
    Dim clsOmigaUser As OmigaUserTable
    Dim rs As ADODB.Recordset
    Dim sUnitID As String
    Dim clsTaskOwner As TaskOwner
    On Error GoTo Failed
    
    Set clsTaskOwner = New TaskOwner
    Set clsOmigaUser = New OmigaUserTable
    Set clsTableAccess = clsOmigaUser
    
    sUnitID = txtUnit(UNIT_ID).Text
    If Len(sUnitID) = 0 Then
        g_clsErrorHandling.RaiseError , "frmEditUnit: UnitID cannot be null"
    End If
    
    'Function gets all omiga users, with activeroles/users
    clsOmigaUser.GetUsersFromUnit (sUnitID)
    
    Set rs = clsTableAccess.GetRecordSet
    clsTaskOwner.SetForm Me
    clsTaskOwner.SetUnitID (sUnitID)
    
    clsTaskOwner.Initialise m_bIsEdit
    
    If rs.RecordCount = 0 Then
        'There are no users with a UserRole.UnitID set to current UnitID
        SSTab1.TabEnabled(UnitTaskOwner) = False
    End If
    dgTaskOwner.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


