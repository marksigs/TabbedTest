VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditUnit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Unit"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
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
      TabIndex        =   20
      Top             =   120
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Unit Details"
      TabPicture(0)   =   "frmEditUnit.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Address Details"
      TabPicture(1)   =   "frmEditUnit.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contact Details"
      TabPicture(2)   =   "frmEditUnit.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Task Owners"
      TabPicture(3)   =   "frmEditUnit.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dgTaskOwner"
      Tab(3).ControlCount=   1
      Begin MSGOCX.MSGDataGrid dgTaskOwner 
         Height          =   3495
         Left            =   -74520
         TabIndex        =   45
         Top             =   1080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6165
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
         Height          =   3315
         Left            =   -74940
         TabIndex        =   39
         Tag             =   "Tab3"
         Top             =   360
         Width           =   8595
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   0
            Left            =   5460
            TabIndex        =   19
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
         Begin MSGOCX.MSGEditBox txtEmailAddress 
            Height          =   315
            Left            =   2100
            TabIndex        =   55
            Top             =   3000
            Width           =   5900
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
         Begin MSGOCX.MSGComboBox cboType2 
            Height          =   315
            Left            =   480
            TabIndex        =   50
            Top             =   2520
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
            Left            =   480
            TabIndex        =   46
            Top             =   2040
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   3
            Left            =   3280
            TabIndex        =   47
            Top             =   2040
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   4
            Left            =   4580
            TabIndex        =   48
            Top             =   2040
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   5
            Left            =   6600
            TabIndex        =   49
            Top             =   2040
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   6
            Left            =   2100
            TabIndex        =   51
            Top             =   2520
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   7
            Left            =   3280
            TabIndex        =   52
            Top             =   2520
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   8
            Left            =   4580
            TabIndex        =   53
            Top             =   2520
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   9
            Left            =   6600
            TabIndex        =   54
            Top             =   2520
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
         Begin MSGOCX.MSGComboBox cboContactType 
            Height          =   315
            Left            =   2100
            TabIndex        =   63
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   64
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
         Begin MSGOCX.MSGComboBox cboContactTitle 
            Height          =   315
            Left            =   2100
            TabIndex        =   65
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   2
            Left            =   2100
            TabIndex        =   66
            Top             =   2040
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
         Begin VB.Label lblContact 
            Caption         =   "Ext. No."
            Height          =   195
            Index           =   11
            Left            =   6600
            TabIndex        =   60
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblContact 
            Caption         =   "Country Code"
            Height          =   255
            Index           =   10
            Left            =   2100
            TabIndex        =   59
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "Area Code"
            Height          =   255
            Left            =   3280
            TabIndex        =   58
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "Telephone Number"
            Height          =   255
            Left            =   4580
            TabIndex        =   57
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label21 
            Caption         =   "Type"
            Height          =   255
            Left            =   480
            TabIndex        =   56
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Type"
            Height          =   255
            Index           =   13
            Left            =   480
            TabIndex        =   44
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Forename"
            Height          =   255
            Index           =   15
            Left            =   4620
            TabIndex        =   43
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Surname"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   42
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Email Address"
            Height          =   195
            Index           =   20
            Left            =   480
            TabIndex        =   41
            Top             =   3060
            Width           =   990
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Title"
            Height          =   255
            Index           =   14
            Left            =   480
            TabIndex        =   40
            Top             =   960
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3195
         Left            =   -74940
         TabIndex        =   30
         Tag             =   "Tab2"
         Top             =   360
         Width           =   8595
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   38
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Building Name"
            Height          =   255
            Index           =   6
            Left            =   420
            TabIndex        =   37
            Top             =   840
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "No."
            Height          =   255
            Index           =   7
            Left            =   6540
            TabIndex        =   36
            Top             =   840
            Width           =   435
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Street"
            Height          =   255
            Index           =   8
            Left            =   420
            TabIndex        =   35
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "District"
            Height          =   255
            Index           =   9
            Left            =   420
            TabIndex        =   34
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Town"
            Height          =   255
            Index           =   10
            Left            =   420
            TabIndex        =   33
            Top             =   1920
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "County"
            Height          =   255
            Index           =   11
            Left            =   420
            TabIndex        =   32
            Top             =   2280
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Country"
            Height          =   255
            Index           =   12
            Left            =   420
            TabIndex        =   31
            Top             =   2640
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5115
         Left            =   120
         TabIndex        =   21
         Tag             =   "Tab1"
         Top             =   360
         Width           =   8475
         Begin VB.TextBox txtUnitName 
            Height          =   315
            Left            =   1860
            MaxLength       =   130
            TabIndex        =   67
            Top             =   600
            Width           =   6135
         End
         Begin VB.CheckBox chkExitFromWrapUp 
            Alignment       =   1  'Right Justify
            Caption         =   "Exit Omiga From Wrap Up Unit?"
            Height          =   255
            Left            =   3840
            TabIndex        =   62
            Top             =   4800
            Width           =   2655
         End
         Begin VB.CheckBox chkLogOnUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "Log On Unit?"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   4800
            Width           =   1815
         End
         Begin VB.CheckBox chkProcessingUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "Processing Unit?"
            Height          =   255
            Left            =   3840
            TabIndex        =   8
            Top             =   4450
            Width           =   2655
         End
         Begin VB.CheckBox chkOrganisationUnit 
            Alignment       =   1  'Right Justify
            Caption         =   "Organisation Unit?"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   4450
            Width           =   1815
         End
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   0
            Left            =   1860
            TabIndex        =   0
            Top             =   200
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
         Begin MSGOCX.MSGDataCombo cboDepartment 
            Height          =   315
            Left            =   1860
            TabIndex        =   1
            Top             =   960
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
            Left            =   1800
            TabIndex        =   2
            Top             =   1380
            Width           =   6315
            _ExtentX        =   11139
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
            Left            =   1860
            TabIndex        =   3
            Top             =   3620
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   3
            Left            =   6285
            TabIndex        =   4
            Top             =   3615
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   4
            Left            =   1860
            TabIndex        =   5
            Top             =   4020
            Width           =   1155
            _ExtentX        =   2037
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
         Begin MSGOCX.MSGEditBox txtUnit 
            Height          =   315
            Index           =   5
            Left            =   6285
            TabIndex        =   6
            Top             =   4020
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
            AutoSize        =   -1  'True
            Caption         =   "DX Location"
            Height          =   195
            Index           =   23
            Left            =   3840
            TabIndex        =   29
            Top             =   4080
            Width           =   885
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "DX ID"
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   28
            Top             =   4080
            Width           =   435
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Department"
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   27
            Top             =   1020
            Width           =   825
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Unit ID"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   260
            Width           =   495
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Unit Name"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   640
            Width           =   750
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Regions"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   1680
            Width           =   585
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Active From"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   3680
            Width           =   840
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Active To"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   22
            Top             =   3675
            Width           =   690
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7590
      TabIndex        =   10
      Top             =   5835
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   6270
      TabIndex        =   9
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
' AA        22/03/01    Added TaskOwner Validation
' DJP       28/08/01    SYS2564 Core functionality
' STB       09/11/01    Added telephone table functionality.
' DB        15/11/02    BMIDS00901 Made DXid max value 99999
' DB        27/11/02    BMIDS01096 DXid now needs a max value of 999999999
' MV        05/03/2003  BM0370  Changed the DXID property to allow alphanumeric
' GHun      17/10/2005  MAR57 Added AllowOmigaLogOn and AllowOmigaExitFromWrap
' TW        13/03/2007  EP2_1934 - DBM341 - Satellite Packagers - Allow unit name of up to 130 characters.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Contains the tab index's for each frmEditUnit Tab
Private tabsUnit As UnitTabs

' Unit Tab
Private Const UNIT_ID = 0
' TW 13/03/2007 EP2_1934
'Private Const UNIT_NAME = 1
' TW 13/03/2007 EP2_1934 End
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
Private Const COUNTRY_CODE1 = 2
Private Const AREA_CODE1 = 3
Private Const TELEPHONE1 = 4
Private Const EXTENSION1 = 5
Private Const COUNTRY_CODE2 = 6
Private Const AREA_CODE2 = 7
Private Const TELEPHONE2 = 8
Private Const EXTENSION2 = 9

'The number of telephone records per contact.
Private Const NUM_TELEPHONE_RECORDS As Long = 2

' Private data
Private m_bIsEdit As Boolean
Private m_clsUnit As UnitTable
Private m_clsAddress As AddressTable
Private m_clsContact As ContactDetailsTable
Private m_clsContactTelephone As ContactDetailsTelephoneTable
Private m_clsTaskOwner As TaskOwner
Private m_clsUnitRegion As UnitRegionTable
Private m_clsUnitContacts As UnitContactDetailsTable

Private m_ReturnCode As MSGReturnCode
Private m_colComboIDS As Collection
Private m_colComboValues As Collection

' Keys to be used on edit to locate the record(s) we need.
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called on form initialisation. Creates all nececcary table classes and reads
'                 combo values.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    ' Create all table classes as required
    Set m_clsUnit = New UnitTable
    Set m_clsAddress = New AddressTable
    Set m_clsContact = New ContactDetailsTable
    Set m_clsTaskOwner = New TaskOwner
    Set m_clsUnitRegion = New UnitRegionTable
    Set m_clsUnitContacts = New UnitContactDetailsTable
    Set m_clsContactTelephone = New ContactDetailsTelephoneTable
    
    ' Do tab initialisation for the UNIT tab
    SSTab1_Click UnitDetails
    
    ' Need to populate the Regions combo with a list of all regions
    GetRegions
    
    ' Read Department record too
    PopulateDepartment
    
    ' Populate all combos as required
    g_clsFormProcessing.PopulateCombo "Country", cboCountry
    g_clsFormProcessing.PopulateCombo "ContactType", cboContactType
    g_clsFormProcessing.PopulateCombo "Title", cboContactTitle
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboType1
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboType2

    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    SSTab1.Tab = 0
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Common method called when the user presses OK or Another. Performs form validation
'                 and if successful, saves all screen data to the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
' TW 13/03/2007 EP2_1934
    If Len(txtUnitName.Text) = 0 Then
        txtUnitName.BackColor = vbRed
        g_clsErrorHandling.RaiseError errMandatoryFieldsRequired
        txtUnitName.SetFocus
        Exit Function
    End If
' TW 13/03/2007 EP2_1934 End
    
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all screen data - needs to ensure that the datagrid has no duplicate
'                 rows, that the unit doesn't already exist (in add mode) and the active to/from
'                 dates are valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
        
    ' Check for duplicates on the databgrid
    bRet = True
    bRet = dgRegions.ValidateRows()
    
    ' Make sure the unit doesn't exist already.
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
    
    ' AA added if bret, as the taskowner validation was always being called, and setting function to TRUE
    If bRet Then
        'Validate taskwoner
        Dim rs As ADODB.Recordset
        
        Set rs = dgTaskOwner.DataSource
        
        m_clsTaskOwner.SetForm Me
        m_clsTaskOwner.SetRecordSet rs
        bRet = m_clsTaskOwner.ValidateTaskOwner()
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesRecordExist
' Description   : Checks if the current unit exists on the database already or not. Returns TRUE
'                 if it does, false if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sUnit As String
    Dim col As New Collection
    
    bRet = False
    sUnit = txtUnit(UNIT_ID).Text
    
    If Len(sUnit) > 0 Then
        col.Add sUnit
        bRet = TableAccess(m_clsUnit).DoesRecordExist(col)
    End If
    
    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateDepartment
' Description   : Populates the department table class, populating all records. This is so we
'                 can populate the Department combo with all departments.s
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateDepartment()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsDepartment As DepartmentTable
    Dim sDepartmentIDField As String
    
    Set clsDepartment = New DepartmentTable
    Set rs = TableAccess(clsDepartment).GetTableData(POPULATE_ALL)

    TableAccess(clsDepartment).ValidateData
    
    ' Set the datacombo to the recordset returned.
    Set cboDepartment.RowSource = rs
    sDepartmentIDField = clsDepartment.GetDepartmentField()
    cboDepartment.ListField = sDepartmentIDField
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateRegion
' Description   : Populates the region table class, populating all records for this unit. The datasource
'                 of the datagrid is then set to the recordset returned.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateRegion(sUnit As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim colValues As New Collection
    
    colValues.Add sUnit
    TableAccess(m_clsUnitRegion).SetKeyMatchValues colValues
    
    Set rs = TableAccess(m_clsUnitRegion).GetTableData()

    TableAccess(m_clsUnitRegion).ValidateData
    Set dgRegions.DataSource = rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Performs all startup processing - load all relevant data from the database
'                 including Unit data, Contact details and Address details.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    On Error GoTo Failed
    Dim vGUID As Variant
    Dim sUnitID As String
    Dim colValues As Collection
    
    Set colValues = New Collection
    
    ' If editing, we can't change the UNIT ID
    txtUnit(UNIT_ID).Enabled = False
    
    ' First, the Unit table
    TableAccess(m_clsUnit).SetKeyMatchValues m_colKeys
    TableAccess(m_clsUnit).GetTableData

    If TableAccess(m_clsUnit).RecordCount > 0 Then
        ' Now the Address table - need the address guid
        vGUID = m_clsUnit.GetAddressGUID()
        sUnitID = m_clsUnit.GetUnitId()
        
        If Len(vGUID) > 0 Then
            colValues.Add vGUID
            TableAccess(m_clsAddress).SetKeyMatchValues colValues
            TableAccess(m_clsAddress).GetTableData
        End If
    
        ' Unit Region
        If Len(sUnitID) > 0 Then
            PopulateRegion sUnitID
        End If
        
        ' Now Contact Details
        If Len(sUnitID) > 0 Then
            Set colValues = New Collection
        
            colValues.Add sUnitID
            TableAccess(m_clsUnitContacts).SetKeyMatchValues colValues
            TableAccess(m_clsUnitContacts).GetTableData
            
            If TableAccess(m_clsUnitContacts).RecordCount > 0 Then
                vGUID = m_clsUnitContacts.GetContactDetailsGUID()
                
                If Len(vGUID) > 0 Then
                    Set colValues = New Collection
                    
                    colValues.Add vGUID
                    TableAccess(m_clsContact).SetKeyMatchValues colValues
                    TableAccess(m_clsContact).GetTableData
                    
                    ' Now Telephone numbers.
                    TableAccess(m_clsContactTelephone).SetKeyMatchValues colValues
                    TableAccess(m_clsContactTelephone).GetTableData
                End If
            End If
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Unit"
    End If
    
    ' Populate all screen fields and setup the datagrid
    PopulateScreenFields
    SetDataGridState
    
    Exit Sub
Failed:
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : This form is made up of Unit, Address and Contact details data. This method
'                 retrieves all relevant data and populates all controls on the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    
    LoadUnit
    LoadAddress
    LoadContact
    
    If m_bIsEdit Then
        LoadTaskOwners
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveUnit()
    On Error GoTo Failed
    Dim vval As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsUnit
    End If

    ' Save the department ID
    m_clsUnit.SetUnitID txtUnit(UNIT_ID).Text
    
    ' Name
' TW 13/03/2007 EP2_1934
'    m_clsUnit.SetUnitName txtUnit(UNIT_NAME).Text
    m_clsUnit.SetUnitName txtUnitName.Text
' TW 13/03/2007 EP2_1934 End
    
    ' Active From
    g_clsFormProcessing.HandleDate txtUnit(UNIT_ACTIVE_FROM), vval, GET_CONTROL_VALUE
    m_clsUnit.SetActiveFrom vval

    ' Active To
    'm_clsUnit.SetActiveTo txtUnit(UNIT_ACTIVE_TO).Text
    g_clsFormProcessing.HandleDate txtUnit(UNIT_ACTIVE_TO), vval, GET_CONTROL_VALUE
    m_clsUnit.SetActiveTo vval

    ' Department Combo
    g_clsFormProcessing.HandleComboText Me.cboDepartment, vval, GET_CONTROL_VALUE
    m_clsUnit.SetDepartmentID CStr(vval)

    ' Organisation Unit
    g_clsFormProcessing.HandleCheckBox chkOrganisationUnit, vval, GET_CONTROL_VALUE
    m_clsUnit.SetOrganisationIndicator CStr(vval)
    
    ' Processing Unit
    g_clsFormProcessing.HandleCheckBox chkProcessingUnit, vval, GET_CONTROL_VALUE
    m_clsUnit.SetProcessingUnitIndicator CStr(vval)

    ' DX ID
    m_clsUnit.SetDXID txtUnit(UNIT_DX_ID).Text
    
    ' DX Location
    m_clsUnit.SetDXLocation txtUnit(UNIT_DX_LOCATION).Text
    
    'MAR57 GHun
    g_clsFormProcessing.HandleCheckBox chkLogOnUnit, vval, GET_CONTROL_VALUE
    m_clsUnit.SetAllowOmigaLogon CStr(vval)
    
    g_clsFormProcessing.HandleCheckBox chkExitFromWrapUp, vval, GET_CONTROL_VALUE
    m_clsUnit.SetAllowOmigaExitFromWrap CStr(vval)
    'MAR57 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadUnit()
    Dim vval As Variant
    On Error GoTo Failed
    
    ' Save the department ID
    txtUnit(UNIT_ID).Text = m_clsUnit.GetUnitId()
    ' Name
' TW 13/03/2007 EP2_1934
'    txtUnit(UNIT_NAME).Text = m_clsUnit.GetUnitName()
    txtUnitName.Text = m_clsUnit.GetUnitName()
' TW 13/03/2007 EP2_1934 End
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
    
    'MAR57 GHun
    g_clsFormProcessing.HandleCheckBox chkLogOnUnit, m_clsUnit.GetAllowOmigaLogon(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox chkExitFromWrapUp, m_clsUnit.GetAllowOmigaExitFromWrap(), SET_CONTROL_VALUE
    'MAR57 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadContact()
    Dim vTmp As Variant
    On Error GoTo Failed
    
    txtContact(CONTACT_FORNAME).Text = m_clsContact.GetContactForname()
    txtContact(CONTACT_SURNAME).Text = m_clsContact.GetContactSurname()
    txtEmailAddress.Text = m_clsContact.GetEmailAddress()

    vTmp = m_clsContact.GetContactType
    g_clsFormProcessing.HandleComboExtra cboContactType, vTmp, SET_CONTROL_VALUE
    
    vTmp = m_clsContact.GetContactTitle
    g_clsFormProcessing.HandleComboExtra cboContactTitle, vTmp, SET_CONTROL_VALUE
    
    'Load both telephone numbers - if available.
    If TableAccess(m_clsContactTelephone).RecordCount = NUM_TELEPHONE_RECORDS Then
        TableAccess(m_clsContactTelephone).MoveFirst
        
        'Set the first telephone number now.
        txtContact(COUNTRY_CODE1).Text = m_clsContactTelephone.GetCOUNTRYCODE()
        txtContact(AREA_CODE1).Text = m_clsContactTelephone.GetAREA_CODE()
        txtContact(TELEPHONE1).Text = m_clsContactTelephone.GetTELEPHONE()
        txtContact(EXTENSION1).Text = m_clsContactTelephone.GetTELEPHONE_EXT()
        
        'Set the Number type combo selection.
        vTmp = m_clsContactTelephone.GetType()
        g_clsFormProcessing.HandleComboExtra cboType1, vTmp, SET_CONTROL_VALUE
        
        'Move onto the next telephone number.
         TableAccess(m_clsContactTelephone).MoveNext
    
         'Update screen elements from the Contact Details Telephone table.
         txtContact(COUNTRY_CODE2).Text = m_clsContactTelephone.GetCOUNTRYCODE()
         txtContact(AREA_CODE2).Text = m_clsContactTelephone.GetAREA_CODE()
         txtContact(TELEPHONE2).Text = m_clsContactTelephone.GetTELEPHONE()
         txtContact(EXTENSION2).Text = m_clsContactTelephone.GetTELEPHONE_EXT()
         
         'Set the Number type combo selection.
         vTmp = m_clsContactTelephone.GetType()
         g_clsFormProcessing.HandleComboExtra cboType2, vTmp, SET_CONTROL_VALUE
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadAddress()
    Dim vTmp As Variant
    
    On Error GoTo Failed
    
    txtAddress(ADDRESS_POSTCODE).Text = m_clsAddress.GetPostcode()
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
        
    m_clsAddress.SetPostcode txtAddress(ADDRESS_POSTCODE).Text
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
    Else
        vGUID = m_clsContact.GetContactDetailsGUID
    End If
    
    m_clsContact.SetContactForname txtContact(CONTACT_FORNAME).Text
    m_clsContact.SetContactSurname txtContact(CONTACT_SURNAME).Text
    m_clsContact.SetEmailAddress txtEmailAddress.Text
    
    g_clsFormProcessing.HandleComboExtra cboContactType, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactType vTmp

    g_clsFormProcessing.HandleComboExtra cboContactTitle, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactTitle vTmp
    
    'Ensure that two records exist in the telephone number table before we try
    EnsureTelephoneRecordsExist vGUID

    'Update both records in the telephone numbers table, only if they are both there.
    If TableAccess(m_clsContactTelephone).RecordCount = NUM_TELEPHONE_RECORDS Then
        'Ensure this is the first telephone number.
        TableAccess(m_clsContactTelephone).MoveFirst

        'Update Telephone table from the screen elements.
        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE1).Text
        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE1).Text
        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE1).Text
        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION1).Text

        'Read the number type value into a variant first.
        g_clsFormProcessing.HandleComboExtra cboType1, vTmp, GET_CONTROL_VALUE
        m_clsContactTelephone.SetType vTmp

        'Move onto the next telephone number.
        TableAccess(m_clsContactTelephone).MoveNext

        'Update Telephone table from the screen elements.
        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE2).Text
        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE2).Text
        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE2).Text
        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION2).Text

        'Read the number type value into a variant first.
        g_clsFormProcessing.HandleComboExtra cboType2, vTmp, GET_CONTROL_VALUE
        m_clsContactTelephone.SetType vTmp
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'Private Sub SaveTelephone()
'
'    Dim sType As String
'    Dim vTmp As Variant
'
'    On Error GoTo Failed
'
'    If m_bIsEdit = False Then
'        'Create a blank record to hold the FIRST phone detail
'        g_clsFormProcessing.CreateNewRecord m_clsContactTelephone
'
'        'Save the primary key detail
'        m_clsContactTelephone.SetCONTACTDETAILSTELEPHONEGUID m_clsContact.GetContactDetailsGUID
'        m_clsContactTelephone.SetTELEPHONE_SEQ_NUMBER 1
'
'        'Save the FIRST phone number detail
'        g_clsFormProcessing.HandleComboExtra Me.cboType1, vTmp, GET_CONTROL_VALUE
'        sType = vTmp
'
'        m_clsContactTelephone.SetType sType
'        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE1).Text
'        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE1).Text
'        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE1).Text
'        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION1).Text
'
'        'Create a blank record to hold the SECOND phone detail
'        g_clsFormProcessing.CreateNewRecord m_clsContactTelephone
'
'        'Save the primary key detail
'        m_clsContactTelephone.SetCONTACTDETAILSTELEPHONEGUID m_clsContact.GetContactDetailsGUID
'        m_clsContactTelephone.SetTELEPHONE_SEQ_NUMBER 2
'
'        'Save the SECOND phone number detail
'        g_clsFormProcessing.HandleComboExtra Me.cboType2, vTmp, GET_CONTROL_VALUE
'        sType = vTmp
'
'        m_clsContactTelephone.SetType sType
'        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE2).Text
'        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE2).Text
'        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE2).Text
'        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION2).Text
'    Else
'        TableAccess(m_clsContactTelephone).MoveFirst
'
'        'Save the FIRST phone number detail
'        'Type
'        g_clsFormProcessing.HandleComboExtra Me.cboType1, vTmp, GET_CONTROL_VALUE
'        sType = vTmp
'        m_clsContactTelephone.SetType sType
'        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE1).Text
'        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE1).Text
'        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE1).Text
'        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION1).Text
'
'        TableAccess(m_clsContactTelephone).MoveNext
'
'        'Save the SECOND phone number detail
'        'Type
'        g_clsFormProcessing.HandleComboExtra Me.cboType2, vTmp, GET_CONTROL_VALUE
'        sType = vTmp
'        m_clsContactTelephone.SetType sType
'        m_clsContactTelephone.SetCOUNTRYCODE txtContact(COUNTRY_CODE2).Text
'        m_clsContactTelephone.SetAREA_CODE txtContact(AREA_CODE2).Text
'        m_clsContactTelephone.SetTELEPHONE txtContact(TELEPHONE2).Text
'        m_clsContactTelephone.SetTELEPHONE_EXT txtContact(EXTENSION2).Text
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub


Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess

    SaveUnit
    SaveAddress
    SaveContact
'    SaveTelephone
    
    ' First, Address
    Set clsTableAccess = m_clsAddress
    clsTableAccess.Update
        
    ' Contact details
    Set clsTableAccess = m_clsContact
    clsTableAccess.Update
        
    ' Contact telephone details
    Set clsTableAccess = m_clsContactTelephone
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
    If SSTab1.TabEnabled(UnitTaskOwner) = True Then
        m_clsTaskOwner.SetUnitID txtUnit(UNIT_ID).Text
        m_clsTaskOwner.SaveTaskOwner
    End If
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
Private Sub txtContact_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtContact(Index).ValidateData()
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Sets the field names and field behaviour of the datagrid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection

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
        g_clsErrorHandling.RaiseError errGeneralError, "SetGridFields: Unable to locate combo group"
    End If

    Me.dgRegions.SetColumns colFields, "EditUnitRegion", "Regions"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetRegions()
    On Error GoTo Failed
    ' Need to get the region name and id's
    Dim clsRegion As RegionTable
    
    Set clsRegion = New RegionTable
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
    End If
    
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
Public Sub SetAddState()
    PopulateRegion "null"
    SSTab1.TabEnabled(UnitTaskOwner) = False
    SetGridFields
End Sub
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
Private Sub dgRegions_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgRegions.ValidateRow(nCurrentRow)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnsureTelephoneRecordsExist
' Description   : Ensure both telephone records exist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureTelephoneRecordsExist(ByVal pvGUID As Variant)
    
    Dim sGuid As String
    Dim colValues As Collection
    
    Select Case TableAccess(m_clsContactTelephone).RecordCount
        Case 0
            ' create 2 new records
            g_clsFormProcessing.CreateNewRecord m_clsContactTelephone
            m_clsContactTelephone.SetCONTACTDETAILSTELEPHONEGUID pvGUID
            m_clsContactTelephone.SetTELEPHONE_SEQ_NUMBER 1
        
            g_clsFormProcessing.CreateNewRecord m_clsContactTelephone
            m_clsContactTelephone.SetCONTACTDETAILSTELEPHONEGUID pvGUID
            m_clsContactTelephone.SetTELEPHONE_SEQ_NUMBER 2
                        
        Case 1
            ' only require one new record - assume sequence 1 exists
            g_clsFormProcessing.CreateNewRecord m_clsContactTelephone
            m_clsContactTelephone.SetCONTACTDETAILSTELEPHONEGUID pvGUID
            m_clsContactTelephone.SetTELEPHONE_SEQ_NUMBER 2
                        
    End Select

    'Create a collection to hold the keys values.
    Set colValues = New Collection
    
    'Convert the contacts guid to one which will work with the filter.
    If TypeName(pvGUID) = "Byte()" Then
        sGuid = g_clsSQLAssistSP.GuidToString(CStr(pvGUID))
    Else
        sGuid = pvGUID
    End If
    
    'Apply a filter to the telephone numbers, by adding records we open the
    'whole table. We are simply interested in the records for this GUID.
    colValues.Add sGuid
    TableAccess(m_clsContactTelephone).SetKeyMatchValues colValues
    TableAccess(m_clsContactTelephone).ApplyFilter
    
End Sub

' TW 13/03/2007 EP2_1934
Private Sub txtUnitName_LostFocus()
    txtUnitName.BackColor = vbWhite
End Sub
' TW 13/03/2007 EP2_1934 End


