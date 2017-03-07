VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditDepartment 
   Caption         =   "Add/Edit Department"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   Icon            =   "frmEditDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8100
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6780
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5460
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   135
      TabIndex        =   20
      Top             =   135
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6694
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Department Details"
      TabPicture(0)   =   "frmEditDepartment.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Address Details"
      TabPicture(1)   =   "frmEditDepartment.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Contact Details"
      TabPicture(2)   =   "frmEditDepartment.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   3255
         Left            =   60
         TabIndex        =   36
         Tag             =   "Tab3"
         Top             =   360
         Width           =   9075
         Begin MSGOCX.MSGComboBox cboContactType 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            Top             =   300
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
            Left            =   2160
            TabIndex        =   18
            Top             =   660
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
            Index           =   0
            Left            =   5640
            TabIndex        =   17
            Top             =   300
            Width           =   2415
            _ExtentX        =   4260
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   19
            Top             =   1020
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
         Begin MSGOCX.MSGComboBox cboType2 
            Height          =   315
            Left            =   720
            TabIndex        =   47
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
            Left            =   720
            TabIndex        =   42
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   43
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   3
            Left            =   3360
            TabIndex        =   44
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   4
            Left            =   4680
            TabIndex        =   45
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   5
            Left            =   6720
            TabIndex        =   46
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   6
            Left            =   2160
            TabIndex        =   48
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   7
            Left            =   3360
            TabIndex        =   49
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   8
            Left            =   4680
            TabIndex        =   50
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
         Begin MSGOCX.MSGEditBox txtContact 
            Height          =   315
            Index           =   9
            Left            =   6720
            TabIndex        =   51
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
         Begin MSGOCX.MSGEditBox txtEmailAddress 
            Height          =   315
            Left            =   2160
            TabIndex        =   57
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
         Begin VB.Label Label21 
            Caption         =   "Type."
            Height          =   255
            Left            =   720
            TabIndex        =   56
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Telephone Number."
            Height          =   255
            Left            =   4680
            TabIndex        =   55
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label23 
            Caption         =   "Area Code."
            Height          =   255
            Left            =   3360
            TabIndex        =   54
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblContact 
            Caption         =   "Country Code."
            Height          =   255
            Index           =   10
            Left            =   2160
            TabIndex        =   53
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblContact 
            Caption         =   "Ext. No."
            Height          =   195
            Index           =   11
            Left            =   6720
            TabIndex        =   52
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Email Address"
            Height          =   255
            Index           =   20
            Left            =   540
            TabIndex        =   41
            Top             =   2760
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Surname"
            Height          =   255
            Index           =   16
            Left            =   540
            TabIndex        =   40
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Forename"
            Height          =   255
            Index           =   15
            Left            =   4740
            TabIndex        =   39
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Title"
            Height          =   255
            Index           =   14
            Left            =   540
            TabIndex        =   38
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Contact Type"
            Height          =   255
            Index           =   13
            Left            =   540
            TabIndex        =   37
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3135
         Left            =   -74940
         TabIndex        =   27
         Tag             =   "Tab2"
         Top             =   360
         Width           =   9075
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   2160
            TabIndex        =   15
            Top             =   2460
            Width           =   2955
            _ExtentX        =   5212
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
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   8
            Top             =   300
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
         End
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   9
            Top             =   660
            Width           =   4275
            _ExtentX        =   7541
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
            Left            =   6960
            TabIndex        =   10
            Top             =   660
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
            MaxLength       =   10
         End
         Begin MSGOCX.MSGEditBox txtAddress 
            Height          =   315
            Index           =   3
            Left            =   2160
            TabIndex        =   11
            Top             =   1020
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
            Left            =   2160
            TabIndex        =   12
            Top             =   1380
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
            Left            =   2160
            TabIndex        =   13
            Top             =   1740
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
            Left            =   2160
            TabIndex        =   14
            Top             =   2100
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
         Begin VB.Label lblLenderDetails 
            Caption         =   "Country"
            Height          =   255
            Index           =   12
            Left            =   540
            TabIndex        =   35
            Top             =   2520
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "County"
            Height          =   255
            Index           =   11
            Left            =   540
            TabIndex        =   34
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Town"
            Height          =   255
            Index           =   10
            Left            =   540
            TabIndex        =   33
            Top             =   1800
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "District"
            Height          =   255
            Index           =   9
            Left            =   540
            TabIndex        =   32
            Top             =   1440
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Street"
            Height          =   255
            Index           =   8
            Left            =   540
            TabIndex        =   31
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "No."
            Height          =   255
            Index           =   7
            Left            =   6600
            TabIndex        =   30
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Building Name"
            Height          =   255
            Index           =   6
            Left            =   540
            TabIndex        =   29
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Postcode"
            Height          =   255
            Index           =   5
            Left            =   540
            TabIndex        =   28
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3135
         Left            =   -74760
         TabIndex        =   21
         Tag             =   "Tab1"
         Top             =   360
         Width           =   8655
         Begin MSGOCX.MSGComboBox cboChannel 
            Height          =   315
            Left            =   1980
            TabIndex        =   2
            Top             =   1260
            Width           =   2535
            _ExtentX        =   4471
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
         Begin MSGOCX.MSGEditBox txtDepartment 
            Height          =   315
            Index           =   0
            Left            =   1980
            TabIndex        =   0
            Top             =   540
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
         Begin MSGOCX.MSGEditBox txtDepartment 
            Height          =   315
            Index           =   1
            Left            =   1980
            TabIndex        =   1
            Top             =   900
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
            MaxLength       =   30
         End
         Begin MSGOCX.MSGEditBox txtDepartment 
            Height          =   315
            Index           =   2
            Left            =   1980
            TabIndex        =   3
            Top             =   1620
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSGOCX.MSGEditBox txtDepartment 
            Height          =   315
            Index           =   3
            Left            =   4800
            TabIndex        =   4
            Top             =   1620
            Width           =   975
            _ExtentX        =   1720
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
         Begin VB.Label lblLenderDetails 
            Caption         =   "Active To"
            Height          =   255
            Index           =   4
            Left            =   3780
            TabIndex        =   26
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Active From"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   25
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Channel"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   24
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Department Name"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   23
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Department ID"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Width           =   1395
         End
      End
   End
End
Attribute VB_Name = "frmEditDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditDepartments
' Description   : Form for editing of Omiga 4 departments.
'
' Change history
' Prog      Date        Description
' DJP       28/08/01    Tidy up.
' STB       09/11/01    Added telephone table functionality.
' STB       06/12/01    SYS1942 - Another button commits current transaction. Also, the
'                       Another button didn't work on this form previously!
' STB       04/01/02    SYS3581 - The last tab was displaying a blank form.
' STB       22/02/02    SYS4091 Altered TelephoneUsage combo to use the -
'                       ContactTelephoneUsage combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Department Tab
Private Const DEPARTMENT_ID = 0
Private Const DEPARTMENT_NAME = 1
Private Const DEPARTMENT_ACTIVE_FROM = 2
Private Const DEPARTMENT_ACTIVE_TO = 3

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

' Enums
Private Enum DepartmentTabs
    Department = 0
    Address
    Contact
End Enum

' Private data
Private m_bIsEdit As Boolean
Private m_clsAddress As AddressTable
Private m_clsContact As ContactDetailsTable
Private m_clsContactTelephone As ContactDetailsTelephoneTable
Private m_clsDepartmentContacts As DepartmentContactTable
Private m_clsDepartment As DepartmentTable
Private m_ReturnCode As MSGReturnCode
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
' Description   : Called when the form is loaded. Performs all initialisation code, creating
'                 table classes, populating all combos and populating all relevant tables for
'                 this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Default return code
    SetReturnCode MSGFailure
    
    ' Tables to be used by this form
    Set m_clsAddress = New AddressTable
    Set m_clsContact = New ContactDetailsTable
    Set m_clsDepartment = New DepartmentTable
    Set m_clsDepartmentContacts = New DepartmentContactTable
    Set m_clsContactTelephone = New ContactDetailsTelephoneTable
    
    ' Force the first tab to be clicked for tab initialisation.
    SSTab1_Click 0
    
    ' Populate all combos
    g_clsFormProcessing.PopulateChannel cboChannel
    g_clsFormProcessing.PopulateCombo "Country", cboCountry
    g_clsFormProcessing.PopulateCombo "ContactType", cboContactType
    g_clsFormProcessing.PopulateCombo "Title", cboContactTitle
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboType1
    g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", cboType2

    ' If editing, read fields from database etc - (always in Edit mode here)
    If m_bIsEdit = True Then
        SetEditState
        PopulateScreenFields
    End If
            
    ' The default tab for when the form loads
    SSTab1.Tab = Department

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOKProcessing
' Description   : Common method called when the user presses ok, but also when the user presses
'                 Another. Performs validation and saves all screen data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
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

    DoOkProcessing = bRet
    
    Exit Function
Failed:
    DoOkProcessing = False
    g_clsErrorHandling.DisplayError
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates whether or not all data has been entered correctly on the screen. Must
'                 first check that the record being created is not already in existence (in add mode),
'                 then must perform field validation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet = True Then
        ' On the Address screen, at least one of House Number or house name must be entered
        If Len(txtAddress(ADDRESS_POSTCODE).Text) > 0 And Len(txtAddress(ADDRESS_BUILDING_NAME).Text) = 0 And Len(txtAddress(ADDRESS_BUILDING_NO).Text) = 0 Then
            MsgBox "At least one of Building Name or Building Number must be entered", vbCritical
            SSTab1.Tab = Address
            txtAddress(ADDRESS_BUILDING_NAME).SetFocus
            bRet = False
        End If
        
        If bRet = True Then
            bRet = g_clsValidation.ValidateActiveFromTo(txtDepartment(DEPARTMENT_ACTIVE_FROM), Me.txtDepartment(DEPARTMENT_ACTIVE_TO))
        End If
        
        If bRet Then
            bRet = g_clsValidation.ValidatePostCode(txtAddress(ADDRESS_POSTCODE).Text)
            If Not bRet Then
                SSTab1.Tab = Address
                txtAddress(ADDRESS_POSTCODE).SetFocus
                g_clsErrorHandling.RaiseError errPostCodeInvalid
            End If
        End If
    Else
        MsgBox "Department already exists"
        SSTab1.Tab = Department
        txtDepartment(DEPARTMENT_ID).SetFocus
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesRecordExist
' Description   : Determines whether or not the department being created exists already or not.
'                 Need to read the department ID and test the database. Returns true if it exists,
'                 false if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim colKeyValues As Collection
    Dim sDepartmentID As String
    Dim clsDepartment As DepartmentTable
    
    sDepartmentID = txtDepartment(DEPARTMENT_ID).Text

    If Len(sDepartmentID) > 0 Then
        Set colKeyValues = New Collection
        Set clsDepartment = New DepartmentTable
        
        colKeyValues.Add sDepartmentID
        bRet = TableAccess(clsDepartment).DoesRecordExist(colKeyValues)
    End If

    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Reads all relevant data from the database. This inclues the Department details,
'                 Contact Details and Address Details. The starting point is always the department ID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    Dim vGuid As Variant
    Dim colValues As New Collection
    Dim sDepartmentID As String
    
    On Error GoTo Failed
    
    ' Editing, so can't create "Another"
    cmdAnother.Enabled = False
    
    ' First, the department table - set the keys and get the data
    TableAccess(m_clsDepartment).SetKeyMatchValues m_colKeys

    TableAccess(m_clsDepartment).GetTableData

    If TableAccess(m_clsDepartment).RecordCount > 0 Then
        ' Now the Address table - need the address guid
        vGuid = m_clsDepartment.GetAddressGUID
        sDepartmentID = m_clsDepartment.GetDepartmentID()
        
        If Len(vGuid) > 0 Then
            
            colValues.Add vGuid
            TableAccess(m_clsAddress).SetKeyMatchValues colValues
                        
            TableAccess(m_clsAddress).GetTableData
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Department GUID is empty"
        End If
    
        ' Now Contact Details
        If Len(sDepartmentID) > 0 Then
            Set colValues = New Collection
            colValues.Add sDepartmentID
            
            TableAccess(m_clsDepartmentContacts).SetKeyMatchValues colValues
            TableAccess(m_clsDepartmentContacts).GetTableData
            
            If TableAccess(m_clsDepartmentContacts).RecordCount > 0 Then
                vGuid = m_clsDepartmentContacts.GetContactDetailsGUID()
                
                If Len(vGuid) > 0 Then
                    Set colValues = New Collection
                    
                    colValues.Add vGuid
                    TableAccess(m_clsContact).SetKeyMatchValues colValues
                    TableAccess(m_clsContact).GetTableData
                    
                    ' Now Telephone numbers.
                    TableAccess(m_clsContactTelephone).SetKeyMatchValues colValues
                    TableAccess(m_clsContactTelephone).GetTableData
                End If
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "No contact details record found"
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Called on startup to populate all screen fields once the tables have been loaded
'                 from the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    
    LoadDepartment
    LoadAddress
    LoadContact
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveDepartment()
    On Error GoTo Failed
    Dim vVal As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsDepartment
    End If
        
    ' Save the department ID
    m_clsDepartment.SetDepartmentID txtDepartment(DEPARTMENT_ID).Text
    
    ' Name
    m_clsDepartment.SetName txtDepartment(DEPARTMENT_NAME).Text
    
    ' Channel
    g_clsFormProcessing.HandleComboExtra cboChannel, vVal, GET_CONTROL_VALUE
    m_clsDepartment.SetChannelID CStr(vVal)

    ' Active From
    g_clsFormProcessing.HandleDate txtDepartment(DEPARTMENT_ACTIVE_FROM), vVal, GET_CONTROL_VALUE
    m_clsDepartment.SetActiveFrom vVal

    ' Active To
    g_clsFormProcessing.HandleDate txtDepartment(DEPARTMENT_ACTIVE_TO), vVal, GET_CONTROL_VALUE
    m_clsDepartment.SetActiveTo vVal
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadDepartment()
    On Error GoTo Failed
    Dim vVal As Variant
    
    ' Save the department ID
    txtDepartment(DEPARTMENT_ID).Text = m_clsDepartment.GetDepartmentID()
    
    ' Name
    txtDepartment(DEPARTMENT_NAME).Text = m_clsDepartment.GetName()
    
    ' Channel
    vVal = m_clsDepartment.GetChannelID()
    g_clsFormProcessing.HandleComboExtra cboChannel, vVal, SET_CONTROL_VALUE

    ' Active From
    txtDepartment(DEPARTMENT_ACTIVE_FROM).Text = m_clsDepartment.GetActiveFrom()

    ' Active To
    txtDepartment(DEPARTMENT_ACTIVE_TO).Text = m_clsDepartment.GetActiveTo()
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadContact()
    On Error GoTo Failed
    Dim vTmp As Variant

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
    On Error GoTo Failed
    Dim vTmp As Variant

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
    Dim vGuid As Variant

    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsAddress
        
        vGuid = m_clsAddress.SetAddressGUID()
        
        If Not IsNull(vGuid) Then
            m_clsDepartment.SetAddressGUID CStr(vGuid)
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveAddress - Address GUID is empty"
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
    Dim vTmp As Variant
    Dim vGuid As Variant
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsContact
    
        vGuid = m_clsContact.SetContactDetailsGUID()
    
        If Not IsNull(vGuid) Then
            g_clsFormProcessing.CreateNewRecord m_clsDepartmentContacts
        
            m_clsDepartmentContacts.SetContactDetailsGUID CStr(vGuid)
            m_clsDepartmentContacts.SetDepartmentID txtDepartment(DEPARTMENT_ID).Text
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Departments, ContactDetails GUID is empty"
        End If
    Else
        vGuid = m_clsContact.GetContactDetailsGUID
    End If
    
    m_clsContact.SetContactForname txtContact(CONTACT_FORNAME).Text
    m_clsContact.SetContactSurname txtContact(CONTACT_SURNAME).Text
    m_clsContact.SetEmailAddress txtEmailAddress.Text
    
    g_clsFormProcessing.HandleComboExtra cboContactType, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactType vTmp

    g_clsFormProcessing.HandleComboExtra cboContactTitle, vTmp, GET_CONTROL_VALUE
    m_clsContact.SetContactTitle vTmp
    
    'Ensure that two records exist in the telephone number table before we try
    EnsureTelephoneRecordsExist vGuid
    
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
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    SaveDepartment
    SaveAddress
    SaveContact
        
    ' Now do the Updates.
    
    ' First, Address
    Set clsTableAccess = m_clsAddress
    clsTableAccess.Update
    
    ' Department
    Set clsTableAccess = m_clsDepartment
    clsTableAccess.Update
    
    ' Contact details
    Set clsTableAccess = m_clsContact
    clsTableAccess.Update
    
    ' Contact telephone details
    Set clsTableAccess = m_clsContactTelephone
    clsTableAccess.Update
    
    ' Department Contact details
    Set clsTableAccess = m_clsDepartmentContacts
    clsTableAccess.Update
    
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveChangeRequest
' Description   : For Promtions. When we edit this form we need to record that it has been changed.
'                 This method records the Department ID so it can be promoted to another database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDeptID As String
    Dim colMatchValues As Collection

    Set colMatchValues = New Collection
    
    sDeptID = txtDepartment(DEPARTMENT_ID).Text
    colMatchValues.Add sDeptID
    
    TableAccess(m_clsDepartment).SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsDepartment
    
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
Private Sub txtDepartment_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtDepartment(Index).ValidateData()
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAnother_Click
' Description   : Add the current record and prepare the screen for another.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = DoOkProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)

        If bRet = True Then
            If txtDepartment(DEPARTMENT_ID).Visible = True Then
                txtDepartment(DEPARTMENT_ID).SetFocus
            End If
        End If
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOkProcessing()
    
    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cboChannel_Validate(Cancel As Boolean)
    Cancel = Not cboChannel.ValidateData()
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnsureTelephoneRecordsExist
' Description   : Ensure both telephone records exist.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureTelephoneRecordsExist(ByVal pvGUID As Variant)

    Dim sGUID As String
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
        sGUID = g_clsSQLAssistSP.GuidToString(CStr(pvGUID))
    Else
        sGUID = pvGUID
    End If
    
    'Apply a filter to the telephone numbers, by adding records we open the
    'whole table. We are simply interested in the records for this GUID.
    colValues.Add sGUID
    TableAccess(m_clsContactTelephone).SetKeyMatchValues colValues
    TableAccess(m_clsContactTelephone).ApplyFilter
    
End Sub

