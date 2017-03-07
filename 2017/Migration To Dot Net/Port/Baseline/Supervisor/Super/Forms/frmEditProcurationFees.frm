VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditProcurationFees 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Procuration Fees"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "frmEditProcurationFees.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6795
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11986
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "All Product Fees"
      TabPicture(0)   =   "frmEditProcurationFees.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProductFees"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboMortgageLender(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboProduct(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFrame(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAddPeriod(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditPeriod(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboPeriod(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Mortgage Product Specific Fees"
      TabPicture(1)   =   "frmEditProcurationFees.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblProcuration(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblProcuration(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblProcuration(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboPeriod(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboProduct(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraFrame(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAddPeriod(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdEditPeriod(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cboMortgageLender(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Insurance Product Fees"
      TabPicture(2)   =   "frmEditProcurationFees.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblProcuration(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblProcuration(4)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboMortgageLender(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboPeriod(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraFrame(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdAddPeriod(3)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdEditPeriod(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboProduct(3)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Packaging Fees"
      TabPicture(3)   =   "frmEditProcurationFees.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblProcuration(6)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cboMortgageLender(4)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cboPeriod(4)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboProduct(4)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fraFrame(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdEditPeriod(4)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdAddPeriod(4)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin MSGOCX.MSGComboBox cboPeriod 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   42
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSGOCX.MSGComboBox cboProduct 
         Height          =   315
         Index           =   3
         Left            =   -72480
         TabIndex        =   39
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGComboBox cboMortgageLender 
         Height          =   315
         Index           =   2
         Left            =   -73080
         TabIndex        =   38
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.CommandButton cmdAddPeriod 
         Caption         =   "Add"
         Height          =   315
         Index           =   4
         Left            =   -69960
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdEditPeriod 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   -69000
         TabIndex        =   36
         Top             =   960
         Width           =   915
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Fee Bands for Date Range"
         Height          =   4935
         Index           =   0
         Left            =   -74640
         TabIndex        =   30
         Top             =   1440
         Width           =   7635
         Begin VB.CommandButton cmdAddFee 
            Caption         =   "Add"
            Height          =   375
            Index           =   4
            Left            =   420
            TabIndex        =   33
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditFee 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   1740
            TabIndex        =   32
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteFee 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   3120
            TabIndex        =   31
            Top             =   4320
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvFeeBandList 
            Height          =   3735
            Index           =   4
            Left            =   360
            TabIndex        =   34
            Top             =   420
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   6588
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.CommandButton cmdEditPeriod 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   -68520
         TabIndex        =   27
         Top             =   1320
         Width           =   915
      End
      Begin VB.CommandButton cmdAddPeriod 
         Caption         =   "Add"
         Height          =   315
         Index           =   3
         Left            =   -69540
         TabIndex        =   26
         Top             =   1320
         Width           =   855
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Fee Bands for Date Range"
         Height          =   4335
         Index           =   3
         Left            =   -74400
         TabIndex        =   21
         Top             =   1740
         Width           =   7275
         Begin VB.CommandButton cmdDeleteFee 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   3120
            TabIndex        =   24
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditFee 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   1740
            TabIndex        =   23
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddFee 
            Caption         =   "Add"
            Height          =   375
            Index           =   3
            Left            =   420
            TabIndex        =   22
            Top             =   3720
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvFeeBandList 
            Height          =   3135
            Index           =   3
            Left            =   300
            TabIndex        =   25
            Top             =   420
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   5530
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.CommandButton cmdEditPeriod 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   -68220
         TabIndex        =   20
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton cmdAddPeriod 
         Caption         =   "Add"
         Height          =   315
         Index           =   2
         Left            =   -69120
         TabIndex        =   19
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton cmdEditPeriod 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   6480
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdAddPeriod 
         Caption         =   "Add"
         Height          =   315
         Index           =   1
         Left            =   5460
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Fee Bands for Date Range"
         Height          =   4335
         Index           =   2
         Left            =   -74760
         TabIndex        =   12
         Top             =   2280
         Width           =   7635
         Begin VB.CommandButton cmdDeleteFee 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   3120
            TabIndex        =   15
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditFee 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1740
            TabIndex        =   14
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddFee 
            Caption         =   "Add"
            Height          =   375
            Index           =   2
            Left            =   420
            TabIndex        =   13
            Top             =   3720
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvFeeBandList 
            Height          =   3135
            Index           =   2
            Left            =   360
            TabIndex        =   16
            Top             =   420
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   5530
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Fee Bands for Date Range"
         Height          =   4335
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   7455
         Begin VB.CommandButton cmdAddFee 
            Caption         =   "Add"
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   5
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdEditFee 
            Caption         =   "Edit"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   1740
            TabIndex        =   4
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteFee 
            Caption         =   "Delete"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   3
            Top             =   3720
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvFeeBandList 
            Height          =   3135
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   420
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   5530
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin MSGOCX.MSGComboBox cboProduct 
         Height          =   315
         Index           =   4
         Left            =   -67560
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
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
      Begin MSGOCX.MSGComboBox cboProduct 
         Height          =   315
         Index           =   2
         Left            =   -73080
         TabIndex        =   41
         Top             =   1320
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSGOCX.MSGComboBox cboPeriod 
         Height          =   315
         Index           =   2
         Left            =   -73080
         TabIndex        =   43
         Top             =   1680
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSGOCX.MSGComboBox cboPeriod 
         Height          =   315
         Index           =   4
         Left            =   -72840
         TabIndex        =   44
         Top             =   960
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGComboBox cboPeriod 
         Height          =   315
         Index           =   3
         Left            =   -72480
         TabIndex        =   45
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSGOCX.MSGComboBox cboProduct 
         Height          =   315
         Index           =   1
         Left            =   7320
         TabIndex        =   46
         Tag             =   "Dummy"
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
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
      Begin MSGOCX.MSGComboBox cboMortgageLender 
         Height          =   315
         Index           =   1
         Left            =   7320
         TabIndex        =   47
         Top             =   5880
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
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
      Begin MSGOCX.MSGComboBox cboMortgageLender 
         Height          =   315
         Index           =   3
         Left            =   -69720
         TabIndex        =   52
         Top             =   960
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
      Begin MSGOCX.MSGComboBox cboMortgageLender 
         Height          =   315
         Index           =   4
         Left            =   -68040
         TabIndex        =   53
         Top             =   960
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
      Begin VB.Label lblProcuration 
         Caption         =   "Packaging Product Period"
         Height          =   375
         Index           =   6
         Left            =   -74520
         TabIndex        =   35
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblProcuration 
         Caption         =   "Product Period"
         Height          =   315
         Index           =   4
         Left            =   -74340
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblProcuration 
         Caption         =   "Insurance Product"
         Height          =   315
         Index           =   3
         Left            =   -74340
         TabIndex        =   28
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lblProcuration 
         Caption         =   "Product Period"
         Height          =   315
         Index           =   2
         Left            =   -74760
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblProcuration 
         Caption         =   "Mortgage Product"
         Height          =   315
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblProcuration 
         Caption         =   "Mortgage Lender"
         Height          =   315
         Index           =   0
         Left            =   -74760
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblProductFees 
         Caption         =   "Product Period"
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   1020
         Width           =   1695
      End
   End
   Begin MSGOCX.MSGComboBox cboPeriod 
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   48
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
   Begin MSGOCX.MSGComboBox cboProduct 
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   49
      Tag             =   "Dummy"
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
   Begin MSGOCX.MSGComboBox cboMortgageLender 
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   50
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
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
   Begin VB.Label lblLabel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "There are a few invisible dummy controls on this form. Leave them there - they allow us to use control indexes for all controls..."
      Height          =   615
      Left            =   5040
      TabIndex        =   51
      Top             =   6960
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmEditProcurationFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditProcurationFees
' Description   : Allows Procuration fees to be modified for the current
'                 Intermediary.
'
' Note: I've used the terms period, band and split in the following comments
' which correspond to the entities found in the underlying tables
' ProcurationFeeType, ProcurationFeeSplit and ProcurationFeeSplitByIntermed
' (respectively).
'
' There is essentially one tab for each 'type' of fee (mortgage, insurance,
' packaging and 'all') and a period drop-down on each tab (showing configured
' periods for the relevant fee-type). Changing the value of this drop-down will
' refresh the listview found on the current tab with the configured bands
' for that type and period (hi/low targets).
'
' Change history
' Prog      Date        Description
' AA        26/06/01    Created
' STB       11/12/01    SYS2550 - SQL Server support.
' STB       22/04/02    SYS4401 - Hidden product selection combo from packaging
'                                 tab.
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Note: The cancel button has been hidden as it served no purpose.

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'A collection of tab-handler objects.
Private m_colForms As Collection

'A collection of primary keys to identify the current intermediary record.
Private m_colKeys As Collection

'These are the underlying table objects required by this form. They are created
'and deleted by the main intermediaries form but the data is controlled by this
'form (and associated tab-handlers).

'This table holds periods.
Private m_clsProcFeeTypeTable As IntermediaryProcFeeTable

'This table holds fee-bands.
Private m_clsProcFeeSplitTable As IntProcFeeSplitTable

'This table holds band-splits.
Private m_clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable

'This enum maps onto the ValueID used in the combo group, the control indexes
'used on this form and also the indexes of the relevant tab-handler classes in
'the m_colForms collection.
Public Enum ProcFeeTypeEnum
    NonSpecificFee = 1
    MortgageFee
    InsuranceFee
    PackagingFee
End Enum

'A flag to stop _Click events being processed when populating combo controls. As
'part of the population, the <select> entries are selected by default which will
'cause a recursive event cycle. Lender combo causes the product combo to refresh
'and the product combo causes the period combo to refresh.
Private m_bBlockComboRecursion As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboMortgageLender_Click
' Description   : When a lender is selected, all the products for that lender should be
'                 populated in the product combo. This will also have the effect of refreshing
'                 the list of periods (and then the listview itself).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboMortgageLender_Click(Index As Integer)
    On Error GoTo Failed

    'Only process if the control is relevant.
    If (Index = MortgageFee) And (m_bBlockComboRecursion = False) Then
        'Broker the call to the specific tab-handler.
        m_colForms(Index).PopulateProductCombo
    End If

    m_colForms(Index).SetFeeButtonState
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboPeriod_Click
' Description   : Enable the relevant edit-period button and refresh the relevant listview
'                 of fee-bands. Index will correspond to the tab/fee-type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboPeriod_Click(Index As Integer)

    On Error GoTo Failed

    If m_bBlockComboRecursion = False Then
        'Route the call down to the relevant tab-handler to refresh the listview.
        m_colForms(Index).PopulateFeeBandList
    End If
    
    m_colForms(Index).SetFeeButtonState
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboInsuranceProduct_Click
' Description   : The product list should goven the available periods. When it changes, lets
'                 lets re-filter the period list.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboProduct_Click(Index As Integer)

    On Error GoTo Failed

    'Only process the click if the control is relevant.
    If (Index <> NonSpecificFee) And (m_bBlockComboRecursion = False) Then
        'Broker the call down to the specific tab-handler.
        m_colForms(Index).PopulatePeriodCombo
    End If
    
    m_colForms(Index).SetFeeButtonState

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAddFee_Click
' Description   : Shows the band/split details form in add mode, if enough information
'                 is present to link the record to a period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddFee_Click(Index As Integer)
    
    Dim bReturn As Boolean
    
    On Error GoTo Failed
    
    'Ensure the relevant combo items are selected.
    bReturn = cboPeriod(Index).CheckMandatory(True)
        
    If bReturn Then
        'Display the ProcFee details form to add a new band/split.
        AddEditFeeBand Index, False
    Else
        'Warn the user if more information is needed.
        cboPeriod(Index).SetFocus
        g_clsErrorHandling.RaiseError errMandatoryFieldsRequired
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAddPeriod_Click
' Description   : Displays a popup to add a new period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddPeriod_Click(Index As Integer)
    
    On Error GoTo Failed
    
    'Displays a popup to add a new period.
    AddEditPeriod Index, False

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdDeleteFee_Click
' Description   : Delete the selected item from the underlying table and listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDeleteFee_Click(Index As Integer)
    
    On Error GoTo Failed
    
    'Broker the call to the relevant tab-handler.
    m_colForms(Index).DeleteFee
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEditFee_Click
' Description   : Edit the selected fee-band.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditFee_Click(Index As Integer)
    
    'Display the ProcFee details form to add a new band/split.
    AddEditFeeBand Index, True
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEditPeriod_Click
' Description   : Displays a popup to edit an existing period.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditPeriod_Click(Index As Integer)
    
    On Error GoTo Failed
    
    'Display a popup to edit an existing period.
    AddEditPeriod Index, True

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
    
    On Error GoTo Failed

    'Ensure all mandatory fields have been populated.
    'The following line is stubbed as the period combos are marked as mandatory
    'but are only validated when a fee is added.
    'bValid = g_clsFormProcessing.DoMandatoryProcessing(Me)
    bValid = True

    'If they have then proceed.
    If bValid = True Then
        'Validate all captured input.
        bValid = ValidateScreenData()
        
        'If the data was valid.
        If bValid = True Then
            'Save the data.
            SaveScreenData
            
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
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate combos and static lists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create the underlying table objects, build the tabs/tab-handlers and
'                 setup the controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Set any underlying state-related table information.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Build the tab-handlers.
    InitialiseTabs
    
    'Populate the combos and static lists.
    PopulateScreenControls
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup tables with new records and GUIDs.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup tables and load data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseTabs
' Description   : Populate m_colforms with all tab-handlers needed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitialiseTabs()
    
    Dim iThisTab As Integer
    Dim clsTabHandler As IntProcFeeTabHandler
    
    On Error GoTo Failed

    'Create a collection to hold the tab-handler objects.
    Set m_colForms = New Collection
        
    'Prepare each tab-handler.
    For iThisTab = 1 To SSTab1.Tabs
        Set clsTabHandler = New IntProcFeeTabHandler
        
        'Add a tab-handler for each tab.
        m_colForms.Add clsTabHandler
    
        'Associate the keys collection with the tab-handler.
        clsTabHandler.SetKeys m_colKeys
        
        'Associate this form with the handler.
        clsTabHandler.SetForm Me
        
        'Set the tab-handler's fee-type (we'll map the index onto the enum).
        clsTabHandler.SetProcFeeType iThisTab
        
        'Associate the table references with the tab-handler.
        clsTabHandler.SetProcFeeTypeTable m_clsProcFeeTypeTable
        clsTabHandler.SetProcFeeSplitTable m_clsProcFeeSplitTable
        clsTabHandler.SetProcFeeSplitByIntTable m_clsProcFeeSplitByIntTable
        
        'Initialise each tab-handler.
        clsTabHandler.Initialise m_bIsEdit
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddEditPeriod
' Description   : Display the period dialog either in add or edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddEditPeriod(ByVal uFeeType As ProcFeeTypeEnum, ByVal bIsEdit As Boolean)
            
    Dim vTmp As Variant
    Dim frmProdPeriod As frmEditProductPeriod
    
    On Error GoTo Failed
    
    'Create a dialog form to add/edit a period.
    Set frmProdPeriod = New frmEditProductPeriod
    
    'Set the key fields on the form.
    frmProdPeriod.SetProcFeeType uFeeType
    frmProdPeriod.SetKeys m_colKeys
    
    'Set the OrganisationID if this fee-type is relevant.
    If uFeeType = MortgageFee Then
        'Convert the LenderCode (used by the combo) into an Organisation ID
        'and set it against the form.
        g_clsFormProcessing.HandleComboExtra cboMortgageLender(uFeeType), vTmp, GET_CONTROL_VALUE
        frmProdPeriod.SetOrganisationID m_colForms(uFeeType).GetOrgIDFromLenderCode(CStr(vTmp))
    End If
    
    'Set the ProductID if this fee-type is relevant.
    Select Case uFeeType
        Case InsuranceFee, MortgageFee
            g_clsFormProcessing.HandleComboExtra cboProduct(uFeeType), vTmp, GET_CONTROL_VALUE
            frmProdPeriod.SetProductID CStr(vTmp)
    End Select
        
    If bIsEdit Then
        'If we're editing then set the Sequence number.
        frmProdPeriod.SetTypeSequence cboPeriod(uFeeType).GetExtra(cboPeriod(uFeeType).ListIndex)
    End If
    
    'Set the add/edit state of the form.
    frmProdPeriod.SetIsEdit bIsEdit
        
    'Associate the underlying period table.
    frmProdPeriod.SetProcFeeTypeTable m_clsProcFeeTypeTable
    
    'Show the form modally.
    frmProdPeriod.Show vbModal

    'Refresh the periods so that the edited/new record is included.
    m_colForms(uFeeType).PopulatePeriodCombo
    
    'Select the edited/added item.
    g_clsFormProcessing.HandleComboExtra cboPeriod(uFeeType), CVar(frmProdPeriod.GetTypeSequence), SET_CONTROL_VALUE
    
    'Refresh the button state.
    m_colForms(uFeeType).SetFeeButtonState
    
    'Ensure the period pop-up is unloaded.
    Unload frmProdPeriod
    Set frmProdPeriod = Nothing
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddEditFeeBand
' Description   : Displays the band/split details form in add or edit mode for the given
'                 fee-type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddEditFeeBand(ByVal uFeeType As ProcFeeTypeEnum, ByVal bIsEdit As Boolean)
            
    Dim vProductID As Variant
    Dim vTypeSequence As Variant
    Dim clsPopulateDetails As PopulateDetails
    Dim frmFeeDetails As frmEditProductProcFee
    
    On Error GoTo Failed
    
    'Create a proc fee details form.
    Set frmFeeDetails = New frmEditProductProcFee
    
    'Set the key fields on the form.
    frmFeeDetails.SetFeeType uFeeType
    frmFeeDetails.SetKeys m_colKeys
    
    'Type Sequence.
    g_clsFormProcessing.HandleComboExtra cboPeriod(uFeeType), vTypeSequence, GET_CONTROL_VALUE
    frmFeeDetails.SetTypeSequence CStr(vTypeSequence)
        
    If bIsEdit Then
        Set clsPopulateDetails = lvFeeBandList(uFeeType).SelectedItem.Tag
        
        'If we're editing, then set the SplitSequence portion of the key.
        frmFeeDetails.SetSplitSequence clsPopulateDetails.GetExtra
    End If
    
    'Set the add/edit state.
    frmFeeDetails.SetIsEdit bIsEdit
    
    'Associate the required tables.
    frmFeeDetails.SetProcFeeTypeTable m_clsProcFeeTypeTable
    frmFeeDetails.SetProcFeeSplitTable m_clsProcFeeSplitTable
    frmFeeDetails.SetProcFeeSplitByIntTable m_clsProcFeeSplitByIntTable
    
    'The product ID will goven a caption displayed on the split details form.
    Select Case uFeeType
        Case InsuranceFee, MortgageFee
            g_clsFormProcessing.HandleComboExtra cboProduct(uFeeType), vProductID, GET_CONTROL_VALUE
            frmFeeDetails.SetProductID CStr(vProductID)
        
        Case Else
            frmFeeDetails.SetProductID ""
    End Select
    
    'Show the form modally.
    frmFeeDetails.Show vbModal

    'Refresh the periods so that the edited/new record is included.
    m_colForms(uFeeType).PopulateFeeBandList
    
    'Refresh button state.
    m_colForms(uFeeType).SetFeeButtonState
    
    'Ensure the period pop-up is unloaded.
    Unload frmFeeDetails
    Set frmFeeDetails = Nothing
    Exit Sub
    
Failed:
    Set frmFeeDetails = Nothing
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Tidy-up object references
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)

    Dim oTabHandler As Object
    
    'Ensure cyclic references are severed.
    For Each oTabHandler In m_colForms
        oTabHandler.Terminate
    Next oTabHandler

    'Before releasing the object references, clear any filters on the tables.
    TableAccess(m_clsProcFeeTypeTable).ApplyFilter "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    TableAccess(m_clsProcFeeSplitTable).ApplyFilter "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))
    TableAccess(m_clsProcFeeSplitByIntTable).ApplyFilter "IntermediaryGUID = " & g_clsSQLAssistSP.FormatString(g_clsSQLAssistSP.ByteArrayToGuidString(CStr(m_colKeys(INTERMEDIARY_KEY))))

    'Release object references.
    Set m_colKeys = Nothing
    Set m_colForms = Nothing
    Set m_clsProcFeeTypeTable = Nothing
    Set m_clsProcFeeSplitTable = Nothing
    Set m_clsProcFeeSplitByIntTable = Nothing

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvFeeBandList_DblClick
' Description   : Do the same processing as-if the edit button was clicked.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvFeeBandList_DblClick(Index As Integer)
    cmdEditFee(Index) = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvFeeBandList_MouseUp
' Description   : Enable/disable the relevant edit/delete buttons.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvFeeBandList_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    m_colForms(Index).SetFeeButtonState
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SSTab1_Click
' Description   : Route the call off to a global bug-fix routine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SSTab1_Click(PreviousTab As Integer)
    
    On Error GoTo Failed
    
    SetTabstops Me
        
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetReturnCode
' Description   :   Sets the return code for the form for any calling method to check. Defaults
'                   to MSG_SUCCESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal uReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = uReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the success code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeType
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeType(ByRef clsProcFeeTypeTable As IntermediaryProcFeeTable)
    Set m_clsProcFeeTypeTable = clsProcFeeTypeTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplit
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplit(ByRef clsProcFeeSplitTable As IntProcFeeSplitTable)
    Set m_clsProcFeeSplitTable = clsProcFeeSplitTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcFeeSplitByInt
' Description   : Associate the specified table with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProcFeeSplitByInt(ByRef clsProcFeeSplitByIntTable As IntProcFeeSplitForIntTable)
    Set m_clsProcFeeSplitByIntTable = clsProcFeeSplitByIntTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Iterates through each tab-handler and asks it to validate the
'                 controls which are relevant to it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    'Stub.
    ValidateScreenData = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Broker the call onto each tab-handler. Saving its data to the underlying
'                 table object(s).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : BlockComboRecursion
' Description   : When this flag is true, the combo _Click events will not be processed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BlockComboRecursion(ByVal bBlock As Boolean)
    m_bBlockComboRecursion = bBlock
End Sub


