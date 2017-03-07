VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditTasks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Task"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "frmEditTasks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   7680
      TabIndex        =   25
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   7800
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7650
      Left            =   75
      TabIndex        =   26
      Top             =   75
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   13494
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Task Details"
      TabPicture(0)   =   "frmEditTasks.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTASQueueName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblInterface"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblRuleRef"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblOwnerType"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblOutputDocument"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTaskType"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTaskName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTaskId"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblDocumentPackId"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblWhenRuleRuns"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblInputProcess"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtRuleRef"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboRuleFrequency"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtTask(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtTask(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTask(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboOwnerType"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTask(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboTaskType"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTask(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtTask(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cboTASQueueName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkTriggerNxtStage"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkTAS"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkAlwaysAuto"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkAutomaticTask"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkRemoteOwnerTask"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkProgressTask"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkEditableTask"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "chkNotApplicable"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "chkApplicantTask"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "chkCustomerInvolved"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkCustomerTask"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkConfirmPrintInd"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Task Timings"
      TabPicture(1)   =   "frmEditTasks.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblChasingPeriod"
      Tab(1).Control(1)=   "lblChasingTask"
      Tab(1).Control(2)=   "lblAdjustments"
      Tab(1).Control(3)=   "lblContactType"
      Tab(1).Control(4)=   "cboContactType"
      Tab(1).Control(5)=   "txtChasingPeriod"
      Tab(1).Control(6)=   "cboChasingTask"
      Tab(1).Control(7)=   "dgAdjustments"
      Tab(1).Control(8)=   "cboChasingPeriodTimescale"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Interface Task Details"
      TabPicture(2)   =   "frmEditTasks.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInitialiseAsComplete"
      Tab(2).Control(1)=   "cmdDeleteCode"
      Tab(2).Control(2)=   "cmdAddMessageCode"
      Tab(2).Control(3)=   "lstSelectedCodes"
      Tab(2).Control(4)=   "cboMessageSubType"
      Tab(2).Control(5)=   "cboMessageType"
      Tab(2).Control(6)=   "cboInterfaceType"
      Tab(2).Control(7)=   "lblShowMessageCodes"
      Tab(2).Control(8)=   "lblMessageSubType"
      Tab(2).Control(9)=   "lblMessageType"
      Tab(2).Control(10)=   "lblInitialiseAsComplete"
      Tab(2).Control(11)=   "lblInterfaceType"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Linked Tasks"
      TabPicture(3)   =   "frmEditTasks.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdDeleteLinkedTask"
      Tab(3).Control(1)=   "cmdAddLinkedTask"
      Tab(3).Control(2)=   "cboLinkTask"
      Tab(3).Control(3)=   "lvLinkedTasks"
      Tab(3).Control(4)=   "lblLinkedTasks"
      Tab(3).Control(5)=   "lblLinkedTask"
      Tab(3).ControlCount=   6
      Begin VB.CheckBox chkConfirmPrintInd 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Print?"
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   5880
         Width           =   2730
      End
      Begin VB.CheckBox chkCustomerTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Task?"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   6780
         Width           =   2730
      End
      Begin VB.CheckBox chkCustomerInvolved 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Contact Involved?"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6360
         Width           =   2730
      End
      Begin VB.CheckBox chkApplicantTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Applicant Task"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   5460
         Width           =   2730
      End
      Begin VB.CheckBox chkNotApplicable 
         Alignment       =   1  'Right Justify
         Caption         =   "'Not Applicable' status valid?"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   2730
      End
      Begin VB.CheckBox chkEditableTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Editable Task"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   4620
         Width           =   2730
      End
      Begin VB.CheckBox chkProgressTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Progress Task"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   5040
         Width           =   2730
      End
      Begin VB.CheckBox chkRemoteOwnerTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote Owner Task?"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   7200
         Value           =   1  'Checked
         Width           =   2730
      End
      Begin VB.Frame fraInitialiseAsComplete 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -72960
         TabIndex        =   64
         Top             =   1770
         Width           =   1575
         Begin VB.OptionButton optInitAsCompYes 
            Caption         =   "Yes"
            Height          =   375
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optInitAsCompNo 
            Caption         =   "No"
            Height          =   375
            Left            =   720
            TabIndex        =   65
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdDeleteCode 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   -68235
         TabIndex        =   54
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddMessageCode 
         Caption         =   "Add Code"
         Height          =   375
         Left            =   -72960
         TabIndex        =   53
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteLinkedTask 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -67635
         TabIndex        =   49
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddLinkedTask 
         Caption         =   "Add"
         Height          =   375
         Left            =   -67635
         TabIndex        =   48
         Top             =   750
         Width           =   1215
      End
      Begin VB.ComboBox cboLinkTask 
         Height          =   315
         Left            =   -74760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   780
         Width           =   6975
      End
      Begin MSGOCX.MSGComboBox cboChasingPeriodTimescale 
         Height          =   315
         Left            =   -72240
         TabIndex        =   38
         Top             =   615
         Width           =   1455
         _ExtentX        =   2566
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
         ListText        =   "days|hours|minutes"
         Text            =   ""
      End
      Begin VB.CheckBox chkAutomaticTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Automatic On Stage Entry"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   4200
         Width           =   3615
      End
      Begin VB.CheckBox chkAlwaysAuto 
         Alignment       =   1  'Right Justify
         Caption         =   "Automatic On Ad Hoc Creation"
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   4620
         Width           =   3615
      End
      Begin VB.CheckBox chkTAS 
         Alignment       =   1  'Right Justify
         Caption         =   "Task Automation Service (TAS)"
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   5040
         Width           =   3615
      End
      Begin VB.CheckBox chkTriggerNxtStage 
         Alignment       =   1  'Right Justify
         Caption         =   "TAS trigger next stage when completed last"
         Height          =   195
         Left            =   4800
         TabIndex        =   21
         Top             =   5490
         Width           =   3615
      End
      Begin MSGOCX.MSGComboBox cboTASQueueName 
         Height          =   315
         Left            =   6480
         TabIndex        =   22
         Top             =   5880
         Width           =   1970
         _ExtentX        =   3466
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
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   0
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
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
         MaxLength       =   20
      End
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Top             =   885
         Width           =   5650
         _ExtentX        =   9975
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
         MaxLength       =   50
      End
      Begin MSGOCX.MSGComboBox cboTaskType 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   1290
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
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   5
         Left            =   2760
         TabIndex        =   4
         Top             =   2100
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSGOCX.MSGComboBox cboOwnerType 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   2910
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
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Top             =   2505
         Width           =   5650
         _ExtentX        =   9975
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
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   6
         Left            =   6475
         TabIndex        =   5
         Top             =   2100
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSGOCX.MSGDataGrid dgAdjustments 
         Height          =   1575
         Left            =   -73260
         TabIndex        =   40
         Top             =   1095
         Width           =   5790
         _ExtentX        =   10213
         _ExtentY        =   2778
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
      Begin MSGOCX.MSGComboBox cboChasingTask 
         Height          =   315
         Left            =   -73200
         TabIndex        =   42
         Top             =   2910
         Width           =   5595
         _ExtentX        =   9869
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
      Begin MSGOCX.MSGEditBox txtChasingPeriod 
         Height          =   315
         Left            =   -73200
         TabIndex        =   37
         Top             =   615
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSGOCX.MSGComboBox cboContactType 
         Height          =   315
         Left            =   -73200
         TabIndex        =   43
         Top             =   3420
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSGOCX.MSGEditBox txtTask 
         Height          =   315
         Index           =   4
         Left            =   2760
         TabIndex        =   3
         Top             =   1695
         Width           =   1515
         _ExtentX        =   2672
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
      Begin MSGOCX.MSGListView lvLinkedTasks 
         Height          =   5655
         Left            =   -74800
         TabIndex        =   50
         Top             =   1620
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9975
         HideSelection   =   0   'False
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
      Begin MSGOCX.MSGListView lstSelectedCodes 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   55
         Top             =   3240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2990
         HideSelection   =   0   'False
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
      Begin MSGOCX.MSGComboBox cboMessageSubType 
         Height          =   315
         Left            =   -72960
         TabIndex        =   56
         Top             =   1395
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
      Begin MSGOCX.MSGComboBox cboMessageType 
         Height          =   315
         Left            =   -72960
         TabIndex        =   57
         Top             =   1005
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
      Begin MSGOCX.MSGComboBox cboInterfaceType 
         Height          =   315
         Left            =   -72960
         TabIndex        =   58
         Top             =   600
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
      Begin MSGOCX.MSGComboBox cboRuleFrequency 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   3315
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
         ListText        =   "Never|On stage entry (default tasks)|When task automatically generated"
         Text            =   ""
      End
      Begin MSGOCX.MSGEditBox txtRuleRef 
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   3720
         Width           =   5655
         _ExtentX        =   9975
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
         MaxLength       =   50
      End
      Begin VB.Label lblShowMessageCodes 
         AutoSize        =   -1  'True
         Caption         =   "Selected Message Codes"
         Height          =   195
         Left            =   -74715
         TabIndex        =   63
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblMessageSubType 
         AutoSize        =   -1  'True
         Caption         =   "Message SubType"
         Height          =   195
         Left            =   -74715
         TabIndex        =   62
         Top             =   1460
         Width           =   1335
      End
      Begin VB.Label lblMessageType 
         AutoSize        =   -1  'True
         Caption         =   "Message Type"
         Height          =   195
         Left            =   -74715
         TabIndex        =   61
         Top             =   1060
         Width           =   1050
      End
      Begin VB.Label lblInitialiseAsComplete 
         AutoSize        =   -1  'True
         Caption         =   "Initialise as Complete?"
         Height          =   195
         Left            =   -74715
         TabIndex        =   60
         Top             =   1860
         Width           =   1560
      End
      Begin VB.Label lblInterfaceType 
         AutoSize        =   -1  'True
         Caption         =   "Interface Type"
         Height          =   195
         Left            =   -74715
         TabIndex        =   59
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lblLinkedTasks 
         AutoSize        =   -1  'True
         Caption         =   "Linked Tasks"
         Height          =   195
         Left            =   -74760
         TabIndex        =   52
         Top             =   1380
         Width           =   960
      End
      Begin VB.Label lblLinkedTask 
         AutoSize        =   -1  'True
         Caption         =   "Task to link"
         Height          =   195
         Left            =   -74760
         TabIndex        =   51
         Top             =   480
         Width           =   825
      End
      Begin VB.Label lblInputProcess 
         AutoSize        =   -1  'True
         Caption         =   "Input Process"
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label lblContactType 
         AutoSize        =   -1  'True
         Caption         =   "Contact Type"
         Height          =   195
         Left            =   -74760
         TabIndex        =   45
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label lblAdjustments 
         AutoSize        =   -1  'True
         Caption         =   "Adjustments"
         Height          =   195
         Left            =   -74760
         TabIndex        =   44
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label lblChasingTask 
         AutoSize        =   -1  'True
         Caption         =   "Chasing Task"
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   2970
         Width           =   975
      End
      Begin VB.Label lblChasingPeriod 
         AutoSize        =   -1  'True
         Caption         =   "Chasing Period"
         Height          =   195
         Left            =   -74760
         TabIndex        =   39
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label lblWhenRuleRuns 
         AutoSize        =   -1  'True
         Caption         =   "When Rule Runs"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   3375
         Width           =   1230
      End
      Begin VB.Label lblDocumentPackId 
         AutoSize        =   -1  'True
         Caption         =   "Document Pack Id"
         Height          =   195
         Left            =   4800
         TabIndex        =   35
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblTaskId 
         AutoSize        =   -1  'True
         Caption         =   "Task ID"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   540
         Width           =   570
      End
      Begin VB.Label lblTaskName 
         AutoSize        =   -1  'True
         Caption         =   "Task Name"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   945
         Width           =   825
      End
      Begin VB.Label lblTaskType 
         AutoSize        =   -1  'True
         Caption         =   "Task Type"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label lblOutputDocument 
         AutoSize        =   -1  'True
         Caption         =   "Output Document"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label lblOwnerType 
         AutoSize        =   -1  'True
         Caption         =   "Owner Type"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2970
         Width           =   870
      End
      Begin VB.Label lblRuleRef 
         AutoSize        =   -1  'True
         Caption         =   "Rule Reference"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3780
         Width           =   1125
      End
      Begin VB.Label lblInterface 
         AutoSize        =   -1  'True
         Caption         =   "Interface"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2565
         Width           =   630
      End
      Begin VB.Label lblTASQueueName 
         AutoSize        =   -1  'True
         Caption         =   "TAS Queue Name"
         Height          =   195
         Left            =   4830
         TabIndex        =   27
         Top             =   5940
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmEditTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditTasks
' Description   :   Form which allows the adding/editing of tasks for Task Management

' Change history
' Prog      Date        Description
' DJP       09/11/00    Created for Phase 2 Task Management
' AA        01/03/01    AQR SYS1893 - Added dgAdjustments Datagrid/Functionality
' AA        05/03/01    AQR SYS1823 - Added Interface Field.
' AA        21/03/01    Changed cboContactType ComboGroup from ThirdPartyType to
'                       TaskContactType
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS History
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' AW    13/05/02    BM092   Added CheckPrint
' GD MERGED for BMIDS 24/05/02 STB       01/04/02    SYS4509 - When toggling day/hour adjustment, don't hold the unused value.
' GD MERGED for BMIDS 24/05/02 STB       08/07/2002  SYS4529 BorderStyle set to 'Fixed Dialog'.
' GD    22/07/02  BMIDS00229 - ensure 'Not applicable status valid?' gets saved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JD    03/09/2005  MAR32 WP8 added new automatic task functionality
' JD    04/09/2005  MAR40 WP2 add interface task details tab
' GHun  20/10/2005  MAR244 change form size and field layout
' MV    20/10/2005  MAR251  Amended SaveScreenData
' GHun  21/10/2005  MAR244 reduced form height
' GHun  15/11/2005  MAR516 Fix saving of interface task details
' HMA   22/11/2005  MAR670 Add Case Assessment to Interface task details
' RF    13/12/2005  MAR202 Handle packs
' PSC   28/02/2006  MAR1341 Add chasing period minutes
' GHun  07/03/2006  MAR1300 Add Linked Tasks tab
' HMA   18/04/2006  MAR1545 Update task name in TaskLink table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change History:
'
' Prog      Date        Ref     Description
' PB        24/04/2006  EP367   Changes merged into Epsom Supervisor
' PB        10/05/2006  EP511   Merged in fix from MARS: MAR1669
' AW        31/08/2006  EP1110  Added EDITABLETASKIND, PROGRESSTASKIND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM2 Change History:
'
' Prog      Date        Ref     Description
' PE        26/10/2006  EP2_24  Added Message type and sub type combos for the Xit2 interface.
' TW        12/12/2006  EP2_453 - E2CR63 Remote Underwriting add Remote Owner Task (Checkbox; default = checked)
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' GHun      27/02/2007  EP2_1363 - Added Email Notification Interface Type
' GHun      02/10/2007  OMIGA00003234 Added RuleFrequency and improved field layout
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_clsTaskPriority As TaskPriority
Private m_clsTaskTable As TaskTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_sTaskID As String
'Private m_bChasingDays As Boolean
'JD MAR40
Private m_clsTaskInterfaceMessage As TaskInterfaceMessageTable
Private m_clsTaskInterfaceSubMessage As TaskInterfaceSubMessageTable

Private m_sMessageTypeComboName As String
Private m_sMessageSubTypeComboName As String

Private m_clsLinkedTaskTable As LinkedTaskTable 'MAR1300 GHun
Private m_clsTaskLinkTable As TaskLinkTable     'MAR1545

' Private constants
Private Const TASK_ID = 0
Private Const TASK_NAME = 1
'Private Const CHASING_PERIOD = 2
Private Const TASK_INTERFACE = 3
Private Const INPUT_PROCESS = 4
Private Const OUTPUT_DOCUMENT = 5
Private Const PACK_CONTROL_ID = 6 ' RF MAR202
'Private Const ADJUSTMENT_DAYS_COL = 4
'Private Const ADJUSTMENT_HOURS_COL = 3
Private Const CHASING_PERIOD_DAYS_FIELD As String = "ADJUSTMENTDAYS"
Private Const CHASING_PERIOD_HOURS_FIELD As String = "ADJUSTMENTHOURS"
' PSC 28/02/2006 MAR1341
Private Const CHASING_PERIOD_MINUTES_FIELD As String = "ADJUSTMENTMINUTES"

'OMIGA00003234 GHun
Private Enum ChasingPeriodTimescale
    CPT_DAYS = 0
    CPT_HOURS = 1
    CPT_MINUTES = 2
End Enum

Private Enum RuleFrequency
    RF_NEVER = 0
    RF_STAGE_ENTRY = 1
    RF_ALWAYS = 2
End Enum
'OMIGA00003234 End

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Private Sub cboInterfaceType_Click()
    Dim vInterfaceType As Variant
    Dim colComboList As Collection
    
    'On setting the interface type, set the correct messageType combo
    g_clsFormProcessing.HandleComboExtra cboInterfaceType, vInterfaceType, GET_CONTROL_VALUE
    
    'COMBO_NONE
    If Len(vInterfaceType) = 0 Then
        If cboMessageType.Enabled Then  'MAR516 GHun
            g_clsFormProcessing.HandleComboExtra cboMessageType, CStr(COMBO_NONE), SET_CONTROL_VALUE
        End If
        If cboMessageSubType.Enabled Then   'MAR516 GHun
            g_clsFormProcessing.HandleComboExtra cboMessageSubType, CStr(COMBO_NONE), SET_CONTROL_VALUE
        End If
        cboMessageType.Enabled = False
        cboMessageSubType.Enabled = False
        cmdAddMessageCode.Enabled = False
    Else
        Set colComboList = GetMessageTypeCombo(vInterfaceType)
        m_sMessageTypeComboName = colComboList.Item(1)
        m_sMessageSubTypeComboName = colComboList.Item(2)
        
        If Len(m_sMessageTypeComboName) > 0 Then
            g_clsFormProcessing.PopulateCombo m_sMessageTypeComboName, cboMessageType
            g_clsFormProcessing.HandleComboExtra cboMessageType, CStr(COMBO_NONE), SET_CONTROL_VALUE
            cboMessageType.Enabled = True
        End If
        
        If Len(m_sMessageSubTypeComboName) > 0 Then
            g_clsFormProcessing.PopulateCombo m_sMessageSubTypeComboName, cboMessageSubType
            g_clsFormProcessing.HandleComboExtra cboMessageSubType, CStr(COMBO_NONE), SET_CONTROL_VALUE
        End If
    End If
    
End Sub

Private Sub cboMessageSubType_Click()
On Error GoTo Failed
    Dim vMessageSubType As Variant
    
    g_clsFormProcessing.HandleComboExtra cboMessageSubType, vMessageSubType, GET_CONTROL_VALUE
    If Len(vMessageSubType) = 0 Then
        cmdAddMessageCode.Enabled = False
    Else
        cmdAddMessageCode.Enabled = True
    End If
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cboMessageType_Click()
'The messageType may allow subtypes. Look at the validation type on the MessageType
'and see if they match any on the subTypes combo

    On Error GoTo Failed
    Dim vMessageType As Variant
    Dim sValidation As String
    Dim clsComboValidation As ComboValidationTable
    Dim nValueID As Long
    
    g_clsFormProcessing.HandleComboExtra cboMessageType, vMessageType, GET_CONTROL_VALUE
    
    If Len(vMessageType) = 0 Then
        cboMessageSubType.Enabled = False
        cmdAddMessageCode.Enabled = False
    Else
        cmdAddMessageCode.Enabled = True
        
        'If there are sub types for this interface type see if the message type matches
        If Len(m_sMessageSubTypeComboName) > 0 Then
            g_clsFormProcessing.GetComboValidation cboMessageType, m_sMessageTypeComboName, sValidation
            
            Set clsComboValidation = New ComboValidationTable
            nValueID = clsComboValidation.GetValueIDForValidationType(sValidation, m_sMessageSubTypeComboName)
            If Not nValueID = 0 Then
                'our messagesubtypes match the messagetype so enable the combo.
                'Don't allow user to add unless a subtype is chosen
                cboMessageSubType.Enabled = True
                cmdAddMessageCode.Enabled = False
            Else
                g_clsFormProcessing.HandleComboExtra cboMessageSubType, CStr(COMBO_NONE), SET_CONTROL_VALUE
                cboMessageSubType.Enabled = False
                cmdAddMessageCode.Enabled = True
            End If
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'OMIGA00003234 GHun
Private Sub cboRuleFrequency_Click()
    If cboRuleFrequency.ListIndex = RF_NEVER Then
        txtRuleRef.Enabled = False
        lblRuleRef.Enabled = False
        txtRuleRef.Mandatory = False
    Else
        txtRuleRef.Enabled = True
        lblRuleRef.Enabled = True
        txtRuleRef.Mandatory = True
    End If
End Sub
'OMIGA00003234 End

Private Sub chkTAS_Click()
'JD MAR32 OnSelectTASCheckBox
        
    If CBool(chkTAS.Value) = True Then
        cboTASQueueName.Enabled = True
        chkTriggerNxtStage.Enabled = True
        lblTASQueueName.Enabled = True
    Else
        g_clsFormProcessing.HandleComboExtra cboTASQueueName, CStr(COMBO_NONE), SET_CONTROL_VALUE
        cboTASQueueName.Enabled = False
        chkTriggerNxtStage.Value = 0
        chkTriggerNxtStage.Enabled = False
        lblTASQueueName.Enabled = False
    End If
End Sub
Private Function LineAlreadyAdded(colInList As Collection) As Boolean
On Error GoTo Failed

    Dim lstCount As Integer
    Dim colGetList As Collection
    Dim colGetValueIDs As Collection
    Dim bAlreadyAdded As Boolean
    
    bAlreadyAdded = False
    
    For lstCount = 1 To lstSelectedCodes.ListItems.Count
        Set colGetList = lstSelectedCodes.GetLine(lstCount, colGetValueIDs)
        If colInList.Count = colGetList.Count Then
            
            If colInList.Item(1) = colGetList.Item(1) And _
               colInList.Item(2) = colGetList.Item(2) Then
               
                If colInList.Count = 3 Then
                    If colInList.Item(3) = colGetList.Item(3) Then
                        bAlreadyAdded = True
                    End If
                Else
                    bAlreadyAdded = True
                End If
            End If
        End If
    Next
    
    LineAlreadyAdded = bAlreadyAdded

Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'MAR1300 GHun
Private Sub cmdAddLinkedTask_Click()
    Dim strTaskId As String
    Dim strLinkedTaskId As String
    Dim strTaskName As String
    Dim intPos As String
    Dim colLine As Collection
    Dim colValues As Collection
    Dim clsPopulateDetails As PopulateDetails
    
    If cboLinkTask.ListIndex > 0 Then
        strTaskId = txtTask.Item(0).Text
        intPos = InStr(1, cboLinkTask.Text, ": ")
        
        strLinkedTaskId = Left$(cboLinkTask.Text, intPos - 1)
        strTaskName = Trim$(Right$(cboLinkTask.Text, Len(cboLinkTask.Text) - intPos - 1))
            
        Set colLine = New Collection
        colLine.Add strLinkedTaskId
        colLine.Add strTaskName
        
        Set colValues = New Collection
        colValues.Add strTaskId
        colValues.Add strLinkedTaskId
         
        'TableAccess(m_clsLinkedTaskTable).SetKeyMatchValues colValues
        
        TableAccess(m_clsLinkedTaskTable).AddRow
        m_clsLinkedTaskTable.SetTaskID strTaskId
        m_clsLinkedTaskTable.SetLinkedTaskId strLinkedTaskId
        
        Set clsPopulateDetails = New PopulateDetails
        clsPopulateDetails.SetKeyMatchValues colValues
        'clsPopulateDetails.SetExtra ""  'sExtraValue
        'clsPopulateDetails.SetObjectDescription ""  'sObjectDescription
        
        lvLinkedTasks.AddLine colLine, clsPopulateDetails
        
        cboLinkTask.RemoveItem cboLinkTask.ListIndex
        cboLinkTask.ListIndex = 0
            
    End If
    
    Set colLine = Nothing
    Set colValues = Nothing
    Set clsPopulateDetails = Nothing
End Sub
'MAR1300 End

Private Sub cmdAddMessageCode_Click()
'Adds a row to the list box
On Error GoTo Failed
    Dim colList As Collection
    Dim colValueIDs As Collection
    Dim vValueID As Variant
    Dim sCode As String
    
    Set colList = New Collection
    Set colValueIDs = New Collection
    
    g_clsFormProcessing.HandleComboExtra cboInterfaceType, vValueID, GET_CONTROL_VALUE
    sCode = g_clsCombo.GetValueNameFromValueID(CStr(vValueID), "InterfaceType")
    colList.Add sCode
    colValueIDs.Add vValueID
    
    g_clsFormProcessing.HandleComboExtra cboMessageType, vValueID, GET_CONTROL_VALUE
    sCode = g_clsCombo.GetValueNameFromValueID(CStr(vValueID), m_sMessageTypeComboName)
    colList.Add sCode
    colValueIDs.Add vValueID
    
    g_clsFormProcessing.HandleComboExtra cboMessageSubType, vValueID, GET_CONTROL_VALUE
    If Len(vValueID) > 0 Then
        sCode = g_clsCombo.GetValueNameFromValueID(CStr(vValueID), m_sMessageSubTypeComboName)
        colList.Add sCode
        colValueIDs.Add vValueID
    Else
        'add empty values. Necessary because of the check LineAlreadyAdded
        colList.Add vbNullString
        colValueIDs.Add vbNullString
    End If

    If Not LineAlreadyAdded(colList) Then
        lstSelectedCodes.AddLine colList, colValueIDs
    End If
            
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user clicks on the Another button. Need to perform the same
'                   processing as if the user has pressed ok. Do validation, save screen data,
'                   clear the screen of existing values and put the focus on the TASK ID field
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    On Error GoTo Failed
        
    ' Perform ok processing. Will raise an error if it fails
    DoOKProcessing
    
    'If the record was saved, commit the transaction and begin another.
    g_clsDataAccess.CommitTrans
    g_clsDataAccess.BeginTrans
    
    ' Clear all fields, set the default focus
    g_clsFormProcessing.ClearScreenFields Me
    
    SetAddState
    
    txtTask(TASK_ID).SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdDeleteCode_Click()
'remove the selected row from the list
On Error GoTo Failed
    Dim lstItem As ListItem
    Set lstItem = lstSelectedCodes.SelectedItem
    
    ' Remove a row from the listview
    If Not lstItem Is Nothing Then
        lstSelectedCodes.RemoveLine lstItem.Index
    End If

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   VB Calls this method when the form is loaded
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    
    Initialise
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Initialise
' Description   :   Funciton which performs all screen initialisation. Populate all combos, read
'                   data from the database if in add mode, create a new record if in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Initialise()
    On Error GoTo Failed
    
    ' Task table contains all database specific code
    Set m_clsTaskTable = New TaskTable
    m_ReturnCode = MSGFailure
    Set m_clsTaskPriority = New TaskPriority
    'JD MAR32
    Set m_clsTaskInterfaceMessage = New TaskInterfaceMessageTable
    Set m_clsTaskInterfaceSubMessage = New TaskInterfaceSubMessageTable
    Set m_clsLinkedTaskTable = New LinkedTaskTable  'MAR1300 GHun
    Set m_clsTaskLinkTable = New TaskLinkTable      'MAR1545
    
    PopulateCombos
    
    
    ' Edit mode?
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    PopulateChasingTask
    dgAdjustments.Enabled = True
    
    SetTabstops Me  'OMIGA00003234
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateCombos
' Description   :   Populates combo boxes with their values. The values may come from a combo
'                   group (PopulateCombo)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateCombos()
    On Error GoTo Failed
    
    ' Task Type
    g_clsFormProcessing.PopulateCombo "TaskType", cboTaskType
    
    ' Contact Type
    g_clsFormProcessing.PopulateCombo "TaskContactType", cboContactType
    
    ' Owner Type
    g_clsFormProcessing.PopulateCombo "UserRole", cboOwnerType
    
    ' JD MAR32 TAS queue name
    g_clsFormProcessing.PopulateCombo "TaskAutomationQueueName", cboTASQueueName
    
    ' JD MAR40 InterfaceType
    g_clsFormProcessing.PopulateCombo "InterfaceType", cboInterfaceType
    'set messageType and messageSubType combos as disabled.
    cboMessageType.Enabled = False
    cboMessageSubType.Enabled = False
    'set add button disabled
    cmdAddMessageCode.Enabled = False

    Exit Sub
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateChasingTask
' Description   :   Populates the Chasing Task combo box with a list of all Tasks read from the
'                   database (the Task table)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateChasingTask()
    On Error GoTo Failed
    Dim colFields As Collection
    Dim colTasks As Collection
    Dim clsTaskTable As TaskTable
    Dim vTask As Variant
    
    ' Create the Task table access class
    Set clsTaskTable = New TaskTable
    Set colTasks = New Collection
    
    ' Get the list of tasks
    colTasks.Add m_sTaskID
    clsTaskTable.GetTasks colTasks
    
    If TableAccess(clsTaskTable).RecordCount() > 0 Then
        Set colFields = clsTaskTable.GetComboFields()
        g_clsFormProcessing.PopulateComboFromTable cboChasingTask, clsTaskTable, colFields
        
        ' Set Chasing task selection if in edit mode
        If m_bIsEdit Then
            vTask = m_clsTaskTable.GetChasingTask()
            g_clsFormProcessing.HandleComboExtra cboChasingTask, vTask, SET_CONTROL_VALUE
        End If
    Else
        cboChasingTask.Enabled = False
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user cliks the Cancel button. All we do is hide the form, which
'                   passes control back to the caller, which checks the status of the form, and closes
'                   it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    On Error GoTo Failed
    
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    
    Dim bRet As Boolean
    
    bRet = ValidateScreenData
    
    If bRet Then
        SaveScreenData
        SaveChangeRequest
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateScreenData
' Description   :   Perform all screen validation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean
    Dim strPackControlId As String
    
    g_clsFormProcessing.DoMandatoryProcessing Me
    
    If Not dgAdjustments.DataSource Is Nothing Then
        bRet = ValidateAdjustments
    Else
        bRet = True
    End If
          
    'BMIDS00719 Start
    'If task is an Applicant task, ensure a Contact Type has been selected.
    'This is required when setting up adhoc tasks.
    If bRet Then
        Dim vVal As Variant
        g_clsFormProcessing.HandleComboExtra cboContactType, vVal, GET_CONTROL_VALUE
    
        If chkApplicantTask.Value <> 0 And Len(vVal) = 0 Then
            g_clsErrorHandling.DisplayError "An Applicant Task must have a Contact Type selected."
            cboContactType.SetComboFocus
            bRet = False
        End If
    End If
    'BMIDS00719 End
    
    ' RF MAR202 Start
    ' Handle link to packs
    If bRet Then
        strPackControlId = txtTask(PACK_CONTROL_ID).Text
        If Len(strPackControlId) > 0 Then
            ' output doc and pack control id are mutually exclusive
            If Len(txtTask(OUTPUT_DOCUMENT).Text) > 0 Then
                g_clsErrorHandling.DisplayError _
                    "Output Document and Pack Control Id must not both be set."
                txtTask(PACK_CONTROL_ID).SetFocus
                bRet = False
            End If
            If bRet Then
                'check PackControl record exists
                bRet = ValidatePackControlNumber(strPackControlId)
                If bRet = False Then
                    g_clsErrorHandling.DisplayError "A valid Pack Control Id must be entered."
                    txtTask(PACK_CONTROL_ID).SetFocus
                    bRet = False
                End If
            End If
        End If
    End If
    ' RF MAR202 End
          
    If Not m_bIsEdit And bRet Then
        DoesTaskExist
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " ValidateScreenData: " & Err.DESCRIPTION
End Function

' RF MAR202 Start
Private Function ValidatePackControlNumber(ByVal vstrPackControlId As String) As Boolean
On Error GoTo Failed
    Dim objPCTable As PackControlTable
    Dim blnPackExists As Boolean
    Dim colMatchData As Collection
    Dim blnIsValid As Boolean
    
    blnIsValid = True
    If Len(vstrPackControlId) > 0 Then
        ' check pack control record exists
        Set objPCTable = New PackControlTable
        Set colMatchData = New Collection
        colMatchData.Add vstrPackControlId
        blnPackExists = TableAccess(objPCTable).DoesRecordExist(colMatchData)
        If blnPackExists = False Then
            blnIsValid = False
        End If
    End If
    ValidatePackControlNumber = blnIsValid
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " ValidatePackControlNumber: " & Err.DESCRIPTION
End Function
' RF MAR202 End

Private Sub DoesTaskExist()
    On Error GoTo Failed
    Dim sTaskID As String
    Dim bExists As Boolean
    Dim colMatchValues As Collection
           
    Set colMatchValues = New Collection
    
    sTaskID = txtTask(TASK_ID).Text
    colMatchValues.Add sTaskID
    
    bExists = TableAccess(m_clsTaskTable).DoesRecordExist(colMatchValues)

    If bExists Then
        txtTask(TASK_ID).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Task already exists. Please enter a unique Task ID"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user clicks the OK button. If DoOKProcessing succeeds (i.e.,
'                   screen is validated and data saved), set the form return code to true and
'                   hide the form thus returning control back to the caller
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    Dim bRet As Boolean
    
    bRet = DoOKProcessing
    
    If bRet Then
        SetReturnCode
        Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetReturnCode
' Description   :   Sets the return code for the form for any calling method to check. Defaults
'                   to MSG_SUCCESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetReturnCode
' Description   :   Returns the return code for this form, either MSG_SUCCESS or MSG_FAILURE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Saves the fact that the current task has been created/edited for future promotion
'                   to other databases
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection
    
    ' Get list of keys
    Set colValues = New Collection
    colValues.Add Me.txtTask(TASK_ID).Text
    
    TableAccess(m_clsTaskTable).SetKeyMatchValues colValues
    
    ' Save this key
    g_clsHandleUpdates.SaveChangeRequest m_clsTaskTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database. All combo values actually store the
'                   combo ID and not the combo text.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim sTaskID As String
    Dim vVal As Variant
    
    sTaskID = txtTask(TASK_ID).Text
    
    m_sTaskID = sTaskID
    
    ' Task ID
    m_clsTaskTable.SetTaskID sTaskID
    
    ' Task Name
    m_clsTaskTable.SetTaskName txtTask(TASK_NAME).Text
    
    'MAR1545 Update task name in TaskLink table
    m_clsTaskLinkTable.GetTasksForTaskID sTaskID
    
    ' EP511/MAR1669
    If (TableAccess(m_clsTaskLinkTable).RecordCount > 0) Then
        m_clsTaskLinkTable.SetTaskName txtTask(TASK_NAME).Text
    End If
       
    TableAccess(m_clsTaskLinkTable).Update
    
    ' Task Type
    g_clsFormProcessing.HandleComboExtra cboTaskType, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetTaskType vVal
    
    If Not dgAdjustments.DataSource Is Nothing Then
        If dgAdjustments.DataSource.RecordCount > 0 Then
            
'            ' PSC 28/02/2006 MAR1341 - Start
'            If optDays.Value = False Then
'                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_DAYS_FIELD
'            End If
'
'            If optHours.Value = False Then
'                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_HOURS_FIELD
'            End If
'
'            If optMinutes.Value = False Then
'                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_MINUTES_FIELD
'            End If
'            ' PSC 28/02/2006 MAR1341 - End
            
            'OMIGA00003234 GHun
            If cboChasingPeriodTimescale.ListIndex <> CPT_DAYS Then    'Not Days
                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_DAYS_FIELD
            End If
            
            If cboChasingPeriodTimescale.ListIndex <> CPT_HOURS Then    'Not Hours
                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_HOURS_FIELD
            End If
            
            If cboChasingPeriodTimescale.ListIndex <> CPT_MINUTES Then    'Not Minutes
                m_clsTaskPriority.ClearChasingPeriod CHASING_PERIOD_MINUTES_FIELD
            End If
            'OMIGA00003234 End

        End If
    End If
    
'    ' PSC 28/02/2006 MAR1341 - Start
'    If optDays.Value = True Then
'        ' Chasing Period Days
'        m_clsTaskTable.SetChasingPeriodDays txtTask(CHASING_PERIOD).Text
'        m_clsTaskTable.SetChasingPeriodHours Null
'        m_clsTaskTable.SetChasingPeriodMinutes Null
'
'    ElseIf optHours.Value = True Then
'        ' Chasing Period Hours
'        m_clsTaskTable.SetChasingPeriodHours txtTask(CHASING_PERIOD).Text
'        m_clsTaskTable.SetChasingPeriodDays Null
'        m_clsTaskTable.SetChasingPeriodMinutes Null
'    ElseIf optMinutes.Value = True Then
'        ' Chasing Period Minutes
'        m_clsTaskTable.SetChasingPeriodMinutes txtTask(CHASING_PERIOD).Text
'        m_clsTaskTable.SetChasingPeriodDays Null
'        m_clsTaskTable.SetChasingPeriodHours Null
'    Else
'        m_clsTaskTable.SetChasingPeriodDays Null
'        m_clsTaskTable.SetChasingPeriodHours Null
'        m_clsTaskTable.SetChasingPeriodMinutes Null
'    End If
'    ' PSC 28/02/2006 MAR1341 - End
    
    'OMIGA00003234 GHun
    Select Case cboChasingPeriodTimescale.ListIndex
        Case CPT_DAYS
            m_clsTaskTable.SetChasingPeriodDays txtChasingPeriod.Text
            m_clsTaskTable.SetChasingPeriodHours Null
            m_clsTaskTable.SetChasingPeriodMinutes Null
        Case CPT_HOURS
            m_clsTaskTable.SetChasingPeriodHours txtChasingPeriod.Text
            m_clsTaskTable.SetChasingPeriodDays Null
            m_clsTaskTable.SetChasingPeriodMinutes Null
        Case CPT_MINUTES
            m_clsTaskTable.SetChasingPeriodMinutes txtChasingPeriod.Text
            m_clsTaskTable.SetChasingPeriodDays Null
            m_clsTaskTable.SetChasingPeriodHours Null
        Case Else
            m_clsTaskTable.SetChasingPeriodDays Null
            m_clsTaskTable.SetChasingPeriodHours Null
            m_clsTaskTable.SetChasingPeriodMinutes Null
    End Select
    'OMIGA00003234 End
    
    ' Chasing Task
    g_clsFormProcessing.HandleComboExtra cboChasingTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetChasingTask vVal
    
    ' Contact Type
    g_clsFormProcessing.HandleComboExtra cboContactType, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetContactType vVal
    
    ' Input Process
    m_clsTaskTable.SetInputProcess txtTask(INPUT_PROCESS).Text
    
    ' Output Document
    m_clsTaskTable.SetOutputDocument txtTask(OUTPUT_DOCUMENT).Text
        
    ' Pack Control Number - RF MAR202
    m_clsTaskTable.SetPackControlNumber txtTask(PACK_CONTROL_ID).Text

    'Interface
    m_clsTaskTable.SetInterface txtTask(TASK_INTERFACE).Text
    
    ' Owner Type
    g_clsFormProcessing.HandleComboExtra cboOwnerType, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetOwnerType vVal
    
    ' Rule Reference
    m_clsTaskTable.SetRuleRef txtRuleRef.Text
    
    ' Applicant Task
    g_clsFormProcessing.HandleCheckBox chkApplicantTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetApplicantTask vVal
        
    ' AW 31/08/2006 EP1110 - Start
    ' Editable Task
    g_clsFormProcessing.HandleCheckBox chkEditableTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetEditableTaskInd vVal
    
    ' Progress Task
    g_clsFormProcessing.HandleCheckBox chkProgressTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetProgressTaskInd vVal
    ' AW 31/08/2006 EP1110 - End
    
    ' Confirm Print?
    g_clsFormProcessing.HandleCheckBox chkConfirmPrintInd, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetConfirmPrint vVal
        
    ' Automatic on stage entry task?
    g_clsFormProcessing.HandleCheckBox chkAutomaticTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetAutomaticTask vVal
    
    'JD MAR32 ...
    ' Always Automatic?
    g_clsFormProcessing.HandleCheckBox chkAlwaysAuto, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetAlwaysAutoTask vVal
    
    'TAS task?
    g_clsFormProcessing.HandleCheckBox chkTAS, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetTASTask vVal
    
    'TAS queue name
    g_clsFormProcessing.HandleComboExtra cboTASQueueName, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetTASQueueName vVal

    'TAS trigger next stage?
    g_clsFormProcessing.HandleCheckBox chkTriggerNxtStage, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetTASTriggerNxtStage vVal
    ' ...MAR32 end

        
    ' Customer involved?
    g_clsFormProcessing.HandleCheckBox chkCustomerInvolved, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetCustomerInvolved vVal
        
    ' Customer task?
    g_clsFormProcessing.HandleCheckBox chkCustomerTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetCustomerTask vVal
    
' TW 12/12/2006 EP2_453
    ' Remote Owner task?
    g_clsFormProcessing.HandleCheckBox chkRemoteOwnerTask, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetRemoteOwnerTask vVal
' TW 12/12/2006 EP2_453 End
    
    'GD BMIDS00229 22/07/2002
    ' Not applicable status valid?
    g_clsFormProcessing.HandleCheckBox chkNotApplicable, vVal, GET_CONTROL_VALUE
    m_clsTaskTable.SetNotApplicable vVal
    
    m_clsTaskTable.SetDeleteFlag
    
    'OMIGA00003234 GHun
    g_clsFormProcessing.HandleComboIndex cboRuleFrequency, vVal, GET_CONTROL_VALUE
    If vVal = 0 Then
        vVal = vbNullString
    End If

    m_clsTaskTable.SetRuleFrequency vVal
    'OMIGA00003234 End
    
    ' Adjustments DG
    If Not dgAdjustments.DataSource Is Nothing Then
        SaveAdjustments
    End If
    
    TableAccess(m_clsTaskTable).Update
    
    'JD MAR40 update interface task details
    If (frmEditTasks.cboInterfaceType.ListIndex > 0 Or _
        Len(frmEditTasks.cboInterfaceType.List(frmEditTasks.cboInterfaceType.ListIndex)) > 0) And _
        frmEditTasks.cboInterfaceType.List(frmEditTasks.cboInterfaceType.ListIndex) <> "<Select>" Then
        SaveInterfaceTaskDetails
    End If
    
    'MAR1300 GHun SaveLinkedTasks
    If lvLinkedTasks.ListItems.Count > 0 Then
        TableAccess(m_clsLinkedTaskTable).Update
    End If
    'MAR1300 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveInterfaceTaskDetails
' Description   :   Any interface codes appearing in the listbox are saved to database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveInterfaceTaskDetails()
    On Error GoTo Failed
    Dim sTaskID As String
    Dim colValues As Collection
    Dim colLine As Collection
    Dim ListCount As Integer
    Dim vCreateAsComplete As Variant     ' PSC 16/12/2005 MAR606
    
    ' PSC 16/12/2005 MAR606
    sTaskID = txtTask(TASK_ID).Text
   
    'delete existing taskInterfaceMessage and SubMessage entries if any
    On Error Resume Next    'MAR516 GHun Ignore errors when there are no rows to delete
    ' PSC 16/12/2005 MAR606 - Start
    m_clsTaskInterfaceSubMessage.GetAllTaskInterfaceMessageSubType (sTaskID)
    TableAccess(m_clsTaskInterfaceSubMessage).DeleteAllRows
    TableAccess(m_clsTaskInterfaceSubMessage).Update
    m_clsTaskInterfaceMessage.GetTaskInterfaceMessage (sTaskID)
    TableAccess(m_clsTaskInterfaceMessage).DeleteAllRows
    TableAccess(m_clsTaskInterfaceMessage).Update
    ' PSC 16/12/2005 MAR606 - End
    On Error GoTo Failed    'MAR516 GHun
    'MAR516 GHun commented out
    'TableAccess(m_clsTaskInterfaceSubMessage).Update
    'TableAccess(m_clsTaskInterfaceMessage).Update
    'MAR516 End
    
    ' PSC 16/12/2005 MAR606
    g_clsFormProcessing.HandleRadioButtons optInitAsCompYes, optInitAsCompNo, vCreateAsComplete, GET_CONTROL_VALUE
    
    'loop around the listbox saving each row.
    For ListCount = 1 To lstSelectedCodes.ListItems.Count
        Set colLine = lstSelectedCodes.GetLine(ListCount, colValues)
        'There will always be 3 items in colValues but if the third
        'is empty we don't have a subtype
        If colValues.Count = 3 And Len(colValues.Item(3)) = 0 Then
            'we have a message type so save to the message type table only
            g_clsFormProcessing.CreateNewRecord m_clsTaskInterfaceMessage
            m_clsTaskInterfaceMessage.SetTaskID sTaskID
            m_clsTaskInterfaceMessage.SetDeleteFlag True
            m_clsTaskInterfaceMessage.SetInterfaceType colValues.Item(1)
            m_clsTaskInterfaceMessage.SetMessageType colValues.Item(2)
            ' PSC 16/12/2005 MAR606
            m_clsTaskInterfaceMessage.SetCreateAsComplete vCreateAsComplete
            TableAccess(m_clsTaskInterfaceMessage).Update
        Else
            'we have a message sub type so save to the message subtype table
            'and the messageType table if it's not there already
            m_clsTaskInterfaceMessage.GetTaskInterfaceMessage sTaskID, colValues.Item(1), colValues.Item(2)
            If TableAccess(m_clsTaskInterfaceMessage).RecordCount() = 0 Then
                g_clsFormProcessing.CreateNewRecord m_clsTaskInterfaceMessage
                m_clsTaskInterfaceMessage.SetTaskID sTaskID
                m_clsTaskInterfaceMessage.SetDeleteFlag True
                m_clsTaskInterfaceMessage.SetInterfaceType colValues.Item(1)
                m_clsTaskInterfaceMessage.SetMessageType colValues.Item(2)
                ' PSC 16/12/2005 MAR606
                m_clsTaskInterfaceMessage.SetCreateAsComplete vCreateAsComplete
                TableAccess(m_clsTaskInterfaceMessage).Update
            End If
            
            g_clsFormProcessing.CreateNewRecord m_clsTaskInterfaceSubMessage
            m_clsTaskInterfaceSubMessage.SetTaskID sTaskID
            m_clsTaskInterfaceSubMessage.SetDeleteFlag True
            m_clsTaskInterfaceSubMessage.SetInterfaceType colValues.Item(1)
            m_clsTaskInterfaceSubMessage.SetMessageType colValues.Item(2)
            m_clsTaskInterfaceSubMessage.SetMessageSubType colValues.Item(3)
            TableAccess(m_clsTaskInterfaceSubMessage).Update
           
        End If
    Next
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets all screen fields to the values stored on the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    
    Dim vVal As Variant
    
    ' Task ID
    m_sTaskID = m_clsTaskTable.GetTaskID()
    txtTask(TASK_ID).Text = m_sTaskID
    
    ' Task Name
    txtTask(TASK_NAME).Text = m_clsTaskTable.GetTaskName()
    
    ' Task Type
    vVal = m_clsTaskTable.GetTaskType()
    g_clsFormProcessing.HandleComboExtra cboTaskType, vVal, SET_CONTROL_VALUE
    
    'Contact Type
    vVal = m_clsTaskTable.GetContactType()
    g_clsFormProcessing.HandleComboExtra cboContactType, vVal, SET_CONTROL_VALUE

    ' Input Process
    txtTask(INPUT_PROCESS).Text = m_clsTaskTable.GetInputProcess()
    
    'InterFace
    txtTask(TASK_INTERFACE).Text = m_clsTaskTable.GetInterface()
    
    ' Output Document
    txtTask(OUTPUT_DOCUMENT).Text = m_clsTaskTable.GetOutputDocument()
    
    ' Pack Control Number - RF MAR202
    txtTask(PACK_CONTROL_ID).Text = m_clsTaskTable.GetPackControlNumber()
    
    ' Owner Type
    vVal = m_clsTaskTable.GetOwnerType()
    g_clsFormProcessing.HandleComboExtra cboOwnerType, vVal, SET_CONTROL_VALUE
    
    ' Rule Reference
    txtRuleRef.Text = m_clsTaskTable.GetRuleRef()

    ' Applicant Task
    vVal = m_clsTaskTable.GetApplicantTask()
    g_clsFormProcessing.HandleCheckBox chkApplicantTask, vVal, SET_CONTROL_VALUE
        
    ' Not Applicable status valid?
    vVal = m_clsTaskTable.GetNotApplicable()
    g_clsFormProcessing.HandleCheckBox chkNotApplicable, vVal, SET_CONTROL_VALUE

    ' AW 31/08/2006 EP1110 - Start
    ' Editable Task?
    vVal = m_clsTaskTable.GetEditableTaskInd()
    g_clsFormProcessing.HandleCheckBox chkEditableTask, vVal, SET_CONTROL_VALUE
    
    ' Progress Task?
    vVal = m_clsTaskTable.GetProgressTaskInd()
    g_clsFormProcessing.HandleCheckBox chkProgressTask, vVal, SET_CONTROL_VALUE
    ' AW 31/08/2006 EP1110 - End
    
    ' Confirm Print?
    vVal = m_clsTaskTable.GetConfirmPrint()
    g_clsFormProcessing.HandleCheckBox chkConfirmPrintInd, vVal, SET_CONTROL_VALUE
    
    'Automatic Task?
    vVal = m_clsTaskTable.GetAutomaticTask()
    g_clsFormProcessing.HandleCheckBox chkAutomaticTask, vVal, SET_CONTROL_VALUE
    
    'JD MAR32 ...
    ' Always Automatic?
    vVal = m_clsTaskTable.GetAlwaysAutoTask()
    g_clsFormProcessing.HandleCheckBox chkAlwaysAuto, vVal, SET_CONTROL_VALUE
    
    'TAS task?
    vVal = m_clsTaskTable.GetTASTask()
    g_clsFormProcessing.HandleCheckBox chkTAS, vVal, SET_CONTROL_VALUE
    
    Call chkTAS_Click
    
    'TAS queue name
    vVal = m_clsTaskTable.GetTASQueueName()
    g_clsFormProcessing.HandleComboExtra cboTASQueueName, vVal, SET_CONTROL_VALUE

    'TAS trigger next stage?
    vVal = m_clsTaskTable.GetTASTriggerNxtStage()
    g_clsFormProcessing.HandleCheckBox chkTriggerNxtStage, vVal, SET_CONTROL_VALUE
    ' ... MAR32 end
        
    ' Customer contact involved?
    vVal = m_clsTaskTable.GetCustomerInvolved()
    g_clsFormProcessing.HandleCheckBox chkCustomerInvolved, vVal, SET_CONTROL_VALUE

    ' Customer Task?
    vVal = m_clsTaskTable.GetCustomerTask()
    g_clsFormProcessing.HandleCheckBox chkCustomerTask, vVal, SET_CONTROL_VALUE
    
' TW 12/12/2006 EP2_453
    ' Remote Owner task?
    vVal = m_clsTaskTable.GetRemoteOwnerTask()
    g_clsFormProcessing.HandleCheckBox chkRemoteOwnerTask, vVal, SET_CONTROL_VALUE
' TW 12/12/2006 EP2_453 End
    
    'OMIGA00003234 GHun
    vVal = m_clsTaskTable.GetRuleFrequency()
    If IsNull(vVal) Or Len(vVal) = 0 Then
        vVal = 0
    End If
    g_clsFormProcessing.HandleComboIndex cboRuleFrequency, vVal, SET_CONTROL_VALUE
    'OMIGA00003234 End
    
    vVal = m_clsTaskTable.GetChasingPeriodDays
    ' PSC 28/02/2006 MAR1341 - Start
    If Not IsNull(vVal) And Len(vVal) > 0 Then
        txtChasingPeriod.Text = vVal
        'optDays.Value = True
        cboChasingPeriodTimescale.ListIndex = CPT_DAYS 'OMIGA00003234 GHun
    Else
        vVal = m_clsTaskTable.GetChasingPeriodHours
        
        If Not IsNull(vVal) And Len(vVal) > 0 Then
            txtChasingPeriod.Text = vVal
            'optHours.Value = True
            cboChasingPeriodTimescale.ListIndex = CPT_HOURS 'OMIGA00003234 GHun
        Else
            vVal = m_clsTaskTable.GetChasingPeriodMinutes
        
            If Not IsNull(vVal) And Len(vVal) > 0 Then
                txtChasingPeriod.Text = vVal
                'optMinutes.Value = True
                cboChasingPeriodTimescale.ListIndex = CPT_MINUTES 'OMIGA00003234 GHun
            Else
                'optDays.Value = True
                cboChasingPeriodTimescale.ListIndex = CPT_DAYS 'OMIGA00003234 GHun
            End If
        End If
    End If
    ' PSC 28/02/2006 MAR1341 - End
    
    'Populate Adjustments
    PopulateDataGrid
    
    'Are there any Adjustments?
    If dgAdjustments.DataSource.RecordCount > 0 Then
        ' Get No Days
        vVal = m_clsTaskPriority.GetAdjustmentDays()
    End If
    
    'JD MAR40
    PopulateInterfaceTaskDetails
    
    PopulateLinkedTasks 'MAR1300 GHun
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub InitialiseList()
    Dim colHeaders As Collection
    Dim lvcolHeaders As listViewAccess
    
    On Error GoTo Failed
        
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    'Correspondence column header.
    lvcolHeaders.nWidth = 30
    lvcolHeaders.sName = "Interface Type"
    colHeaders.Add lvcolHeaders
    
    lvcolHeaders.nWidth = 30
    lvcolHeaders.sName = "Message Type"
    colHeaders.Add lvcolHeaders
    
    lvcolHeaders.nWidth = 30
    lvcolHeaders.sName = "Message Sub Type"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the correspondence listview.
    lstSelectedCodes.AddHeadings colHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateInterfaceTaskDetails
' Description   :   populate the interfrace task details tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateInterfaceTaskDetails()
    On Error GoTo Failed
    
    Dim colListLine As Collection
    Dim colComboList As Collection
    Dim colValueIDs As Collection
    Dim bIsPopulated As Boolean
    Dim rsTaskInterfaceMessage  As ADODB.Recordset
    Dim rsTaskInterfaceSubMessage  As ADODB.Recordset
    Dim vInterfaceTypeValueID As Variant
    Dim sMessageCode As String
    Dim vMessageTypeValueID As Variant
    Dim vMessageSubTypeValueID As Variant
    Dim vCreateAsComplete As Variant 'PSC 16/11/2005 MAR606
    
    InitialiseList
    
    bIsPopulated = False
    
    'Get Any existing message subtypes
    m_clsTaskInterfaceSubMessage.GetAllTaskInterfaceMessageSubType m_sTaskID
    If TableAccess(m_clsTaskInterfaceSubMessage).RecordCount() > 0 Then
        Set rsTaskInterfaceSubMessage = TableAccess(m_clsTaskInterfaceSubMessage).GetRecordSet()
            
        rsTaskInterfaceSubMessage.MoveFirst
        Do While Not rsTaskInterfaceSubMessage.EOF
            Set colListLine = New Collection
            Set colValueIDs = New Collection
            
            vInterfaceTypeValueID = m_clsTaskInterfaceSubMessage.GetInterfaceType
            sMessageCode = g_clsCombo.GetValueNameFromValueID(CStr(vInterfaceTypeValueID), "InterfaceType")
            colListLine.Add sMessageCode
            colValueIDs.Add vInterfaceTypeValueID

            'get the correct combo for the interface type
            Set colComboList = GetMessageTypeCombo(vInterfaceTypeValueID)
            m_sMessageTypeComboName = colComboList.Item(1)
            m_sMessageSubTypeComboName = colComboList.Item(2)
            
            vMessageTypeValueID = m_clsTaskInterfaceSubMessage.GetMessageType
            sMessageCode = g_clsCombo.GetValueNameFromValueID(CStr(vMessageTypeValueID), m_sMessageTypeComboName)
            colListLine.Add sMessageCode
            colValueIDs.Add vMessageTypeValueID
            
            vMessageSubTypeValueID = m_clsTaskInterfaceSubMessage.GetMessageSubType
            sMessageCode = g_clsCombo.GetValueNameFromValueID(CStr(vMessageSubTypeValueID), m_sMessageSubTypeComboName)
            colListLine.Add sMessageCode
            colValueIDs.Add vMessageSubTypeValueID

            
            lstSelectedCodes.AddLine colListLine, colValueIDs
            bIsPopulated = True
            
            rsTaskInterfaceSubMessage.MoveNext
        Loop
    End If
    
    'Get any existing interface messages
    m_clsTaskInterfaceMessage.GetTaskInterfaceMessage m_sTaskID
    If TableAccess(m_clsTaskInterfaceMessage).RecordCount() > 0 Then
        Set rsTaskInterfaceMessage = TableAccess(m_clsTaskInterfaceMessage).GetRecordSet()
            
        rsTaskInterfaceMessage.MoveFirst
        Do While Not rsTaskInterfaceMessage.EOF
            Set colListLine = New Collection
            Set colValueIDs = New Collection
            vInterfaceTypeValueID = m_clsTaskInterfaceMessage.GetInterfaceType
            sMessageCode = g_clsCombo.GetValueNameFromValueID(CStr(vInterfaceTypeValueID), "InterfaceType")
            colListLine.Add sMessageCode
            colValueIDs.Add vInterfaceTypeValueID

            'get the correct combo for the interface type
            Set colComboList = GetMessageTypeCombo(vInterfaceTypeValueID)
            m_sMessageTypeComboName = colComboList.Item(1)
            m_sMessageSubTypeComboName = colComboList.Item(2)
            
            vMessageTypeValueID = m_clsTaskInterfaceMessage.GetMessageType
            sMessageCode = g_clsCombo.GetValueNameFromValueID(CStr(vMessageTypeValueID), m_sMessageTypeComboName)
            colListLine.Add sMessageCode
            colValueIDs.Add vMessageTypeValueID
            
            'check if there is a subtype for this meesage type. Only add to the listbox
            'if there isn't
            m_clsTaskInterfaceSubMessage.GetTaskInterfaceMessageSubType m_sTaskID, CLng(vInterfaceTypeValueID), CLng(vMessageTypeValueID)
            If TableAccess(m_clsTaskInterfaceSubMessage).RecordCount() = 0 Then
                'make sure there are 3 rows in the collection
                colListLine.Add vbNullString
                colValueIDs.Add vbNullString
                lstSelectedCodes.AddLine colListLine, colValueIDs
                bIsPopulated = True
            End If
            
            'PSC 16/12/2005 MAR606 - Start
            vCreateAsComplete = m_clsTaskInterfaceMessage.GetCreateAsComplete()
            g_clsFormProcessing.HandleRadioButtons optInitAsCompYes, optInitAsCompNo, vCreateAsComplete, SET_CONTROL_VALUE
            'PSC 16/12/2005 MAR606 - End
            
            rsTaskInterfaceMessage.MoveNext
        Loop
    End If
    
    
    If bIsPopulated Then
        cmdDeleteCode.Enabled = True
    Else
        cmdDeleteCode.Enabled = False
    End If
   
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function GetMessageTypeCombo(ByVal vValueID As Variant) As Collection
    On Error GoTo Failed
    Dim colCombos As Collection
    Dim sValidation As String
    Dim sSwapListCombo As String
    Dim sSwapListSubCombo As String
    
    'get the validation type for this combo
    g_clsFormProcessing.HandleComboExtra cboInterfaceType, vValueID, SET_CONTROL_VALUE
    g_clsFormProcessing.GetComboValidation cboInterfaceType, "InterfaceType", sValidation
    
    'find the combo to get the values from
    sSwapListCombo = vbNullString
    sSwapListSubCombo = vbNullString
    Select Case sValidation
    Case "FT"
        sSwapListCombo = "FirstTitleMessageType"
        sSwapListSubCombo = "FirstTitleMessageSubType"
    Case "ES"
        sSwapListCombo = "ESurvMessageType"
    Case "EXP"
        sSwapListCombo = "SMDecisionCode"
        sSwapListSubCombo = "SMReasonCode"
    Case "CAT"
        sSwapListCombo = "CategorisationMessageType"
    Case "CA"
        sSwapListCombo = "CaseAssessementMessageType"
        sSwapListSubCombo = "CaseAssessementMessageSubType"
    Case "XIT2"
        sSwapListCombo = "Xit2MessageType"
        sSwapListSubCombo = "Xit2MessageSubType"
    'EP2_1363 GHun
    Case "FE"
        sSwapListCombo = "FulfilmentEmailMessageType"
    'EP2_1363 End
    Case Else
        'validation type error
    End Select
    
    Set colCombos = New Collection
    colCombos.Add sSwapListCombo
    colCombos.Add sSwapListSubCombo
    
    Set GetMessageTypeCombo = colCombos
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Called on startup to do any Add specific processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    On Error GoTo Failed
    
    g_clsFormProcessing.CreateNewRecord m_clsTaskTable
    
    PopulateDataGrid
    InitialiseList  'MAR516 GHun
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Called on startup to do any Edit specific processing - will need to read the
'                   record required from the database via the task table class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsTaskPriority
    TableAccess(m_clsTaskTable).SetKeyMatchValues m_colKeys
    
    ' Get the data from the database
    TableAccess(m_clsTaskTable).GetTableData
    
    ' Validate we have the record
    If TableAccess(m_clsTaskTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit Tasks - Unable to locate task"
    End If
    
    ' If we get here, we have the data we need
    PopulateScreenFields
    
    cmdAnother.Enabled = False
    txtTask(TASK_ID).Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetKeys
' Description   :   Called from a calling method to set the key values used by this form to locate
'                   a record from the database when this form is initialsed
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(colKeys As Collection)
    On Error GoTo Failed
    Dim sFunctionName As String
    
    sFunctionName = "SetKeys"
    
    If Not colKeys Is Nothing Then
        If colKeys.Count > 0 Then
            Set m_colKeys = colKeys
        Else
            g_clsErrorHandling.RaiseError errKeysEmpty, "EditTasks, " & sFunctionName
        End If
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "EditTasks, " & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Initialize()
    m_bIsEdit = False
End Sub

Private Sub SetGridFields()
    On Error GoTo Failed
    'Dim bRet As Boolean
    Dim fields As FieldData
    Dim colFields As Collection
    Dim colValues As Collection
    Dim colIDS As Collection
    
    'bRet = True
    Set colValues = New Collection
    Set colIDS = New Collection
    Set colFields = New Collection

    ' TaskID
    fields.sField = "TASKID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sTaskID
    fields.sError = vbNullString
    fields.sTitle = "Task ID"
    fields.sOtherField = vbNullString
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = vbNullString
    fields.bDateField = False
    colFields.Add fields

    ' PriorityID
    fields.sField = "CASEPRIORITY"
    fields.bRequired = False
    fields.bVisible = False
    fields.sDefault = vbNullString
    fields.sError = vbNullString
    fields.sTitle = vbNullString
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = vbNullString
    fields.bDateField = False
    colFields.Add fields

    ' Priority
    fields.sField = "CASEPRIORITYTEXT"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = vbNullString
    fields.sError = "A case priority must be entered"
    fields.sTitle = "Case Priority"
    
    g_clsCombo.FindComboGroup "ApplicationPriority", colValues, colIDS, , True
    
    Set fields.colComboValues = colValues
    Set fields.colComboIDS = colIDS
    fields.sOtherField = "CASEPRIORITY"
    fields.bDateField = False
    colFields.Add fields
    
    ' Adjustment Hours
    fields.sField = "ADJUSTMENTHOURS"
    'If this adjustment is in Hours then this field is required!
    'If optHours.Value = True Then
    If cboChasingPeriodTimescale.ListIndex = CPT_HOURS Then 'OMIGA00003234 GHun
        fields.bRequired = True
        fields.sDefault = "0"
    Else
        fields.bRequired = False
        fields.sDefault = vbNullString
    End If
    fields.bVisible = (cboChasingPeriodTimescale.ListIndex = CPT_HOURS) 'OMIGA00003234 GHun
    fields.sError = "Adjustment Hours must be entered"
    fields.sTitle = "Adjustment +/-"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = vbNullString
    fields.bDateField = False
    fields.sMinValue = "-32767"
    colFields.Add fields
        
    ' Adjustment Days
    fields.sField = "ADJUSTMENTDAYS"
    
    ' PSC 28/02/2006 MAR1341
    'If optDays.Value = True Then
    If cboChasingPeriodTimescale.ListIndex = CPT_DAYS Then 'OMIGA00003234 GHun
        fields.bRequired = True
        fields.sDefault = "0"
    Else
        fields.sDefault = vbNullString
        fields.bRequired = False
    End If

    fields.bVisible = (cboChasingPeriodTimescale.ListIndex = CPT_DAYS) 'OMIGA00003234 GHun
    fields.sError = "Adjustment Days must be entered"
    fields.sTitle = "Adjustment +/-"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sMinValue = "-32767"
    fields.sOtherField = vbNullString
    fields.bDateField = False
    colFields.Add fields
    
    ' PSC 28/02/2006 MAR1341 - Start
    ' Adjustment Days
    fields.sField = "ADJUSTMENTMINUTES"
    
    'If optMinutes.Value = True Then
    If cboChasingPeriodTimescale.ListIndex = CPT_MINUTES Then 'OMIGA00003234 GHun
        fields.bRequired = True
        fields.sDefault = "0"
    Else
        fields.sDefault = vbNullString
        fields.bRequired = False
    End If

    fields.bVisible = (cboChasingPeriodTimescale.ListIndex = CPT_MINUTES) 'OMIGA00003234 GHun
    fields.sError = "Adjustment Minutes must be entered"
    fields.sTitle = "Adjustment +/-"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sMinValue = "-32767"
    fields.sOtherField = vbNullString
    fields.bDateField = False
    colFields.Add fields
    ' PSC 28/02/2006 MAR1341 - End

    dgAdjustments.SetColumns colFields, "TaskPriority", "Task Priority"
    dgAdjustments.Enabled = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR1300 GHun
Private Sub lvLinkedTasks_ItemClick(ByVal Item As MSComctlLib.IListItem)
    If Not lvLinkedTasks.SelectedItem Is Nothing Then
        cmdDeleteLinkedTask.Enabled = True
    Else
        cmdDeleteLinkedTask.Enabled = False
    End If
End Sub
'MAR1300 End

'OMIGA00003234 GHun
Private Sub cboChasingPeriodTimescale_Click()
    Dim rsAdjustments As ADODB.Recordset
    
On Error GoTo Failed

    'SYS4509 - If days is selected then clear the value held in hours.
    If Not dgAdjustments.DataSource Is Nothing Then
        'Detach the recordset from the grid prior to updating the field values in code.
        'If we don't, we'll get a cyclic event call chain.
        Set rsAdjustments = dgAdjustments.DataSource
        Set dgAdjustments.DataSource = Nothing
        
        'Clear the days value for each record.
        If Not (rsAdjustments.BOF And rsAdjustments.EOF) Then
            rsAdjustments.MoveFirst
            
            ' PSC 28/02/2006 MAR1341 - Start
            Do While Not rsAdjustments.EOF
                'Copy the value into the new field and then clear the old one.
                
                Select Case cboChasingPeriodTimescale.ListIndex
                    Case CPT_DAYS
                        If Not IsNull(rsAdjustments.fields("ADJUSTMENTHOURS").Value) Then
                            rsAdjustments.fields("ADJUSTMENTDAYS").Value = rsAdjustments.fields("ADJUSTMENTHOURS").Value
                            rsAdjustments.fields("ADJUSTMENTHOURS").Value = Null
                        ElseIf Not IsNull(rsAdjustments.fields("ADJUSTMENTMINUTES").Value) Then
                            rsAdjustments.fields("ADJUSTMENTDAYS").Value = rsAdjustments.fields("ADJUSTMENTMINUTES").Value
                            rsAdjustments.fields("ADJUSTMENTMINUTES").Value = Null
                        End If

                    Case CPT_HOURS
                        If Not IsNull(rsAdjustments.fields("ADJUSTMENTDAYS").Value) Then
                            rsAdjustments.fields("ADJUSTMENTHOURS").Value = rsAdjustments.fields("ADJUSTMENTDAYS").Value
                            rsAdjustments.fields("ADJUSTMENTDAYS").Value = Null
                        ElseIf Not IsNull(rsAdjustments.fields("ADJUSTMENTMINUTES").Value) Then
                            rsAdjustments.fields("ADJUSTMENTHOURS").Value = rsAdjustments.fields("ADJUSTMENTMINUTES").Value
                            rsAdjustments.fields("ADJUSTMENTMINUTES").Value = Null
                        End If
                        
                    Case CPT_MINUTES
                        If Not IsNull(rsAdjustments.fields("ADJUSTMENTDAYS").Value) Then
                            rsAdjustments.fields("ADJUSTMENTMINUTES").Value = rsAdjustments.fields("ADJUSTMENTDAYS").Value
                            rsAdjustments.fields("ADJUSTMENTDAYS").Value = Null
                        ElseIf Not IsNull(rsAdjustments.fields("ADJUSTMENTHOURS").Value) Then
                            rsAdjustments.fields("ADJUSTMENTMINUTES").Value = rsAdjustments.fields("ADJUSTMENTHOURS").Value
                            rsAdjustments.fields("ADJUSTMENTHOURS").Value = Null
                        End If
                End Select
                rsAdjustments.MoveNext
            Loop
            ' PSC 28/02/2006 MAR1341 - End
        End If
        
        'Re-attach the data source.
        Set dgAdjustments.DataSource = rsAdjustments
    End If
    'SYS4509 - End.
        
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'OMIGA00003234 End

Private Function ValidateAdjustments() As Boolean
 
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = True

    ' Validate grid, whichever one it is
    bRet = dgAdjustments.ValidateRows()
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTaskPriority)
        
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Case Priority must be unique"
        End If
    
    End If
    
    ValidateAdjustments = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub PopulateDataGrid()
    
    On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    Set clsTableAccess = m_clsTaskPriority
    
    'Get the Adjustments data
    If m_bIsEdit Then
        clsTableAccess.SetKeyMatchValues m_colKeys
        Set rs = clsTableAccess.GetTableData(POPULATE_KEYS)
    
    Else
        Set rs = clsTableAccess.GetTableData(POPULATE_EMPTY)
    End If
    
    Set dgAdjustments.DataSource = rs

    dgAdjustments.Enabled = True
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveAdjustments
' Description   : Saves the adjustment records associated with the current task record
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveAdjustments()
    
    On Error GoTo Failed
    Dim colUpdateValues As Collection
    Dim clsBandedTable As BandedTable
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsTaskPriority
    
     If Not m_bIsEdit Then
        Set colUpdateValues = New Collection
        
        colUpdateValues.Add m_sTaskID
        Set clsBandedTable = m_clsTaskPriority
        
        clsBandedTable.SetUpdateValues colUpdateValues
        clsBandedTable.SetUpdateSets
        clsBandedTable.DoUpdateSets
        
    End If
    
    TableAccess(m_clsTaskPriority).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR1300 GHun
Private Sub PopulateLinkedTasks()
    Dim colHeaders As Collection
    Dim lvcolHeaders As listViewAccess
    Dim clsTempLinkedTaskTable As LinkedTaskTable
    
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    'Correspondence column header.
    lvcolHeaders.nWidth = 40
    lvcolHeaders.sName = "Task Id"
    colHeaders.Add lvcolHeaders
    
    lvcolHeaders.nWidth = 100
    lvcolHeaders.sName = "Task name"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the correspondence listview.
    lvLinkedTasks.AddHeadings colHeaders
    
    lvLinkedTasks.LoadColumnDetails TypeName(Me)
    
    lvLinkedTasks.HideSelection = False
    
    lvLinkedTasks.AllowColumnReorder = False
    lvLinkedTasks.Sorted = True
    
    'Populate listview
    If Len(txtTask(0).Text) > 0 Then
        m_clsLinkedTaskTable.GetLinkedTasksForTaskId txtTask(0).Text
        g_clsFormProcessing.PopulateFromRecordset lvLinkedTasks, m_clsLinkedTaskTable
    End If
    
    'Populate combo
    Dim colFields As Collection
    Set clsTempLinkedTaskTable = New LinkedTaskTable
    
    clsTempLinkedTaskTable.GetTasksForLinking txtTask(0).Text
    
    Set colFields = clsTempLinkedTaskTable.GetComboFields()
    
    PopulateLinkTaskCombo clsTempLinkedTaskTable, colFields
    
    Set colHeaders = Nothing
    Set clsTempLinkedTaskTable = Nothing
End Sub
'MAR1300 End

'MAR1300 GHun
Private Sub cmdDeleteLinkedTask_Click()
        
    On Error GoTo Failed
    Dim colValues As Collection
    Dim intIndex As Integer
    
    If Not lvLinkedTasks.SelectedItem Is Nothing Then
        GetLinkedTaskKeys colValues
        TableAccess(m_clsLinkedTaskTable).SetKeyMatchValues colValues
        TableAccess(m_clsLinkedTaskTable).DeleteRecords
        
        'Add the deleted item back into the combo
        cboLinkTask.AddItem lvLinkedTasks.SelectedItem.Text & ": " & lvLinkedTasks.SelectedItem.SubItems(1)
        
        intIndex = lvLinkedTasks.SelectedItem.Index
        lvLinkedTasks.RemoveLine lvLinkedTasks.SelectedItem.Index
        
        If intIndex > lvLinkedTasks.ListItems.Count Then
            intIndex = intIndex - 1
        End If
        If (intIndex > 0 And lvLinkedTasks.ListItems.Count > 0) Then
            lvLinkedTasks.ListItems.Item(intIndex).Selected = True
        End If
    End If
    
    Set colValues = Nothing
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'MAR1300 End

'MAR1300 GHun
Private Sub GetLinkedTaskKeys(colKeyValues As Collection)
    
    On Error GoTo Failed
    
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = lvLinkedTasks.SelectedItem.Index
    Set clsPopulateDetails = lvLinkedTasks.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colKeyValues = clsPopulateDetails.GetKeyMatchValues
    End If
    
    If colKeyValues Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "LinkedTasks, unable to obtain Keys"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'MAR1300 End

'MAR1300 GHun
Private Sub PopulateLinkTaskCombo(clsTableAccess As TableAccess, colFields As Collection)
    Dim colTaskNames As Collection
    Dim colTaskIds As Collection
    Dim intCount As Integer
    Dim varTaskName As Variant
    Dim strTaskName As String
    Dim strTaskDesc As String
    
    'Set m_colListTasks = New Collection
    Set colTaskNames = New Collection
    Set colTaskIds = New Collection
      
    g_clsFormProcessing.PopulateCollectionFromTable clsTableAccess, colFields, colTaskNames, colTaskIds

    cboLinkTask.Clear
    
    intCount = 0
    For Each varTaskName In colTaskNames
        strTaskName = CStr(varTaskName)
        intCount = intCount + 1
        
        strTaskDesc = colTaskIds(intCount) & ": " & strTaskName
        
        cboLinkTask.AddItem strTaskDesc
    Next
    cboLinkTask.AddItem COMBO_NONE, 0
    cboLinkTask.ListIndex = 0
End Sub
'MAR1300 End

'OMIGA00003234 GHun
Private Sub SSTab1_Click(PreviousTab As Integer)
    SetTabstops Me
End Sub
'OMIGA00003234 End
