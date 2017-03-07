VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmLenderDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lender Details"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmLenderDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7860
      TabIndex        =   1
      Top             =   8040
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   8040
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTabLenderDetails 
      Height          =   7900
      Left            =   72
      TabIndex        =   2
      Top             =   60
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   13917
      _Version        =   393216
      Tabs            =   7
      TabHeight       =   617
      TabCaption(0)   =   "Lender Details"
      TabPicture(0)   =   "frmLenderDetails.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeLenderDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmeMIG"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fmeAdditional1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Contact Details"
      TabPicture(1)   =   "frmLenderDetails.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calculation Parameters"
      TabPicture(2)   =   "frmLenderDetails.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Legal Fees"
      TabPicture(3)   =   "frmLenderDetails.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Other Fees"
      TabPicture(4)   =   "frmLenderDetails.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "HLC Rate Sets"
      TabPicture(5)   =   "frmLenderDetails.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame10"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "BACS/CHAPS Parameters"
      TabPicture(6)   =   "frmLenderDetails.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraFrame(2)"
      Tab(6).Control(1)=   "fraFrame(1)"
      Tab(6).Control(2)=   "fraFrame(0)"
      Tab(6).ControlCount=   3
      Begin VB.Frame Frame13 
         Caption         =   "Add to Loan"
         Height          =   1416
         Left            =   75
         TabIndex        =   200
         Top             =   6330
         Width           =   8685
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Credit Limit Increase Fee"
            Height          =   255
            Index           =   12
            Left            =   6495
            TabIndex        =   171
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Additional Borrowing Fee"
            Height          =   255
            Index           =   11
            Left            =   4410
            TabIndex        =   170
            Top             =   1080
            Width           =   2085
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Re Valuation Fee"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   162
            Top             =   1080
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Valuation Fee"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   161
            Top             =   790
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Administration Fee"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   160
            Top             =   500
            Width           =   1680
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Porting Fee"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   159
            Top             =   210
            Width           =   1575
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Reinspection Fee"
            Height          =   255
            Index           =   3
            Left            =   2325
            TabIndex        =   163
            Top             =   210
            Width           =   1575
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Arrangement Fee"
            Height          =   255
            Index           =   4
            Left            =   2325
            TabIndex        =   164
            Top             =   500
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "HLC Fee"
            Height          =   255
            Index           =   5
            Left            =   2325
            TabIndex        =   165
            Top             =   790
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "TT Fee"
            Height          =   255
            Index           =   6
            Left            =   4410
            TabIndex        =   167
            Top             =   210
            Width           =   1575
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Sealing Fee"
            Height          =   255
            Index           =   7
            Left            =   4410
            TabIndex        =   168
            Top             =   500
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Deeds Release Fee"
            Height          =   255
            Index           =   8
            Left            =   4410
            TabIndex        =   169
            Top             =   790
            Width           =   1755
         End
         Begin VB.CheckBox chkLenderDetails 
            Caption         =   "Product Switch Fee"
            Height          =   255
            Index           =   10
            Left            =   2325
            TabIndex        =   166
            Top             =   1080
            Width           =   1755
         End
      End
      Begin VB.Frame Frame5 
         Height          =   480
         Left            =   60
         TabIndex        =   199
         Top             =   5850
         Width           =   8700
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   2790
            TabIndex        =   203
            Top             =   120
            Width           =   2040
            Begin VB.OptionButton optAddToLoan 
               Caption         =   "Last"
               Height          =   255
               Index           =   1
               Left            =   705
               TabIndex        =   158
               Top             =   45
               Width           =   735
            End
            Begin VB.OptionButton optAddToLoan 
               Caption         =   "First"
               Height          =   255
               Index           =   0
               Left            =   30
               TabIndex        =   157
               Top             =   45
               Width           =   675
            End
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Add Costs to Loan Component"
            Height          =   192
            Index           =   29
            Left            =   60
            TabIndex        =   204
            Top             =   189
            Width           =   2172
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1270
         Left            =   60
         TabIndex        =   192
         Top             =   4580
         Width           =   8685
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   16
            Left            =   2760
            TabIndex        =   151
            Top             =   168
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   17
            Left            =   7460
            TabIndex        =   152
            Top             =   168
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   18
            Left            =   2760
            TabIndex        =   153
            Top             =   516
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   19
            Left            =   7460
            TabIndex        =   154
            Top             =   516
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   21
            Left            =   7460
            TabIndex        =   156
            Top             =   876
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   20
            Left            =   2760
            TabIndex        =   155
            Top             =   876
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Highest Earner Income Multiple"
            Height          =   195
            Index           =   20
            Left            =   45
            TabIndex        =   198
            Top             =   923
            Width           =   2205
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Joint Income Multiple"
            Height          =   192
            Index           =   21
            Left            =   4656
            TabIndex        =   197
            Top             =   924
            Width           =   1488
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Second Earner Income Multiple"
            Height          =   192
            Index           =   19
            Left            =   4656
            TabIndex        =   196
            Top             =   564
            Width           =   2220
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Single Income Multiple"
            Height          =   195
            Index           =   18
            Left            =   45
            TabIndex        =   195
            Top             =   563
            Width           =   1590
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Guaranteed Income Multiple"
            Height          =   192
            Index           =   17
            Left            =   4656
            TabIndex        =   194
            Top             =   216
            Width           =   1992
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Non Guaranteed Income Multiple"
            Height          =   195
            Index           =   16
            Left            =   45
            TabIndex        =   193
            Top             =   215
            Width           =   2340
         End
      End
      Begin VB.Frame fmeAdditional1 
         Height          =   940
         Left            =   60
         TabIndex        =   185
         Top             =   3650
         Width           =   8685
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   10
            Left            =   1920
            TabIndex        =   145
            Top             =   168
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   13
            Left            =   2640
            TabIndex        =   147
            Top             =   516
            Width           =   996
            _ExtentX        =   1746
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   9
            Left            =   6600
            TabIndex        =   146
            Top             =   168
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   11
            Left            =   4200
            TabIndex        =   148
            Top             =   516
            Width           =   696
            _ExtentX        =   1217
            _ExtentY        =   503
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MaxValue        =   "100"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            MaxLength       =   6
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   12
            Left            =   7580
            TabIndex        =   150
            Top             =   516
            Width           =   996
            _ExtentX        =   1746
            _ExtentY        =   503
            TextType        =   6
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MaxValue        =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            MaxLength       =   4
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   15
            Left            =   5793
            TabIndex        =   149
            Top             =   516
            Width           =   1000
            _ExtentX        =   1746
            _ExtentY        =   503
            TextType        =   6
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MaxValue        =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            MaxLength       =   4
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   192
            Index           =   25
            Left            =   1920
            TabIndex        =   205
            Top             =   564
            Width           =   540
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Deeds Release Fee"
            Height          =   195
            Index           =   12
            Left            =   60
            TabIndex        =   191
            Top             =   210
            Width           =   1410
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Arrangement Fee"
            Height          =   195
            Index           =   15
            Left            =   60
            TabIndex        =   190
            Top             =   563
            Width           =   1335
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Sealing Fee"
            Height          =   192
            Index           =   11
            Left            =   4680
            TabIndex        =   189
            Top             =   216
            Width           =   876
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   192
            Index           =   13
            Left            =   3930
            TabIndex        =   188
            Top             =   564
            Width           =   144
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Max"
            Height          =   192
            Index           =   14
            Left            =   7080
            TabIndex        =   187
            Top             =   564
            Width           =   312
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Min"
            Height          =   192
            Index           =   24
            Left            =   5400
            TabIndex        =   186
            Top             =   564
            Width           =   312
         End
      End
      Begin VB.Frame fmeMIG 
         Height          =   910
         Left            =   60
         TabIndex        =   180
         Top             =   2730
         Width           =   8685
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   5
            Left            =   6600
            TabIndex        =   142
            Top             =   168
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   7
            Left            =   1920
            TabIndex        =   143
            Top             =   510
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   6
            Left            =   6600
            TabIndex        =   144
            Top             =   516
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   14
            Left            =   1920
            TabIndex        =   141
            Top             =   165
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            TextType        =   7
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
            MaxLength       =   4
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   315
            Index           =   8
            Left            =   1815
            TabIndex        =   201
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            TextType        =   2
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
            MaxLength       =   12
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "TT Fee"
            Height          =   195
            Index           =   10
            Left            =   195
            TabIndex        =   202
            Top             =   1020
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Min. HLC Premium"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   184
            Top             =   557
            Width           =   1335
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "HLC Start LTV%"
            Height          =   195
            Index           =   6
            Left            =   4650
            TabIndex        =   183
            Top             =   212
            Width           =   1335
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "Ignore HLC Premium"
            Height          =   195
            Index           =   7
            Left            =   4665
            TabIndex        =   182
            Top             =   557
            Width           =   1695
         End
         Begin VB.Label lblLenderDetails 
            Caption         =   "HLC Threshold"
            Height          =   195
            Index           =   22
            Left            =   60
            TabIndex        =   181
            Top             =   212
            Width           =   1335
         End
      End
      Begin VB.Frame fmeLenderDetails 
         Height          =   1650
         Left            =   60
         TabIndex        =   130
         Top             =   1080
         Width           =   8685
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   6465
            TabIndex        =   131
            Top             =   1224
            Width           =   1365
            Begin VB.OptionButton optAddIPT 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   660
               TabIndex        =   140
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optAddIPT 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   139
               Top             =   0
               Width           =   675
            End
         End
         Begin MSGOCX.MSGComboBox cboOrganisationType 
            Height          =   315
            Left            =   6480
            TabIndex        =   136
            Top             =   155
            Width           =   2130
            _ExtentX        =   3757
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
         Begin MSGOCX.MSGComboBox cboAddressType 
            Height          =   315
            Left            =   6480
            TabIndex        =   137
            Top             =   503
            Width           =   2130
            _ExtentX        =   3757
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
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   3
            Left            =   6480
            TabIndex        =   138
            Top             =   875
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   0
            Left            =   1920
            TabIndex        =   132
            Top             =   168
            Width           =   1140
            _ExtentX        =   2011
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
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   1
            Left            =   1920
            TabIndex        =   133
            Top             =   516
            Width           =   2292
            _ExtentX        =   4048
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
            MaxLength       =   50
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   4
            Left            =   1920
            TabIndex        =   135
            Top             =   1230
            Width           =   1140
            _ExtentX        =   2011
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
            MaxLength       =   2
         End
         Begin MSGOCX.MSGEditBox txtLenderDetails 
            Height          =   288
            Index           =   2
            Left            =   1920
            TabIndex        =   134
            Top             =   875
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
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
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Organisation Type"
            Height          =   195
            Index           =   1
            Left            =   4635
            TabIndex        =   179
            Top             =   215
            Width           =   1290
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Lender Code"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   178
            Top             =   215
            Width           =   915
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Lender Name"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   177
            Top             =   563
            Width           =   960
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Name && Address Type"
            Height          =   195
            Index           =   23
            Left            =   4650
            TabIndex        =   176
            Top             =   563
            Width           =   1575
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Max. Loan Components"
            Height          =   192
            Index           =   5
            Left            =   60
            TabIndex        =   175
            Top             =   1278
            Width           =   1680
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   174
            Top             =   922
            Width           =   720
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "End Date"
            Height          =   195
            Index           =   4
            Left            =   4650
            TabIndex        =   173
            Top             =   922
            Width           =   675
         End
         Begin VB.Label lblLenderDetails 
            AutoSize        =   -1  'True
            Caption         =   "Add IPT to Premium"
            Height          =   195
            Index           =   9
            Left            =   4650
            TabIndex        =   172
            Top             =   1277
            Width           =   1410
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Hexagon CHAPS Parameters"
         Height          =   975
         Index           =   2
         Left            =   -74640
         TabIndex        =   82
         Top             =   4800
         Width           =   8295
         Begin MSGOCX.MSGEditBox txtHexagonCHAPSAccountNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   83
            Top             =   360
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
         Begin VB.Label lblLabel 
            Caption         =   "Account Number"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   84
            Top             =   390
            Width           =   1335
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "Barclays CHAPS Parameters"
         Height          =   1695
         Index           =   1
         Left            =   -74640
         TabIndex        =   75
         Top             =   3000
         Width           =   8295
         Begin MSGOCX.MSGEditBox txtBarclaysCHAPSSortCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   76
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Mask            =   "99-99-99"
            TextType        =   4
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
         Begin MSGOCX.MSGEditBox txtBarclaysCHAPSAccountNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   77
            Top             =   720
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
         Begin MSGOCX.MSGEditBox txtBarclaysCHAPSOrderingInstitution 
            Height          =   315
            Left            =   1680
            TabIndex        =   78
            Top             =   1080
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
            MaxLength       =   35
         End
         Begin VB.Label lblLabel 
            Caption         =   "Ordering Institution"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   81
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label lblLabel 
            Caption         =   "Account Number"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   80
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblLabel 
            Caption         =   "Sort Code"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraFrame 
         Caption         =   "BACS Parameters"
         Height          =   1335
         Index           =   0
         Left            =   -74640
         TabIndex        =   70
         Top             =   1560
         Width           =   8295
         Begin MSGOCX.MSGEditBox txtBACSAccountNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   72
            Top             =   720
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
         Begin MSGOCX.MSGEditBox txtBACSSortCode 
            Height          =   315
            Left            =   1680
            TabIndex        =   71
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            Mask            =   "99-99-99"
            TextType        =   4
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
         Begin VB.Label lblLabel 
            Caption         =   "Account Number"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   74
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblLabel 
            Caption         =   "Sort Code"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   6195
         Left            =   -74700
         TabIndex        =   47
         Tag             =   "Tab6"
         Top             =   1215
         Width           =   8355
         Begin VB.CommandButton cmdMIGAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   540
            TabIndex        =   22
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton cmdMIGEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1860
            TabIndex        =   23
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton cmdMIGDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3180
            TabIndex        =   24
            Top             =   4200
            Width           =   1155
         End
         Begin MSGOCX.MSGListView lvMIGRateSets 
            Height          =   3015
            Left            =   420
            TabIndex        =   21
            Top             =   720
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   5318
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame9"
         Height          =   6255
         Left            =   -74700
         TabIndex        =   46
         Tag             =   "Tab5"
         Top             =   1140
         Width           =   8415
         Begin VB.Frame frameFeeSetBands 
            Caption         =   "Other Fees"
            Height          =   4332
            Left            =   0
            TabIndex        =   48
            Top             =   240
            Width           =   8235
            Begin MSGOCX.MSGDataGrid dgOtherFees 
               Height          =   3855
               Left            =   420
               TabIndex        =   20
               Top             =   240
               Width           =   7635
               _ExtentX        =   13467
               _ExtentY        =   6800
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
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   6015
         Left            =   -74640
         TabIndex        =   45
         Tag             =   "Tab4"
         Top             =   1200
         Width           =   8175
         Begin VB.CommandButton cmdDeleteLegalFees 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3000
            TabIndex        =   19
            Top             =   3780
            Width           =   1155
         End
         Begin VB.CommandButton cmdEditLegalFees 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1680
            TabIndex        =   18
            Top             =   3780
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddLegalFees 
            Caption         =   "&Add"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   3780
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvLegalFees 
            Height          =   3015
            Left            =   420
            TabIndex        =   16
            Top             =   480
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   5318
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   6555
         Left            =   -74760
         TabIndex        =   37
         Tag             =   "Tab3"
         Top             =   1100
         Width           =   8535
         Begin VB.Frame Frame11 
            Caption         =   "Rounding"
            Height          =   2055
            Left            =   0
            TabIndex        =   85
            Top             =   4370
            Width           =   8295
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   6
               Left            =   1920
               TabIndex        =   119
               Top             =   480
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               TextType        =   2
               FontSize        =   8.25
               FontName        =   "MS Sans Serif"
               FirstCharUpper  =   0   'False
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
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   7
               Left            =   1920
               TabIndex        =   122
               Top             =   840
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               TextType        =   2
               FontSize        =   8.25
               FontName        =   "MS Sans Serif"
               FirstCharUpper  =   0   'False
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
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   8
               Left            =   1920
               TabIndex        =   125
               Top             =   1200
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               TextType        =   2
               FontSize        =   8.25
               FontName        =   "MS Sans Serif"
               FirstCharUpper  =   0   'False
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
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   9
               Left            =   1920
               TabIndex        =   128
               Top             =   1560
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               TextType        =   2
               FontSize        =   8.25
               FontName        =   "MS Sans Serif"
               FirstCharUpper  =   0   'False
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
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   11
               Left            =   3480
               TabIndex        =   120
               Top             =   480
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
               Mandatory       =   -1  'True
               Text            =   ""
            End
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   12
               Left            =   3480
               TabIndex        =   123
               Top             =   840
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
               Mandatory       =   -1  'True
               Text            =   ""
            End
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   13
               Left            =   3480
               TabIndex        =   126
               Top             =   1200
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
               Mandatory       =   -1  'True
               Text            =   ""
            End
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   14
               Left            =   3480
               TabIndex        =   129
               Top             =   1560
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
               Mandatory       =   -1  'True
               Text            =   ""
            End
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   15
               Left            =   5400
               TabIndex        =   121
               Top             =   480
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
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   16
               Left            =   5400
               TabIndex        =   124
               Top             =   840
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
            Begin MSGOCX.MSGComboBox cboAdditionalParams 
               Height          =   315
               Index           =   17
               Left            =   5400
               TabIndex        =   127
               Top             =   1200
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
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Charge"
               Height          =   195
               Index           =   20
               Left            =   240
               TabIndex        =   92
               Top             =   1560
               Width           =   510
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Balance"
               Height          =   195
               Index           =   19
               Left            =   240
               TabIndex        =   91
               Top             =   1200
               Width           =   585
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Payment"
               Height          =   195
               Index           =   18
               Left            =   240
               TabIndex        =   90
               Top             =   840
               Width           =   615
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Accrued Interest"
               Height          =   195
               Index           =   17
               Left            =   240
               TabIndex        =   89
               Top             =   480
               Width           =   1170
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "When Rounding Occurs"
               Height          =   195
               Index           =   23
               Left            =   5400
               TabIndex        =   88
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Rounding Direction"
               Height          =   195
               Index           =   22
               Left            =   3480
               TabIndex        =   87
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Rounding Factor"
               Height          =   195
               Index           =   21
               Left            =   1920
               TabIndex        =   86
               ToolTipText     =   "Rounding factor (e.g. 0.01 represents a penny)"
               Top             =   240
               Width           =   1185
            End
         End
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   2
            Left            =   7320
            TabIndex        =   115
            Top             =   930
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            TextType        =   6
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            MinValue        =   "0"
            MaxValue        =   "11"
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   1
            Left            =   7320
            TabIndex        =   114
            Top             =   555
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            TextType        =   5
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   0
            Left            =   2160
            TabIndex        =   102
            Top             =   180
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   3
            Left            =   2160
            TabIndex        =   105
            Top             =   1320
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   4
            Left            =   2160
            TabIndex        =   106
            Top             =   1695
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   5
            Left            =   2160
            TabIndex        =   107
            Top             =   2076
            Width           =   2532
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   6
            Left            =   2160
            TabIndex        =   108
            Top             =   2445
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   10
            Left            =   2160
            TabIndex        =   112
            Top             =   3960
            Width           =   3195
            _ExtentX        =   5636
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   103
            Top             =   570
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   3
            Left            =   7320
            TabIndex        =   116
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   4
            Left            =   7320
            TabIndex        =   117
            Top             =   1695
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            TextType        =   6
            FontSize        =   8.25
            FontName        =   "MS Sans Serif"
            FirstCharUpper  =   0   'False
            MinValue        =   "1"
            MaxValue        =   "28"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
            MaxLength       =   2
         End
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   5
            Left            =   7320
            TabIndex        =   118
            Top             =   2070
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            TextType        =   2
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   315
            Index           =   0
            Left            =   7320
            TabIndex        =   113
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            TextType        =   5
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
            MaxLength       =   5
         End
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   104
            Top             =   930
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   7
            Left            =   2160
            TabIndex        =   109
            Top             =   2835
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   8
            Left            =   2160
            TabIndex        =   110
            Top             =   3210
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
         Begin MSGOCX.MSGComboBox cboAdditionalParams 
            Height          =   315
            Index           =   9
            Left            =   2160
            TabIndex        =   111
            Top             =   3585
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
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Output Balance Schedule"
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   101
            ToolTipText     =   "Indicates the interval at which the current account balance for the mortgage should be output"
            Top             =   3642
            Width           =   1830
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Output Payment Schedule"
            Height          =   195
            Index           =   8
            Left            =   0
            TabIndex        =   100
            ToolTipText     =   "Indicates the interval at which the monthly mortgage payments should be output"
            Top             =   3264
            Width           =   1860
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "APR Month"
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   99
            ToolTipText     =   $"frmLenderDetails.frx":0506
            Top             =   2886
            Width           =   825
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Interest Charge Month"
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   98
            ToolTipText     =   $"frmLenderDetails.frx":05AA
            Top             =   2508
            Width           =   1575
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Days Per Year"
            Height          =   195
            Index           =   16
            Left            =   5160
            TabIndex        =   97
            ToolTipText     =   "The number of days assumed in a calendar year"
            Top             =   2130
            Width           =   1020
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Specified Payment Day"
            Height          =   195
            Index           =   15
            Left            =   5160
            TabIndex        =   96
            ToolTipText     =   $"frmLenderDetails.frx":063F
            Top             =   1752
            Width           =   1650
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Cooling Off Period"
            Height          =   195
            Index           =   14
            Left            =   5160
            TabIndex        =   95
            ToolTipText     =   "The number of days in any ""cooling off"" period. Not required if Payment Timing is set to 0"
            Top             =   1374
            Width           =   1275
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Accrued Days Added"
            Height          =   195
            Index           =   11
            Left            =   5160
            TabIndex        =   94
            ToolTipText     =   "Number of case-specific additional days' interest charged."
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Accrued Days Included"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   93
            ToolTipText     =   "Indicator showing how the days left in the month/year are counted"
            Top             =   996
            Width           =   1665
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Accrued Interest Payable"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   69
            ToolTipText     =   "Indicator showing when accrued interest is payable"
            Top             =   618
            Width           =   1785
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Completion Days Adjustment"
            Height          =   195
            Index           =   12
            Left            =   5160
            TabIndex        =   44
            ToolTipText     =   "Final adjustment in days to determine the default Completion Date"
            Top             =   615
            Width           =   2010
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Completion Timing"
            Height          =   195
            Index           =   10
            Left            =   0
            TabIndex        =   43
            ToolTipText     =   $"frmLenderDetails.frx":06F0
            Top             =   4020
            Width           =   1290
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Months To Completion"
            Height          =   195
            Index           =   13
            Left            =   5160
            TabIndex        =   42
            ToolTipText     =   "Number of calendar months from the Application Date to the default Completion Date"
            Top             =   996
            Width           =   1590
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Interest Charged Indicator"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   41
            ToolTipText     =   "Indicator showing the balance on which interest is charged"
            Top             =   2130
            Width           =   1830
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Accounting Start Month"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   40
            ToolTipText     =   "The calendar month when the lender's accounting year starts"
            Top             =   1752
            Width           =   1680
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Payment Timing"
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   39
            ToolTipText     =   "Indicator showing when the first full monthly payment is due"
            Top             =   1374
            Width           =   1125
         End
         Begin VB.Label lblAdditionalParams 
            AutoSize        =   -1  'True
            Caption         =   "Accrued Interest Indicator"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   38
            ToolTipText     =   "Indicator showing how accrued interest is calculated"
            Top             =   240
            Width           =   1830
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6372
         Left            =   -74925
         TabIndex        =   25
         Tag             =   "Tab2"
         Top             =   1250
         Width           =   8595
         Begin VB.Frame Frame6 
            Caption         =   "Address"
            Height          =   2895
            Left            =   360
            TabIndex        =   28
            Top             =   0
            Width           =   7935
            Begin MSGOCX.MSGComboBox cboCountry 
               Height          =   315
               Left            =   1560
               TabIndex        =   10
               Top             =   2415
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
               Height          =   288
               Index           =   0
               Left            =   1560
               TabIndex        =   3
               Top             =   250
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
               MaxLength       =   8
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   1
               Left            =   1560
               TabIndex        =   4
               Top             =   610
               Width           =   3795
               _ExtentX        =   6694
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   3
               Left            =   1560
               TabIndex        =   6
               Top             =   970
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   4
               Left            =   1560
               TabIndex        =   7
               Top             =   1330
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   5
               Left            =   1560
               TabIndex        =   8
               Top             =   1690
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   6
               Left            =   1560
               TabIndex        =   9
               Top             =   2050
               Width           =   6135
               _ExtentX        =   10821
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   2
               Left            =   6420
               TabIndex        =   5
               Top             =   610
               Width           =   1275
               _ExtentX        =   2249
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
               MaxLength       =   10
            End
            Begin VB.Label lblDetails 
               Caption         =   "Country"
               Height          =   195
               Index           =   9
               Left            =   300
               TabIndex        =   36
               Top             =   2457
               Width           =   975
            End
            Begin VB.Label lblCounty 
               Caption         =   "County"
               Height          =   195
               Index           =   8
               Left            =   300
               TabIndex        =   35
               Top             =   2097
               Width           =   1095
            End
            Begin VB.Label lblDetails 
               Caption         =   "Town"
               Height          =   195
               Index           =   7
               Left            =   300
               TabIndex        =   34
               Top             =   1737
               Width           =   975
            End
            Begin VB.Label lblDetails 
               Caption         =   "District"
               Height          =   195
               Index           =   6
               Left            =   300
               TabIndex        =   33
               Top             =   1377
               Width           =   1095
            End
            Begin VB.Label lblDetails 
               Caption         =   "Street"
               Height          =   195
               Index           =   5
               Left            =   300
               TabIndex        =   32
               Top             =   1017
               Width           =   915
            End
            Begin VB.Label lblDetails 
               Caption         =   "No."
               Height          =   195
               Index           =   4
               Left            =   5940
               TabIndex        =   31
               Top             =   657
               Width           =   495
            End
            Begin VB.Label lblDetails 
               Caption         =   "Building Name"
               Height          =   195
               Index           =   3
               Left            =   300
               TabIndex        =   30
               Top             =   657
               Width           =   1155
            End
            Begin VB.Label lblDetails 
               Caption         =   "Postcode"
               Height          =   195
               Index           =   2
               Left            =   300
               TabIndex        =   29
               Top             =   297
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Contact Details"
            Height          =   2970
            Left            =   360
            TabIndex        =   26
            Top             =   3060
            Width           =   7935
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   315
               Index           =   18
               Left            =   5580
               TabIndex        =   13
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
               TabIndex        =   11
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
               TabIndex        =   12
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
               Height          =   288
               Index           =   7
               Left            =   1560
               TabIndex        =   14
               Top             =   960
               Width           =   2715
               _ExtentX        =   4789
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   8
               Left            =   5580
               TabIndex        =   15
               Top             =   960
               Width           =   2055
               _ExtentX        =   3625
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
               MaxLength       =   40
            End
            Begin MSGOCX.MSGComboBox cboType2 
               Height          =   315
               Left            =   240
               TabIndex        =   58
               Top             =   2090
               Width           =   1332
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
               TabIndex        =   53
               Top             =   1710
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
               Height          =   288
               Index           =   9
               Left            =   1680
               TabIndex        =   54
               Top             =   1700
               Width           =   1035
               _ExtentX        =   1826
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
               MaxLength       =   3
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   10
               Left            =   2880
               TabIndex        =   55
               Top             =   1700
               Width           =   1155
               _ExtentX        =   2037
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
               MaxLength       =   5
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   11
               Left            =   4200
               TabIndex        =   56
               Top             =   1700
               Width           =   1875
               _ExtentX        =   3307
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
               MaxLength       =   30
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   12
               Left            =   6240
               TabIndex        =   57
               Top             =   1700
               Width           =   1395
               _ExtentX        =   2461
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   13
               Left            =   1680
               TabIndex        =   59
               Top             =   2090
               Width           =   1032
               _ExtentX        =   1826
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
               MaxLength       =   3
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   14
               Left            =   2880
               TabIndex        =   60
               Top             =   2090
               Width           =   1155
               _ExtentX        =   2037
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
               MaxLength       =   5
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   15
               Left            =   4200
               TabIndex        =   61
               Top             =   2090
               Width           =   1875
               _ExtentX        =   3307
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
               MaxLength       =   30
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   16
               Left            =   6240
               TabIndex        =   62
               Top             =   2090
               Width           =   1395
               _ExtentX        =   2461
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
               MaxLength       =   10
            End
            Begin MSGOCX.MSGEditBox txtContactDetails 
               Height          =   288
               Index           =   17
               Left            =   1680
               TabIndex        =   68
               Top             =   2520
               Width           =   5952
               _ExtentX        =   10504
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
               MaxLength       =   30
            End
            Begin VB.Label lblOtherTitle 
               Caption         =   "Other Title"
               Height          =   255
               Left            =   4500
               TabIndex        =   206
               Top             =   660
               Width           =   800
            End
            Begin VB.Label lblContact 
               Caption         =   "Ext. No."
               Height          =   195
               Index           =   11
               Left            =   6240
               TabIndex        =   67
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblContact 
               Caption         =   "Country Code"
               Height          =   255
               Index           =   10
               Left            =   1680
               TabIndex        =   66
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "Area Code"
               Height          =   255
               Left            =   2880
               TabIndex        =   65
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Telephone Number"
               Height          =   255
               Left            =   4200
               TabIndex        =   64
               Top             =   1440
               Width           =   1815
            End
            Begin VB.Label Label21 
               Caption         =   "Type"
               Height          =   255
               Left            =   240
               TabIndex        =   63
               Top             =   1440
               Width           =   855
            End
            Begin VB.Label lblContact 
               Caption         =   "Forename"
               Height          =   195
               Index           =   1
               Left            =   4500
               TabIndex        =   52
               Top             =   1020
               Width           =   1035
            End
            Begin VB.Label lblContact 
               Caption         =   "Surname"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   51
               Top             =   1020
               Width           =   1155
            End
            Begin VB.Label lblContact 
               Caption         =   "Contact Title"
               Height          =   252
               Index           =   18
               Left            =   240
               TabIndex        =   50
               Top             =   660
               Width           =   1272
            End
            Begin VB.Label lblContact 
               Caption         =   "Contact Type"
               Height          =   252
               Index           =   5
               Left            =   240
               TabIndex        =   49
               Top             =   300
               Width           =   1272
            End
            Begin VB.Label lblContact 
               Caption         =   "E-Mail Address"
               Height          =   192
               Index           =   13
               Left            =   240
               TabIndex        =   27
               Top             =   2568
               Width           =   1152
            End
         End
      End
   End
End
Attribute VB_Name = "frmLenderDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : LenderDetails
' Description   : Contains all tabs for the Lender details form. This form is responsible
'                 for initialising all tabs.
'
' Change history
' Prog      Date        Description
' DJP       01/08/01    SYS2564 Make Ledger Codes only visible if the tables are there.
' DJP       22/10/01    SYS2831 - change Lender to support client variants. This form now contains
'                       only events and basic validation of controls. All other functionality has been
'                       moved to the Lender class.
' STB       08/11/01    Added telephone table functionality.
' DJP       21/11/01    SYS2831/SYS2912 added field lengths to tel numbers.
' DJP       22/11/01    SYS2912 SQL Server locking problem.
' MO        10/6/02     BMIDS00040 Arrangement Fee percentages
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private m_clsLenderCS As Lender

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS History
' MO        10/6/02     BMIDS00040
' AW        27/06/02    BMIDS00093  Added NetGrossProfileInd combo
' AW        11/07/02    BMIDS00177  Removed Ledger Codes tab
' GHun      27/07/2004  BMIDS817 Reduced form size so it is all visable at 1024x768 with large fonts
' JD            29/03/2005      BMIDS982  Changed all screen text MIG to HLC
' TW        09/10/2006  EP2_7 Added Fields for Additional Borrowing Fee and Credit Limit Increase Fee
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const ARRANGEMENTFEEAMOUNT = 13
Private Const ARRANGEMENTFEEPERCENT = 11
Private Const ARRANGEMENTFEEPERCENTMIN = 15
Private Const ARRANGEMENTFEEPERCENTMAX = 12

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLender
' Description   : Sets the Lender support class this form is going to be used. All events being
'                 handled in this form will be delegated to the Lender class. It is upto the
'                 Initialisation of the Lender class call this method to initialise this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetLender(clsLender As Lender)
    Set m_clsLenderCS = clsLender
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Activate
' Description   : Called by VB when this form is activated. Need to remove any previous selections
'                 the listviews may have had.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
    On Error GoTo Failed
      
    Set lvLegalFees.SelectedItem = Nothing
    Set lvMIGRateSets.SelectedItem = Nothing
   
   Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvLegalFees_DeletePressed
' Description   : Called when Delete is pressed on the Legal Fees tab. Delegate to the Lender
'                 class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvLegalFees_DeletePressed()
    On Error GoTo Failed
    
    m_clsLenderCS.Delete LegalFees

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvMIGRateSets_DeletePressed
' Description   : Called when Delete is pressed on the MIG Rates tab. Delegate to the Lender
'                 class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvMIGRateSets_DeletePressed()
    On Error GoTo Failed
    
    m_clsLenderCS.Delete MIGRateSets
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvLegalFees_ItemClick
' Description   : Called when a row is selected on the Legal Fees listview. Delegate to the Lender
'                 class and set the state of the Add/Edit/Delete buttons
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvLegalFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    m_clsLenderCS.SetButtonsState LegalFees

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvMIGRateSets_ItemClick
' Description   : Called when a row is selected on the MIG Rates listview. Delegate to the Lender
'                 class and set the state of the Add/Edit/Delete buttons
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvMIGRateSets_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    m_clsLenderCS.SetButtonsState MIGRateSets

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SSTabLenderDetails_Click
' Description   : Called when a new tab on the Lender tabbed control is clicked. Need to perform
'                 initialisation on that tab. Delegate to the Lender class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SSTabLenderDetails_Click(PreviousTab As Integer)
    On Error GoTo Failed
    
    m_clsLenderCS.InitialiseTab SSTabLenderDetails.Tab
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtContactDetails_KeyPress
' Description   : Called when a a key is pressed on any of the contact detail text ontrols.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtContactDetails_KeyPress(nIndex As Integer, nKeyAscii As Integer)
    On Error GoTo Failed
    
    m_clsLenderCS.ValidateContactDetailsKey nIndex, nKeyAscii
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub txtLenderDetails_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not txtLenderDetails(Index).ValidateData()

End Sub
Private Sub cmdAddLegalFees_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Add LegalFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdDeleteLegalFees_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Delete LegalFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdEditLegalFees_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Edit LegalFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdMIGAdd_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Add MIGRateSets
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdMIGDelete_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Delete MIGRateSets
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdMIGEdit_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Edit MIGRateSets
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub dgOtherFees_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    On Error GoTo Failed
      
    If nCurrentRow >= 0 Then
        bCancel = Not dgOtherFees.ValidateRow(nCurrentRow)
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub dgOtherFees_Validate(Cancel As Boolean)
    On Error GoTo Failed

    If SSTabLenderDetails.Tab + 1 = OtherFees Then
        Cancel = Not dgOtherFees.DoControlValidation()
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cboAdditionalParams_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not cboAdditionalParams(Index).ValidateData()
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo Failed
    
    m_clsLenderCS.Cancel
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdOK_Click()
On Error GoTo Failed
    
    m_clsLenderCS.OK
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   MO      10/06/02    BMIDS00040
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    'validate that only an arrangement fee amount or an arrangement fee percentage can be entered
    If Val(txtLenderDetails(ARRANGEMENTFEEAMOUNT).Text) > 0 And Val(txtLenderDetails(ARRANGEMENTFEEPERCENT).Text) > 0 Then
        bRet = False
        g_clsErrorHandling.RaiseError errGeneralError, "Either an arrangement fee amount or percentage must be entered"
    End If
    
    'validate that if an arrangement fee amount has been specified then minimum and maximums havent
    If Val(txtLenderDetails(ARRANGEMENTFEEAMOUNT).Text) > 0 And (Val(txtLenderDetails(ARRANGEMENTFEEPERCENTMIN).Text) > 0 Or Val(txtLenderDetails(ARRANGEMENTFEEPERCENTMAX).Text) > 0) Then
        bRet = False
        g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum arrangement fee values can only be applied against the arrangement fee percentage."
    End If
    
    'validate that the minimum fee is not greater than the maximum fee, if a maximum has been specified
    If Len(txtLenderDetails(ARRANGEMENTFEEPERCENTMAX).Text) > 0 Then
        If Val(txtLenderDetails(ARRANGEMENTFEEPERCENTMAX).Text) < Val(txtLenderDetails(ARRANGEMENTFEEPERCENTMIN).Text) Then
            bRet = False
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum arrangement fee must be greater than or equal to the minimum fee."
        End If
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

