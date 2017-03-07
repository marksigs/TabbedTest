VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmProductDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Details"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   Icon            =   "frmProductDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7770
      TabIndex        =   133
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6450
      TabIndex        =   132
      Top             =   8760
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   64
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   12
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Product Details"
      TabPicture(0)   =   "frmProductDetails.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tied Insurances"
      TabPicture(1)   =   "frmProductDetails.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Application Eligibility 1"
      TabPicture(2)   =   "frmProductDetails.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Application Eligibility 2"
      TabPicture(3)   =   "frmProductDetails.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame30"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Application Eligibility 3"
      TabPicture(4)   =   "frmProductDetails.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame40"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame41"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame42"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Fee Sets 1"
      TabPicture(5)   =   "frmProductDetails.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame25"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame22"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame21"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Fee Sets 2"
      TabPicture(6)   =   "frmProductDetails.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fmeIASet"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "fmeProdSwitchSet"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "fmeTTSet"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Fee Sets 3"
      TabPicture(7)   =   "frmProductDetails.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame39"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Frame43"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Interest Rates"
      TabPicture(8)   =   "frmProductDetails.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame26"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Frame20"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Special Conditions"
      TabPicture(9)   =   "frmProductDetails.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame34"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Other Rates/Groups"
      TabPicture(10)  =   "frmProductDetails.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblProductDetails(31)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "Label8"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).Control(2)=   "Label5"
      Tab(10).Control(2).Enabled=   0   'False
      Tab(10).Control(3)=   "Label3"
      Tab(10).Control(3).Enabled=   0   'False
      Tab(10).Control(4)=   "lblMIGStartLTV"
      Tab(10).Control(4).Enabled=   0   'False
      Tab(10).Control(5)=   "Label4"
      Tab(10).Control(5).Enabled=   0   'False
      Tab(10).Control(6)=   "lblMIGRateSet"
      Tab(10).Control(6).Enabled=   0   'False
      Tab(10).Control(7)=   "txtERCFreePercentage"
      Tab(10).Control(7).Enabled=   0   'False
      Tab(10).Control(8)=   "cboMIGRateSet"
      Tab(10).Control(8).Enabled=   0   'False
      Tab(10).Control(9)=   "cboIncomeMultipleSet"
      Tab(10).Control(9).Enabled=   0   'False
      Tab(10).Control(10)=   "txtMPMIGStartLTV"
      Tab(10).Control(10).Enabled=   0   'False
      Tab(10).Control(11)=   "cboRedemptionFeeSet"
      Tab(10).Control(11).Enabled=   0   'False
      Tab(10).Control(12)=   "cboRentalIncomeRateSet"
      Tab(10).Control(12).Enabled=   0   'False
      Tab(10).Control(13)=   "Frame33"
      Tab(10).Control(13).Enabled=   0   'False
      Tab(10).Control(14)=   "Frame35"
      Tab(10).Control(14).Enabled=   0   'False
      Tab(10).ControlCount=   15
      TabCaption(11)  =   "Other Fees/Incentives"
      TabPicture(11)  =   "frmProductDetails.frx":0576
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame23"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "Frame24"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).ControlCount=   2
      Begin VB.Frame Frame43 
         Caption         =   "Transfer of Equity Fee"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   228
         Top             =   3480
         Width           =   7935
         Begin VB.TextBox txtTransferOfEquityFeeSelected 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1740
            TabIndex        =   231
            Top             =   1800
            Width           =   1000
         End
         Begin VB.CommandButton cmdTransferOfEquityFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   230
            Top             =   1800
            Width           =   1035
         End
         Begin VB.CommandButton cmdTransferOfEquityFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   229
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvTransferOfEquityFees 
            Height          =   1560
            Left            =   120
            TabIndex        =   232
            Top             =   240
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label Label11 
            Caption         =   "Selected Set No."
            Height          =   255
            Left            =   120
            TabIndex        =   233
            Top             =   1860
            Width           =   1335
         End
      End
      Begin VB.Frame Frame42 
         Caption         =   "Nature of Loan"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   224
         Top             =   5640
         Width           =   7935
         Begin MSGOCX.MSGHorizontalSwapList MSGSwapNatureOfLoan 
            Height          =   1695
            Left            =   240
            TabIndex        =   227
            Top             =   240
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   2990
         End
      End
      Begin VB.Frame Frame41 
         Caption         =   "Product Class"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   223
         Top             =   3360
         Width           =   7935
         Begin MSGOCX.MSGHorizontalSwapList MSGSwapProductClass 
            Height          =   1695
            Left            =   240
            TabIndex        =   226
            Top             =   240
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   2990
         End
      End
      Begin VB.Frame Frame40 
         Caption         =   "Income Status"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   222
         Top             =   1080
         Width           =   7935
         Begin MSGOCX.MSGHorizontalSwapList MSGSwapIncomeStatus 
            Height          =   1695
            Left            =   240
            TabIndex        =   225
            Top             =   240
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   2990
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Administration Fees"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   216
         Top             =   1080
         Width           =   7935
         Begin VB.CommandButton cmdAdminFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   219
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtAdminFeeSelected 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            TabIndex        =   218
            Top             =   1800
            Width           =   1000
         End
         Begin VB.CommandButton cmdAdminFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   217
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvAdminFees 
            Height          =   1560
            Left            =   72
            TabIndex        =   220
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label Label6 
            Caption         =   "Selected Set No."
            Height          =   255
            Left            =   120
            TabIndex        =   221
            Top             =   1845
            Width           =   1335
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Valuation Fees"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   210
         Top             =   3360
         Width           =   7935
         Begin VB.TextBox txtValuationFeeSelected 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            TabIndex        =   213
            Top             =   1800
            Width           =   1000
         End
         Begin VB.CommandButton cmdValuationFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   212
            Top             =   1800
            Width           =   1035
         End
         Begin VB.CommandButton cmdValuationFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   211
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvValuationFees 
            Height          =   1560
            Left            =   72
            TabIndex        =   214
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label Label7 
            Caption         =   "Selected Set No."
            Height          =   255
            Left            =   120
            TabIndex        =   215
            Top             =   1845
            Width           =   1335
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Additional Borrowing Fee Set"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   204
         Top             =   5640
         Width           =   7935
         Begin VB.TextBox txtAdditionalBorrowingFeeSelected 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            TabIndex        =   207
            Top             =   1800
            Width           =   1000
         End
         Begin VB.CommandButton cmdAdditionalBorrowingFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   206
            Top             =   1800
            Width           =   1035
         End
         Begin VB.CommandButton cmdAdditionalBorrowingFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   205
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvAdditionalBorrowingFees 
            Height          =   1560
            Left            =   72
            TabIndex        =   208
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label Label9 
            Caption         =   "Selected Set No."
            Height          =   255
            Left            =   120
            TabIndex        =   209
            Top             =   1845
            Width           =   1335
         End
      End
      Begin VB.Frame fmeTTSet 
         Caption         =   "TT Fee Sets"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   198
         Top             =   1080
         Width           =   7905
         Begin VB.CommandButton cmdTTFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   201
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtTTFeeSelected 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1740
            TabIndex        =   200
            Top             =   1798
            Width           =   1000
         End
         Begin VB.CommandButton cmdTTFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   199
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvTTFeeSet 
            Height          =   1560
            Left            =   72
            TabIndex        =   202
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label lblFeeSet2Sheet 
            AutoSize        =   -1  'True
            Caption         =   "Selected Set No."
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   203
            Top             =   1845
            Width           =   1215
         End
      End
      Begin VB.Frame fmeProdSwitchSet 
         Caption         =   "Product Switch Fee Set"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   192
         Top             =   3360
         Width           =   7905
         Begin VB.CommandButton cmdProdSwitchFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   195
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtProdSwitchFeeSelected 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1740
            TabIndex        =   194
            Top             =   1798
            Width           =   1000
         End
         Begin VB.CommandButton cmdProdSwitchFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   193
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvProdSwitchFeeSet 
            Height          =   1560
            Left            =   72
            TabIndex        =   196
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label lblFeeSet2Sheet 
            Caption         =   "Selected Set No."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   197
            Top             =   1815
            Width           =   1335
         End
      End
      Begin VB.Frame fmeIASet 
         Caption         =   "Insurance Admin Fee Sets"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   186
         Top             =   5640
         Width           =   7905
         Begin VB.CommandButton cmdIAFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   189
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtIAFeeSelected 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1740
            TabIndex        =   188
            Top             =   1798
            Width           =   1000
         End
         Begin VB.CommandButton cmdIAFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   187
            Top             =   1800
            Width           =   1035
         End
         Begin MSGOCX.MSGListView lvIAFeeSet 
            Height          =   1560
            Left            =   72
            TabIndex        =   190
            Top             =   192
            Width           =   7776
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label lblFeeSet2Sheet 
            AutoSize        =   -1  'True
            Caption         =   "Selected Set No."
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   191
            Top             =   1845
            Width           =   1215
         End
      End
      Begin VB.Frame Frame39 
         Caption         =   "Credit Limit Increase Fee"
         Height          =   2200
         Left            =   -74640
         TabIndex        =   180
         Top             =   1080
         Width           =   7935
         Begin VB.CommandButton cmdCreditLimitIncreaseFeeSelect 
            Caption         =   "&Select"
            Height          =   285
            Left            =   5610
            TabIndex        =   183
            Top             =   1800
            Width           =   1035
         End
         Begin VB.CommandButton cmdCreditLimitIncreaseFeeDeSelect 
            Caption         =   "&Deselect"
            Height          =   285
            Left            =   6735
            TabIndex        =   182
            Top             =   1800
            Width           =   1035
         End
         Begin VB.TextBox txtCreditLimitIncreaseFeeSelected 
            Enabled         =   0   'False
            Height          =   288
            Left            =   1740
            TabIndex        =   181
            Top             =   1800
            Width           =   1000
         End
         Begin MSGOCX.MSGListView lvCreditLimitIncreaseFees 
            Height          =   1560
            Left            =   120
            TabIndex        =   184
            Top             =   240
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   2752
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
         Begin VB.Label Label10 
            Caption         =   "Selected Set No."
            Height          =   255
            Left            =   120
            TabIndex        =   185
            Top             =   1860
            Width           =   1335
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Manual Interest Rate Adjustment Parameters"
         Height          =   1750
         Left            =   -74760
         TabIndex        =   173
         Top             =   4800
         Width           =   7995
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   310
            Index           =   1
            Left            =   6130
            TabIndex        =   174
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   310
            Index           =   13
            Left            =   2160
            TabIndex        =   175
            Top             =   1100
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   310
            Index           =   12
            Left            =   2160
            TabIndex        =   176
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
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
         Begin VB.Label lblInterestRateDecrease 
            Caption         =   "Minimum Resolved Rate"
            Height          =   195
            Left            =   4000
            TabIndex        =   179
            Top             =   480
            Width           =   1800
         End
         Begin VB.Label lblAdditionalParams 
            Caption         =   "Maximum Increase %"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   178
            Top             =   1100
            Width           =   1590
         End
         Begin VB.Label lblAdditionalParams 
            Caption         =   "Maximum Decrease %"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   177
            Top             =   480
            Width           =   1590
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Interest Rate Periods:"
         Height          =   3015
         Left            =   -74760
         TabIndex        =   168
         Top             =   1200
         Width           =   7995
         Begin VB.CommandButton cmdInterestRateEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4860
            TabIndex        =   171
            Top             =   2340
            Width           =   1215
         End
         Begin VB.CommandButton cmdInterestRateDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6180
            TabIndex        =   170
            Top             =   2340
            Width           =   1155
         End
         Begin VB.CommandButton cmdInterestRateAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3480
            TabIndex        =   169
            Top             =   2340
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvInterestRates 
            Height          =   1875
            Left            =   360
            TabIndex        =   172
            Top             =   360
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   3307
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "Typical APR's"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   157
         Top             =   1080
         Width           =   8100
         Begin MSGOCX.MSGDataGrid dgTypicalAPR 
            Height          =   1875
            Left            =   120
            TabIndex        =   158
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3307
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
      Begin VB.Frame Frame33 
         Caption         =   "Special Groups"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   155
         Top             =   3435
         Width           =   8115
         Begin MSGOCX.MSGDataGrid dgSpecialGroup 
            Height          =   1875
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3307
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
      Begin VB.Frame Frame24 
         Caption         =   "Incentives"
         Height          =   2772
         Left            =   -74640
         TabIndex        =   145
         Top             =   4080
         Width           =   7815
         Begin VB.CommandButton cmdIncentivesEdit 
            Caption         =   "&Edit"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4860
            TabIndex        =   148
            Top             =   2220
            Width           =   1215
         End
         Begin VB.CommandButton cmdIncentivesDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6180
            TabIndex        =   147
            Top             =   2220
            Width           =   1155
         End
         Begin VB.CommandButton cmdIncentivesAdd 
            Caption         =   "&Add"
            Height          =   375
            Left            =   3540
            TabIndex        =   146
            Top             =   2220
            Width           =   1215
         End
         Begin MSGOCX.MSGListView lvIncentives 
            Height          =   1872
            Left            =   360
            TabIndex        =   149
            Top             =   240
            Width           =   7152
            _ExtentX        =   12621
            _ExtentY        =   3307
            Sorted          =   -1  'True
            AllowColumnReorder=   0   'False
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Other Fees"
         Height          =   2775
         Left            =   -74640
         TabIndex        =   143
         Top             =   1200
         Width           =   7815
         Begin MSGOCX.MSGHorizontalSwapList MSGSwapOtherFees 
            Height          =   2310
            Left            =   360
            TabIndex        =   144
            Top             =   240
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   4075
         End
      End
      Begin VB.Frame Frame34 
         BorderStyle     =   0  'None
         Caption         =   "Frame34"
         Height          =   6555
         Left            =   -74760
         TabIndex        =   77
         Tag             =   "Tab8"
         Top             =   1100
         Width           =   8175
         Begin MSGOCX.MSGHorizontalSwapList MSGSwapProdCond 
            Height          =   3615
            Left            =   240
            TabIndex        =   167
            Top             =   240
            Width           =   7500
            _ExtentX        =   13229
            _ExtentY        =   6376
         End
      End
      Begin VB.Frame Frame30 
         BorderStyle     =   0  'None
         Caption         =   "Frame30"
         Height          =   6855
         Left            =   -74895
         TabIndex        =   120
         Tag             =   "Tab4"
         Top             =   1040
         Width           =   8295
         Begin VB.Frame Frame36 
            Caption         =   "Property Location"
            Height          =   2175
            Left            =   60
            TabIndex        =   124
            Top             =   4440
            Width           =   8235
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapPropLocation 
               Height          =   1900
               Left            =   300
               TabIndex        =   76
               Top             =   190
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   3334
            End
         End
         Begin VB.Frame Frame32 
            Caption         =   "Employment Type"
            Height          =   2175
            Left            =   60
            TabIndex        =   122
            Top             =   2200
            Width           =   8235
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapEmpElig 
               Height          =   1900
               Left            =   300
               TabIndex        =   75
               Top             =   190
               Width           =   7305
               _ExtentX        =   12885
               _ExtentY        =   3334
            End
         End
         Begin VB.Frame Frame31 
            Caption         =   "Distribution Channel"
            Height          =   2175
            Left            =   60
            TabIndex        =   121
            Top             =   0
            Width           =   8235
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapChannel 
               Height          =   1900
               Left            =   300
               TabIndex        =   74
               Top             =   190
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   3334
            End
         End
      End
      Begin VB.Frame Frame18 
         BorderStyle     =   0  'None
         Caption         =   "Frame18"
         Height          =   6730
         Left            =   -74895
         TabIndex        =   107
         Tag             =   "Tab3"
         Top             =   1040
         Width           =   8355
         Begin VB.Frame Frame29 
            Caption         =   "Type of Buyer"
            Height          =   2175
            Left            =   60
            TabIndex        =   119
            Top             =   4440
            Width           =   8250
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapTypeOfBuyer 
               Height          =   1900
               Left            =   300
               TabIndex        =   73
               Top             =   190
               Width           =   7260
               _ExtentX        =   12806
               _ExtentY        =   3334
            End
         End
         Begin VB.Frame Frame28 
            Caption         =   "Purpose of Loan"
            Height          =   2175
            Left            =   60
            TabIndex        =   118
            Top             =   2200
            Width           =   8250
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapPurpOfLoan 
               Height          =   1900
               Left            =   300
               TabIndex        =   72
               Top             =   190
               Width           =   7392
               _ExtentX        =   13044
               _ExtentY        =   3334
            End
         End
         Begin VB.Frame Frame27 
            Caption         =   "Type of Application Eligibility"
            Height          =   2175
            Left            =   60
            TabIndex        =   117
            Top             =   0
            Width           =   8250
            Begin MSGOCX.MSGHorizontalSwapList MSGSwapTypeOfAppEligibility 
               Height          =   1900
               Left            =   300
               TabIndex        =   71
               Top             =   190
               Width           =   7275
               _ExtentX        =   12832
               _ExtentY        =   3334
            End
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6660
         Left            =   -74880
         TabIndex        =   80
         Tag             =   "Tab2"
         Top             =   1200
         Width           =   8325
         Begin VB.Frame Frame19 
            Caption         =   "Parameters"
            Height          =   1635
            Left            =   0
            TabIndex        =   111
            Top             =   -15
            Width           =   8340
            Begin VB.Frame Frame13 
               BorderStyle     =   0  'None
               Caption         =   "Frame13"
               Height          =   420
               Left            =   2910
               TabIndex        =   78
               Top             =   120
               Width           =   1455
               Begin VB.OptionButton optCompulsoryBC 
                  Caption         =   "Yes"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   65
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   735
               End
               Begin VB.OptionButton optCompulsoryBC 
                  Caption         =   "No"
                  Height          =   255
                  Index           =   1
                  Left            =   840
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   975
               End
            End
            Begin VB.Frame Frame14 
               BorderStyle     =   0  'None
               Caption         =   "Frame14"
               Height          =   375
               Left            =   2805
               TabIndex        =   79
               Top             =   540
               Width           =   1635
               Begin VB.OptionButton optCompulsoryPP 
                  Caption         =   "No"
                  Height          =   255
                  Index           =   1
                  Left            =   960
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   975
               End
               Begin VB.OptionButton optCompulsoryPP 
                  Caption         =   "Yes"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   735
               End
            End
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   0
               Left            =   6375
               TabIndex        =   67
               Top             =   240
               Width           =   735
               _ExtentX        =   1296
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
               MaxLength       =   4
            End
            Begin MSGOCX.MSGEditBox txtAdditionalParams 
               Height          =   315
               Index           =   1
               Left            =   6375
               TabIndex        =   70
               Top             =   660
               Width           =   735
               _ExtentX        =   1296
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
               MaxLength       =   4
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Term"
               Height          =   195
               Index           =   16
               Left            =   5655
               TabIndex        =   116
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Months"
               Height          =   195
               Index           =   3
               Left            =   7410
               TabIndex        =   115
               Top             =   300
               Width           =   525
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Term"
               Height          =   195
               Index           =   2
               Left            =   5655
               TabIndex        =   114
               Top             =   300
               Width           =   360
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Compulsory Payment Protection?"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   113
               Top             =   660
               Width           =   2325
            End
            Begin VB.Label lblAdditionalParams 
               AutoSize        =   -1  'True
               Caption         =   "Compulsory Buildings && Contents?"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   112
               Top             =   240
               Width           =   2385
            End
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7290
         Left            =   120
         TabIndex        =   85
         Tag             =   "Tab1"
         Top             =   1000
         Width           =   8415
         Begin VB.Frame Frame44 
            BorderStyle     =   0  'None
            Caption         =   "Frame44"
            Height          =   255
            Left            =   2600
            TabIndex        =   235
            Top             =   4440
            Width           =   1515
            Begin VB.OptionButton optRefundOfValuation 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   29
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optRefundOfValuation 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdFindIntroducers 
            Caption         =   "Find Introducers"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   142
            Top             =   6765
            Width           =   1815
         End
         Begin VB.Frame fraFreeLegalFees 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   6600
            TabIndex        =   30
            Top             =   4425
            Width           =   1515
            Begin VB.OptionButton optFreeLegalFees 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optFreeLegalFees 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   32
               Top             =   0
               Width           =   555
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2600
            TabIndex        =   33
            Top             =   4755
            Width           =   1515
            Begin VB.OptionButton optMortCalcAvailable 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   35
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optMortCalcAvailable 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame10 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   6600
            TabIndex        =   36
            Top             =   4755
            Width           =   1515
            Begin VB.OptionButton optAvailableQuickQuote 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   38
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optAvailableQuickQuote 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame16 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2600
            TabIndex        =   45
            Top             =   5445
            Width           =   1515
            Begin VB.OptionButton optImpairedCredit 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   46
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optImpairedCredit 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   47
               Top             =   0
               Width           =   555
            End
         End
         Begin VB.Frame Frame12 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6600
            TabIndex        =   48
            Top             =   5445
            Width           =   1515
            Begin VB.OptionButton optFlexibleProduct 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   50
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optFlexibleProduct 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2600
            TabIndex        =   39
            Top             =   5100
            Width           =   1515
            Begin VB.OptionButton optExistingCustomer 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   41
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optExistingCustomer 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   40
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame11 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6600
            TabIndex        =   42
            Top             =   5100
            Width           =   1515
            Begin VB.OptionButton optMemberOfStaffProduct 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   44
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optMemberOfStaffProduct 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   43
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame37 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2600
            TabIndex        =   51
            Top             =   5775
            Width           =   1515
            Begin VB.OptionButton optCanBePorted 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   53
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optCanBePorted 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   2600
            TabIndex        =   57
            Top             =   6120
            Width           =   1515
            Begin VB.OptionButton optCATStandard 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   58
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optCATStandard 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   59
               Top             =   0
               Width           =   555
            End
         End
         Begin VB.CommandButton CmdAdditionalFeatures 
            Caption         =   "Additional Features"
            Height          =   315
            Left            =   3720
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   107
            Width           =   2175
         End
         Begin VB.Frame Frame38 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6600
            TabIndex        =   60
            Top             =   6120
            Width           =   1515
            Begin VB.OptionButton optExclusiveOrSemi 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   62
               Top             =   0
               Width           =   555
            End
            Begin VB.OptionButton optExclusiveOrSemi 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdMortgageProductCode 
            Caption         =   "Product Grouping"
            Height          =   315
            Left            =   5940
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   107
            Width           =   2175
         End
         Begin MSGOCX.MSGTextMulti txtDetails 
            Height          =   675
            Left            =   2280
            TabIndex        =   63
            Top             =   6480
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   1191
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
         Begin VB.Frame Frame17 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   6600
            TabIndex        =   54
            Top             =   5775
            Width           =   1515
            Begin VB.OptionButton optCashback 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   55
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optCashback 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   56
               Top             =   0
               Width           =   555
            End
         End
         Begin MSGOCX.MSGComboBox cboCountry 
            Height          =   315
            Left            =   2280
            TabIndex        =   21
            Top             =   3261
            Width           =   1692
            _ExtentX        =   2990
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
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   255
            Left            =   5160
            TabIndex        =   5
            Top             =   842
            Width           =   1680
            Begin VB.OptionButton optNonPanelLender 
               Caption         =   "No"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   7
               Top             =   0
               Width           =   570
            End
            Begin VB.OptionButton optNonPanelLender 
               Caption         =   "Yes"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   735
            End
         End
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   7
            Left            =   5160
            TabIndex        =   12
            Top             =   1531
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   8
            Left            =   2280
            TabIndex        =   13
            Top             =   1877
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   12
            Left            =   2280
            TabIndex        =   15
            Top             =   2223
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   18
            Left            =   6240
            TabIndex        =   22
            Top             =   3261
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   13
            Left            =   6240
            TabIndex        =   16
            Top             =   2223
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   14
            Left            =   2280
            TabIndex        =   17
            Top             =   2569
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   15
            Left            =   6240
            TabIndex        =   18
            Top             =   2569
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   16
            Left            =   2280
            TabIndex        =   19
            Top             =   2915
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   17
            Left            =   6240
            TabIndex        =   20
            Top             =   2915
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   3
            Left            =   7320
            TabIndex        =   10
            Top             =   1185
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
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
            Height          =   285
            Index           =   15
            Left            =   3840
            TabIndex        =   24
            Top             =   3600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
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
            MaxLength       =   6
         End
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   285
            Index           =   16
            Left            =   7320
            TabIndex        =   26
            Top             =   3600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   285
            Index           =   14
            Left            =   2280
            TabIndex        =   23
            Top             =   3600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   4
            Left            =   2280
            TabIndex        =   27
            Top             =   3960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox txtStartTime 
            Height          =   288
            Left            =   5160
            TabIndex        =   9
            Top             =   1185
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "c"
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   0
            Left            =   2280
            TabIndex        =   0
            Top             =   120
            Width           =   1395
            _ExtentX        =   2461
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
            MaxLength       =   6
         End
         Begin MSGOCX.MSGTextMulti txtProductName 
            Height          =   288
            Left            =   2280
            TabIndex        =   3
            Top             =   466
            Width           =   5835
            _ExtentX        =   10292
            _ExtentY        =   503
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
            Text            =   ""
            Mandatory       =   -1  'True
            MaxLength       =   50
         End
         Begin MSGOCX.MSGDataCombo cboLenderCode 
            Height          =   315
            Left            =   2280
            TabIndex        =   4
            Top             =   812
            Width           =   1275
            _ExtentX        =   2249
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
            Mandatory       =   -1  'True
         End
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   5
            Left            =   2280
            TabIndex        =   8
            Top             =   1185
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   6
            Left            =   2280
            TabIndex        =   11
            Top             =   1531
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSGOCX.MSGEditBox txtProductDetails 
            Height          =   288
            Index           =   10
            Left            =   6240
            TabIndex        =   14
            Top             =   1877
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
         Begin MSGOCX.MSGEditBox txtAdditionalParams 
            Height          =   285
            Index           =   17
            Left            =   5520
            TabIndex        =   25
            Top             =   3600
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
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
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Refund of Valuation?"
            Height          =   195
            Index           =   33
            Left            =   240
            TabIndex        =   234
            Top             =   4440
            Width           =   1500
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Free Legal Fees?"
            Height          =   195
            Index           =   32
            Left            =   4500
            TabIndex        =   141
            Top             =   4440
            Width           =   1230
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Available for Mortgage Calc?"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   140
            Top             =   4755
            Width           =   2040
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Available for Quick Quote?"
            Height          =   195
            Index           =   12
            Left            =   4500
            TabIndex        =   139
            Top             =   4755
            Width           =   1905
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Flexible Product?"
            Height          =   195
            Index           =   22
            Left            =   4500
            TabIndex        =   138
            Top             =   5430
            Width           =   1215
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Available for Impaired Cred?"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   137
            Top             =   5430
            Width           =   1980
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Existing Cust. Mortgage Prod.?"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   136
            Top             =   5085
            Width           =   2175
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Member of Staff Product?"
            Height          =   195
            Index           =   13
            Left            =   4500
            TabIndex        =   135
            Top             =   5085
            Width           =   1815
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "CAT Standard?"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   134
            Top             =   6135
            Width           =   1095
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Amount"
            Height          =   195
            Index           =   50
            Left            =   1680
            TabIndex        =   131
            Top             =   3630
            Width           =   540
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Arrangement Fee"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   130
            Top             =   3630
            Width           =   1215
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Exclusive/Semi-Exclusive?"
            Height          =   195
            Index           =   49
            Left            =   4500
            TabIndex        =   129
            Top             =   6135
            Width           =   1905
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Can Be Ported?"
            Height          =   195
            Index           =   48
            Left            =   240
            TabIndex        =   128
            Top             =   5760
            Width           =   1125
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Min"
            Height          =   195
            Index           =   47
            Left            =   5100
            TabIndex        =   127
            Top             =   3630
            Width           =   255
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Applicant Age"
            Height          =   192
            Index           =   20
            Left            =   4320
            TabIndex        =   126
            Top             =   1896
            Width           =   1728
         End
         Begin VB.Label lblMainProdDetails 
            AutoSize        =   -1  'True
            Caption         =   "Start Time"
            Height          =   195
            Index           =   5
            Left            =   3660
            TabIndex        =   125
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "months"
            Height          =   195
            Index           =   46
            Left            =   7320
            TabIndex        =   123
            Top             =   3315
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Porting Fee"
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   4005
            Width           =   810
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Max"
            Height          =   195
            Index           =   18
            Left            =   6840
            TabIndex        =   109
            Top             =   3630
            Width           =   300
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   17
            Left            =   3600
            TabIndex        =   108
            Top             =   3630
            Width           =   120
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Cashback Product?"
            Height          =   195
            Index           =   30
            Left            =   4500
            TabIndex        =   106
            Top             =   5760
            Width           =   1410
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Country of Origin"
            Height          =   195
            Index           =   28
            Left            =   240
            TabIndex        =   105
            Top             =   3285
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Display Order"
            Height          =   195
            Left            =   6240
            TabIndex        =   104
            Top             =   1230
            Width           =   945
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "years"
            Height          =   195
            Index           =   27
            Left            =   7320
            TabIndex        =   103
            Top             =   2265
            Width           =   375
         End
         Begin VB.Label lblProductDetails 
            Caption         =   "years"
            Height          =   195
            Index           =   3
            Left            =   3060
            TabIndex        =   102
            Top             =   2270
            Width           =   555
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Non Panel Lender"
            Height          =   195
            Index           =   21
            Left            =   3660
            TabIndex        =   101
            Top             =   840
            Width           =   1290
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Mortgage Product Code"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Top             =   135
            Width           =   1695
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Mortgage Product Name"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   99
            Top             =   480
            Width           =   1740
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Lender Code"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   98
            Top             =   840
            Width           =   915
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Start Date"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   97
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "End Date"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   96
            Top             =   1545
            Width           =   675
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Withdrawn Date"
            Height          =   195
            Index           =   6
            Left            =   3660
            TabIndex        =   95
            Top             =   1545
            Width           =   1155
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Applicant Age"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   94
            Top             =   1890
            Width           =   1650
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Term"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   93
            Top             =   2235
            Width           =   1020
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Pref. Rate Period"
            Height          =   195
            Index           =   9
            Left            =   4320
            TabIndex        =   92
            Top             =   3285
            Width           =   1215
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Details"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   91
            Top             =   6480
            Width           =   480
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Term"
            Height          =   195
            Index           =   19
            Left            =   4320
            TabIndex        =   90
            Top             =   2235
            Width           =   1065
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Minimum Loan Amount"
            Height          =   195
            Index           =   23
            Left            =   240
            TabIndex        =   89
            Top             =   2580
            Width           =   1605
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Loan Amount"
            Height          =   195
            Index           =   24
            Left            =   4320
            TabIndex        =   88
            Top             =   2580
            Width           =   1650
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Minimum LTV"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   87
            Top             =   2925
            Width           =   960
         End
         Begin VB.Label lblProductDetails 
            AutoSize        =   -1  'True
            Caption         =   "Maximum LTV"
            Height          =   195
            Index           =   26
            Left            =   4320
            TabIndex        =   86
            Top             =   2925
            Width           =   1005
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -72600
         TabIndex        =   84
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -72660
         TabIndex        =   83
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2280
         TabIndex        =   82
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6480
         TabIndex        =   81
         Top             =   5340
         Width           =   1575
      End
      Begin MSGOCX.MSGDataCombo cboRentalIncomeRateSet 
         Height          =   315
         Left            =   -72840
         TabIndex        =   150
         Top             =   7080
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSGOCX.MSGDataCombo cboRedemptionFeeSet 
         Height          =   315
         Left            =   -72840
         TabIndex        =   151
         Top             =   6720
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSGOCX.MSGEditBox txtMPMIGStartLTV 
         Height          =   330
         Left            =   -67800
         TabIndex        =   152
         Top             =   5925
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
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
      End
      Begin MSGOCX.MSGDataCombo cboIncomeMultipleSet 
         Height          =   315
         Left            =   -72840
         TabIndex        =   153
         Top             =   6315
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSGOCX.MSGDataCombo cboMIGRateSet 
         Height          =   315
         Left            =   -72840
         TabIndex        =   154
         Top             =   5925
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSGOCX.MSGEditBox txtERCFreePercentage 
         Height          =   330
         Left            =   -67800
         TabIndex        =   159
         Top             =   6690
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         TextType        =   6
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         MinValue        =   "0"
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
         MaxLength       =   3
      End
      Begin VB.Label lblMIGRateSet 
         AutoSize        =   -1  'True
         Caption         =   "HLC Rate Set"
         Height          =   195
         Left            =   -74760
         TabIndex        =   166
         Top             =   5985
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Income Multiple Set"
         Height          =   195
         Left            =   -74760
         TabIndex        =   165
         Top             =   6375
         Width           =   1395
      End
      Begin VB.Label lblMIGStartLTV 
         AutoSize        =   -1  'True
         Caption         =   "HLC Threshold LTV"
         Height          =   195
         Left            =   -69480
         TabIndex        =   164
         Top             =   5985
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Early Repayment Charge"
         Height          =   195
         Left            =   -74760
         TabIndex        =   163
         Top             =   6750
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rental Income Rate Set"
         Height          =   195
         Left            =   -74760
         TabIndex        =   162
         Top             =   7140
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ERC Free Percentage"
         Height          =   195
         Left            =   -69480
         TabIndex        =   161
         Top             =   6750
         Width           =   1560
      End
      Begin VB.Label lblProductDetails 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   31
         Left            =   -66705
         TabIndex        =   160
         Top             =   6750
         Width           =   120
      End
   End
End
Attribute VB_Name = "frmProductDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmProductDetails
' Description   : Form which allows the user to edit Mortgage Products. This form contains
'                 Multiple tabs, where there exists a supoprt class for each tab.
'
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       28/08/01    SYS2564 - Add label for Manual decrease percent so we can disable it if it
'                       isn't on the table.
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       04/12/01    SYS2831 Client variants.
' DJP/SA    21/12/01    SYS3545 Fix tab order.
' STB       10/01/02    SYS3225 Products are now keyed on StartTime.
' STB       08/03/02    SYS4158 Dbl-clicking on lvIncentives and lvInterestRates when nothing is selected - no-longer crashes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        10/05/2002  Client Specific : BMIDS Changes,
'                       Deuce Ref : BMIDS00002(MASTER) AND BMIDS00007(CHILD)
'                       For Requirement BM014("Can Be Ported") and BM019("Exclusive Or Semi Exclusive")
' GD        16/05/2002  BMIDS00009. Amended :InterestRate Tab.
'                                   Added   :None.
' GD        20/05/2002  Client Specific : BMIDS Changes,Deuce Ref : BMIDS0011
'                       For requirement BM017(part of)
' GD        22/05/2002  BMIDS0012 : Addition of Special Conditions Tab with Swap Boxes.
' GD        06/06/2002  BMIDS00016 : Slight layout change
' MO        11/06/2002  BMIDS00050 : TabIndex re-arrange
' MO        03/07/2002  BMIDS00067 and BMIDS00069 : TabIndex re-arrange and rename of MIG Start LTV to MIG Threshold LTV
' SA        20/09/2002  BMIDS00246  Removed Available for Self Employed option from first tab. It is now dealt with in other eligibility tab.
' JD        16/06/04    BMIDS765 Moved Redemption fee set combo from Other Fees tab to Misc tab. Renamed Misc tab to Other Rates/Groups.
'                                Renamed 'Redemption fee set' 'Early Repayment Charge'.
' MC        21/06/2004  BMIDS775 - Default TAB set to "Product Details".
' MC        07/07/2004  BMIDS763 - Select & Deselect toggle. disable deselect unless selected value present.
' GHun      27/07/2004  BMIDS817 - Reduced form size so it is all visable at 1024x768 with large fonts
' HMA       09/12/2004  BMIDS959 - Removed Mortgage Product Bands.
'JLD        01/03/2005  BMIDS982 - corrected spelling error
'JLD        01/03/2005  BMIDS982 - Changed all screen text MIG to HLC
' HM        23/08/2005  WP16 MAR42 - ERC Free Percentage is added on Other Rates/Groups Tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Added processing for Additional Borrowing Fee and Credit Limit Increase Fee
' GHun      17/11/2006  EP2_19 Added FreeLegalFees radio buttons
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        30/11/2006  EP2_253 - Changes related to Mortgage Product Application Eligibility
' TW        11/12/2006  EP2_20 Added processing for Transfer Of Equity Fee
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' TW        23/02/2007  EP2_1354 - DBM183 Introduce refunded valuation products.
' TW        13/04/2007  EP2_2384 - Error in supervisor when trying to get to the Introducer for Exclusive Products
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_clsMortgageProduct As MortgageProduct
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetMortgageProduct
' Description   : Sets the Mortgage Product support class this form is going to be used. All events being
'                 handled in this form will be delegated to the Mortgage Product class. It is upto the
'                 Initialisation of the Mortgage Product class call this method to initialise this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetMortgageProduct(clsMortgageProduct As MortgageProduct)
    Set m_clsMortgageProduct = clsMortgageProduct
End Sub


Private Sub cboRedemptionFeeSet_Click(Area As Integer)
    'WP16 MAR42 if all of the loan is subject to an early repayment charge then
    'ERCFreePercentage should be left blank
    If cboRedemptionFeeSet.BoundText = "<None>" Or _
       cboRedemptionFeeSet.BoundText = COMBO_NONE Or _
       Len(cboRedemptionFeeSet.BoundText) = 0 Then
        Me.txtERCFreePercentage.Enabled = False
        Me.txtERCFreePercentage.Text = vbNullString
    Else
        Me.txtERCFreePercentage.Enabled = True
    End If
    
End Sub

Private Sub cmdAdditionalBorrowingFeeDeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.DeselectItem mortAreaTabAdditionalBorrowingFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdAdditionalBorrowingFeeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.SelectItem mortAreaTabAdditionalBorrowingFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub CmdAdditionalFeatures_Click()
    'Call frmAdditionalFeatures.SetupForm
     m_clsMortgageProduct.AdditionalFeaturesPressed
     
End Sub

Private Sub cmdCreditLimitIncreaseFeeDeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.DeselectItem mortAreaTabCreditLimitIncreaseFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCreditLimitIncreaseFeeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.SelectItem mortAreaTabCreditLimitIncreaseFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdFindIntroducers_Click()
' TW 23/11/2006 EP2_172
Dim colKeys As New Collection
    On Error GoTo Failed:
    colKeys.Add txtProductDetails(0).Text       'Product Code
    colKeys.Add "Product"
' TW 13/04/2007 EP2_2384
'    colKeys.Add txtProductDetails(5).Text       'Start Date
    If Len(txtStartTime.Text) = 0 Then
        colKeys.Add txtProductDetails(5).Text       'Start Date
    Else
        colKeys.Add txtProductDetails(5).Text & " " & txtStartTime.Text      'Start Date
    End If
' TW 13/04/2007 EP2_2384 End
    colKeys.Add cboCountry.GetExtra(cboCountry.ListIndex)       'Language
    frmMaintainProductExclusivity.SetKeys colKeys
    frmMaintainProductExclusivity.SetIsEdit True
    frmMaintainProductExclusivity.Show 1
    
    Unload frmMaintainProductExclusivity
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
' TW 23/11/2006 EP2_172 End
End Sub

'*=[MC]BMIDS763

Private Sub cmdIAFeeDeSelect_Click()
    m_clsMortgageProduct.DeselectItem mortAreaTabIAFees
End Sub

Private Sub cmdProdSwitchFeeDeSelect_Click()
    m_clsMortgageProduct.DeselectItem mortAreaTabProdSwitchFees
End Sub


Private Sub cmdTransferOfEquityFeeDeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.DeselectItem mortAreaTabTransferOfEquityFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdTransferOfEquityFeeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.SelectItem mortAreaTabTransferOfEquityFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdTTFeeDeSelect_Click()
     m_clsMortgageProduct.DeselectItem mortAreaTabTTFees
End Sub

'*=SECTION END

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Activate
' Description   : Called by VB when this form is activated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
    On Error GoTo Failed
    
    '*=[MC]Enable deselect button, If item has been already selected.
    
    cmdValuationFeeDeSelect.Enabled = txtValuationFeeSelected.Text <> ""
    cmdAdminFeeDeSelect.Enabled = txtAdminFeeSelected.Text <> ""
' TW 09/10/2006 EP2_7
    cmdAdditionalBorrowingFeeDeSelect.Enabled = txtAdditionalBorrowingFeeSelected.Text <> ""
    cmdCreditLimitIncreaseFeeDeSelect.Enabled = txtCreditLimitIncreaseFeeSelected.Text <> ""
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    cmdTransferOfEquityFeeDeSelect.Enabled = txtTransferOfEquityFeeSelected.Text <> ""
' TW 11/12/2006 EP2_20 End
    cmdTTFeeDeSelect.Enabled = txtTTFeeSelected.Text <> ""
    cmdIAFeeDeSelect.Enabled = txtIAFeeSelected.Text <> ""
    cmdProdSwitchFeeDeSelect.Enabled = txtProdSwitchFeeSelected.Text <> ""
    '*=SECTION END
    
    'WP16 MAR42 if all of the loan is subject to an early repayment charge then
    'ERCFreePercentage should be left blank
    If Len(cboRedemptionFeeSet.BoundText) = 0 Then
        Me.txtERCFreePercentage.Enabled = False
        Me.txtERCFreePercentage.Text = vbNullString
    End If
    
' TW 23/11/2006 EP2_172
    cmdFindIntroducers.Visible = m_clsMortgageProduct.IsEdit
' TW 23/11/2006 EP2_172 End
    
   Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''Control Event Handlers ''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user presses the OK button. Delegates to m_clsMortgageProduct.OK
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo Failed

    m_clsMortgageProduct.OK

Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdCancel_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.Cancel
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdValuationFeeDeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.DeselectItem mortAreaTabValuationFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdAdminFeeDeSelect_Click()
On Error GoTo Failed

    m_clsMortgageProduct.DeselectItem mortAreaTabAdminFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'Private Sub cmdBaseRateDeSelect_Click()
'On Error GoTo Failed
    
    'm_clsMortgageProduct.DeselectItem mortAreaTabBaseRates
    
'Exit Sub
'Failed:
 '   g_clsErrorHandling.DisplayError
'End Sub

Private Sub cmdAdminFeeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.SelectItem mortAreaTabAdminFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'*=[MC]BMIDS763
Private Sub cmdIAFeeSelect_Click()
    
    Call m_clsMortgageProduct.SelectItem(mortAreaTabIAFees)
    
End Sub

Private Sub cmdProdSwitchFeeSelect_Click()
    Call m_clsMortgageProduct.SelectItem(mortAreaTabProdSwitchFees)
End Sub

Private Sub cmdTTFeeSelect_Click()
    Call m_clsMortgageProduct.SelectItem(mortAreaTabTTFees)
End Sub

'*=SECTION END


'Private Sub cmdBaseRateSelect_Click()
'On Error GoTo Failed
    
    'm_clsMortgageProduct.SelectItem mortAreaTabBaseRates
    
'Exit Sub
'Failed:
    'g_clsErrorHandling.DisplayError
'End Sub

Private Sub cmdValuationFeeSelect_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.SelectItem mortAreaTabValuationFees
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvAdditionalBorrowingFees_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdAdditionalBorrowingFeeSelect.Enabled = True
End Sub


Private Sub lvCreditLimitIncreaseFees_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdCreditLimitIncreaseFeeSelect.Enabled = True
End Sub


Private Sub lvIAFeeSet_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdIAFeeSelect.Enabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvIncentives_DblClick
' Description   : Delegate the call down to a class/tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvIncentives_DblClick()
        
    If Not lvIncentives.SelectedItem Is Nothing Then
        m_clsMortgageProduct.Edit mortAreaTabIncentives
    End If

End Sub


Private Sub cmdInterestRateEdit_Click()
    On Error GoTo Failed
    
    m_clsMortgageProduct.Edit mortAreaTabInterestRates
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvInterestRates_DblClick
' Description   : Delegate the call down to a class/tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvInterestRates_DblClick()
    
    If Not lvInterestRates.SelectedItem Is Nothing Then
        m_clsMortgageProduct.Edit mortAreaTabInterestRates
    End If

End Sub


Private Sub cmdIncentivesEdit_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.Edit mortAreaTabIncentives
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdInterestRateAdd_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.Add mortAreaTabInterestRates
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdIncentivesAdd_Click()
On Error GoTo Failed
    
    m_clsMortgageProduct.Add mortAreaTabIncentives
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub lvIncentives_DeletePressed()
On Error GoTo Failed
    
    m_clsMortgageProduct.Delete mortAreaTabIncentives
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub lvInterestRates_DeletePressed()
On Error GoTo Failed
        
    m_clsMortgageProduct.Delete mortAreaTabInterestRates
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvTransferOfEquityFees_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdTransferOfEquityFeeSelect.Enabled = True
End Sub


Private Sub optExclusiveOrSemi_Click(Index As Integer)
    cmdFindIntroducers.Enabled = (optExclusiveOrSemi(0).Value = True)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo Failed
    
    m_clsMortgageProduct.InitialiseTab
    
    SetTabstops Me

Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub txtERCFreePercentage_Validate(Cancel As Boolean)
    'WP16 MAR42
    Cancel = Not txtERCFreePercentage.ValidateData()
End Sub

Private Sub txtProductDetails_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not txtProductDetails(Index).ValidateData()

End Sub

' BMIDS959  Remove routines to validate Redemption fields

'*=[MD]BMIDS763 - CC075 - FEESETS
Private Sub lvAdminFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    cmdAdminFeeSelect.Enabled = True

End Sub

Private Sub lvProdSwitchFeeSet_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdProdSwitchFeeSelect.Enabled = True
End Sub

Private Sub lvTTFeeSet_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdTTFeeSelect.Enabled = True
End Sub
'*=[MC]SECTION END


Private Sub cmdIncentivesDelete_Click()
On Error GoTo Failed
        
    m_clsMortgageProduct.Delete mortAreaTabIncentives
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdInterestRateDelete_Click()
On Error GoTo Failed
    'Me.lvInterestRates.ListItems
    'If Me.lvInterestRates.ListItems.Count > 1 Then
        m_clsMortgageProduct.Delete mortAreaTabInterestRates
   ' Else
    '    MsgBox "You must have at least 1 interest rate type"
    'End If
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'Private Sub lvBaseRates_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
'    cmdBaseRateSelect.Enabled = True

'End Sub
Private Sub lvIncentives_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    cmdIncentivesEdit.Enabled = True
    cmdIncentivesDelete.Enabled = True

End Sub
Private Sub lvInterestRates_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    cmdInterestRateEdit.Enabled = True
    cmdInterestRateDelete.Enabled = True

End Sub
Private Sub lvValuationFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    cmdValuationFeeSelect.Enabled = True

End Sub
Private Sub cboLenderCode_Click(Area As Integer)
    On Error GoTo Failed
    
    m_clsMortgageProduct.LenderChange
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdMortgageProductCode_Click()
On Error GoTo Failed

    m_clsMortgageProduct.ProductGroupingPressed

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



