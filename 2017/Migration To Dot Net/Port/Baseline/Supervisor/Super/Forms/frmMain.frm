VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMain 
   Caption         =   "Supervisor"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   10320
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   10320
   Begin VB.Timer timerCancelForm 
      Enabled         =   0   'False
      Left            =   6360
      Top             =   720
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5625
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   936
      TabIndex        =   5
      Top             =   705
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   720
   End
   Begin MSComctlLib.TreeView tvwDB 
      Height          =   4800
      Left            =   15
      TabIndex        =   4
      Top             =   735
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8467
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   512
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlSmallIcons"
      Appearance      =   1
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   10320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10320
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   2205
         TabIndex        =   3
         Tag             =   " ListView:"
         Top             =   75
         Width           =   2850
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Supervisor functions"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   " TreeView:"
         Top             =   75
         Width           =   2190
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0554
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0666
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0778
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":088A
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":099C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AAE
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC0
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD2
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DE4
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EF6
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   6840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1008
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12CA
            Key             =   "Options"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158C
            Key             =   "INDIVIDUAL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DE
            Key             =   "LEADER"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E30
            Key             =   "COMPANY"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2282
            Key             =   "ADMINISTRATION"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26D4
            Key             =   "closed"
         EndProperty
      EndProperty
   End
   Begin MSGOCX.MSGListView lvListView 
      Height          =   4755
      Left            =   3000
      TabIndex        =   6
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8387
      Sorted          =   -1  'True
      AllowColumnReorder=   0   'False
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12091
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "28/11/2007"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "09:03"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2805
      MousePointer    =   9  'Size W E
      Top             =   720
      Width           =   120
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "New"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuFileNewComboEntry 
            Caption         =   "Combobox Entry"
         End
         Begin VB.Menu mnuFileNewGlobalParameter 
            Caption         =   "GlobalParameter"
            Begin VB.Menu mnuFileNewGlobalParameterFixed 
               Caption         =   "Fixed"
            End
            Begin VB.Menu mnuFileNewGlobalParameterBanded 
               Caption         =   "Banded"
            End
         End
         Begin VB.Menu mnuFileNewRates 
            Caption         =   "Rates and Fees"
            Begin VB.Menu mnuFileNewRatesAdminFee 
               Caption         =   "Admin Fee"
            End
            Begin VB.Menu mnuFileNewRatesValuationFee 
               Caption         =   "Valuation Fee"
            End
            Begin VB.Menu mnuFileNewRatesBaseRateSet 
               Caption         =   "Base Rate Sets"
            End
            Begin VB.Menu mnuFileNewRatesBaseRate 
               Caption         =   "Base Rate"
            End
            Begin VB.Menu mnuFileNewRatesRentalIncomeRateSet 
               Caption         =   "Rental Income Rate Sets"
            End
            Begin VB.Menu mnuFileNewIncomeMultiple 
               Caption         =   "Income Multiples"
            End
            Begin VB.Menu mnuFileNewMPMigRate 
               Caption         =   "Product MIG Rates"
            End
            Begin VB.Menu mnuFileNewRedemptionFee 
               Caption         =   "Redemption Fees"
            End
            Begin VB.Menu mnuFileNewRatesProdSwitchFee 
               Caption         =   "Product Switch Fee"
            End
            Begin VB.Menu mnuFileNewRatesInsuAdminFeeSet 
               Caption         =   "Insurance Admin Fee Set"
            End
         End
         Begin VB.Menu mnuFileNewLender 
            Caption         =   "Lender"
         End
         Begin VB.Menu mnuFileNewBatchProcessing 
            Caption         =   "Batch Processing"
            Begin VB.Menu mnuFileNewBatchScheduler 
               Caption         =   "Batch Scheduler"
            End
            Begin VB.Menu mnuFileNewPaymentsForCompletion 
               Caption         =   "Payments for Completion"
            End
         End
         Begin VB.Menu mnuFileNewProducts 
            Caption         =   "Products"
            Begin VB.Menu mnuFileNewProductsProduct 
               Caption         =   "Mortgage Product"
            End
            Begin VB.Menu mnuFileNewProductsLifeCoverRates 
               Caption         =   "Life Cover Rates"
            End
            Begin VB.Menu mnuFileNewProductsBAndC 
               Caption         =   "Buildings && Contents Products"
            End
            Begin VB.Menu mnuFileNewProductsPayProtRates 
               Caption         =   "Payment Protection Rates"
            End
            Begin VB.Menu mnuFileNewProductsPayProtProducts 
               Caption         =   "Payment Protection Products"
            End
         End
         Begin VB.Menu mnuFileNewOrg 
            Caption         =   "Organisation"
            Begin VB.Menu mnuFileNewOrgCountry 
               Caption         =   "Country"
            End
            Begin VB.Menu mnuFileNewOrgDistChannel 
               Caption         =   "Distribution Channel"
            End
            Begin VB.Menu mnuFileNewOrgDepartment 
               Caption         =   "Department"
            End
            Begin VB.Menu mnuFileNewOrgRegion 
               Caption         =   "Region"
            End
            Begin VB.Menu mnuFileNewOrgUnit 
               Caption         =   "Unit"
            End
            Begin VB.Menu mnuFileNewOrgCompetencies 
               Caption         =   "Competencies"
            End
            Begin VB.Menu mnuFileNewOrgWorkingHours 
               Caption         =   "Working Hours"
            End
            Begin VB.Menu mnuFileNewOrgUser 
               Caption         =   "User"
            End
         End
         Begin VB.Menu mnuFileNewTP 
            Caption         =   "Third Parties"
            Begin VB.Menu mnuFileNewTPPanel 
               Caption         =   "Panel"
               Begin VB.Menu mnuFileNewTPLegalRep 
                  Caption         =   "Legal Rep."
               End
               Begin VB.Menu mnuFileNewTPValuer 
                  Caption         =   "Valuer"
               End
            End
            Begin VB.Menu mnuFileNewTPOther 
               Caption         =   "Other"
            End
         End
         Begin VB.Menu mnuFileNewIncomeFactor 
            Caption         =   "Income Factors"
         End
         Begin VB.Menu mnuFileNewErrorMessage 
            Caption         =   "Error Message"
         End
         Begin VB.Menu FileNewAppLock 
            Caption         =   "Lock"
            Begin VB.Menu mnuFileNewOnlineAppLock 
               Caption         =   "Online Application Lock"
            End
            Begin VB.Menu mnuFileNewOfflineAppLock 
               Caption         =   "Offline Application Lock"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuFileNewTaskManage 
            Caption         =   "Task Management"
            Begin VB.Menu mnuFileNewTaskManageTask 
               Caption         =   "Tasks"
            End
            Begin VB.Menu mnuFileNewTaskManageStage 
               Caption         =   "Stages"
            End
            Begin VB.Menu mnuFileNewTaskManageActivity 
               Caption         =   "Activities"
            End
            Begin VB.Menu mnuFileNewTaskManageCase 
               Caption         =   "Case Tracking"
               Begin VB.Menu mnuFileNewTaskManageCaseBusiness 
                  Caption         =   "Business Groups"
               End
            End
         End
         Begin VB.Menu mnuFileNewQuestion 
            Caption         =   "Question"
         End
         Begin VB.Menu mnuFileNewCondition 
            Caption         =   "Condition"
         End
         Begin VB.Menu mnuFileNewBusinessGroup 
            Caption         =   "Business Group"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileNewPrinting 
            Caption         =   "Printing"
            Begin VB.Menu mnuFileNewPrintingTemplate 
               Caption         =   "Printing Template"
            End
            Begin VB.Menu mnuFileNewPrintingDocument 
               Caption         =   "Document Locations"
            End
            Begin VB.Menu mnuFileNewPrintingPack 
               Caption         =   "Printing Pack"
            End
         End
         Begin VB.Menu mnuFileNewCurrencies 
            Caption         =   "Currencies"
         End
         Begin VB.Menu mnuStub 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "&View"
      Index           =   2
      Begin VB.Menu mnuView 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "&Tools"
      Index           =   3
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsDatabaseOptions 
         Caption         =   "Database Options"
      End
      Begin VB.Menu mnuToolsManagementOptions 
         Caption         =   "Database Promotions"
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "&Help"
      Index           =   4
      Begin VB.Menu mnuHelp 
         Caption         =   "&About "
         Index           =   3
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Lenders"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Contact Details"
         Index           =   2
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Additional Parameters"
         Index           =   3
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Legal Fees"
         Index           =   4
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Other Fees"
         Index           =   5
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "MIG Rate Sets"
         Index           =   6
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Ledger Codes"
         Index           =   7
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuLVLenders 
         Caption         =   "Mark for promotion"
         Index           =   9
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Products"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu mnuLVProducts 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuLVProducts 
         Caption         =   "Copy  Ctrl+C"
         Index           =   1
      End
      Begin VB.Menu mnuLVProducts 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu mnuLVProducts 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuLVProducts 
         Caption         =   "Mark for promotion"
         Index           =   4
      End
      Begin VB.Menu mnuLVProducts 
         Caption         =   "View Errors"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "EditDelete"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu mnuLVEdit 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuLVEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLVEdit 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu mnuLVEdit 
         Caption         =   "Mark for promotion"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Options"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Add New"
         Index           =   0
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Find"
         Index           =   1
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Retrieve All"
         Index           =   2
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Retrieve Omiga Units"
         Index           =   3
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Retrieve Intro Units"
         Index           =   4
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Retrieve Omiga Users"
         Index           =   5
      End
      Begin VB.Menu mnuTVOptions 
         Caption         =   "Retrieve Introducer Users"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Panel Addresses"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Activate"
         Index           =   1
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Contact Details"
         Index           =   3
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Legal Rep Details"
         Index           =   4
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Valuer Details"
         Index           =   5
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Bank Details"
         Index           =   6
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Delete"
         Index           =   8
      End
      Begin VB.Menu mnuLVAddress 
         Caption         =   "Mark for promotion"
         Index           =   9
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Users"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu mnuLVUsers 
         Caption         =   "Edit"
         Index           =   0
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "Qualifications"
         Index           =   2
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "Competency History"
         Index           =   3
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "User Roles"
         Index           =   4
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLVUsers 
         Caption         =   "Mark for promotion"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Locks"
      Index           =   12
      Visible         =   0   'False
      Begin VB.Menu mnuLVLocks 
         Caption         =   "Find"
         Index           =   0
      End
      Begin VB.Menu mnuLVLocks 
         Caption         =   "Retrieve all"
         Index           =   1
      End
      Begin VB.Menu mnuLVLocks 
         Caption         =   "Create"
         Index           =   2
      End
      Begin VB.Menu mnuLVLocks 
         Caption         =   "Remove"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Application Processing"
      Index           =   13
      Visible         =   0   'False
      Begin VB.Menu mnuLVAppProcessing 
         Caption         =   "Cancel Application"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Batch Scheduler"
      Index           =   14
      Visible         =   0   'False
      Begin VB.Menu mnuLVBatch 
         Caption         =   "View"
         Index           =   0
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "Cancel Batch"
         Index           =   3
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "Restart Batch"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "View Scheduled Jobs"
         Index           =   6
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuLVBatch 
         Caption         =   "Launch Batch"
         Index           =   8
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Intermediaries"
      Index           =   15
      Visible         =   0   'False
      Begin VB.Menu mnuLVIntermediaries 
         Caption         =   "Add New"
         Index           =   0
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Administration Centre"
            Index           =   0
         End
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Company"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Lead Agent"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Individual"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Packager"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuLVIntermediariesAdd 
            Caption         =   "Broker"
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuLVIntermediaries 
         Caption         =   "Refresh"
         Index           =   1
      End
      Begin VB.Menu mnuLVIntermediaries 
         Caption         =   "Find"
         Enabled         =   0   'False
         Index           =   2
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "Third Parties"
      Index           =   16
      Visible         =   0   'False
      Begin VB.Menu mnuLVThirdParties 
         Caption         =   "Add New"
         Index           =   0
      End
      Begin VB.Menu mnuLVThirdParties 
         Caption         =   "Find"
         Index           =   1
      End
      Begin VB.Menu mnuLVThirdParties 
         Caption         =   "Retrieve All"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "EpsomIntermediaries"
      Index           =   17
      Visible         =   0   'False
      Begin VB.Menu mnuLVEpsomIntermediaries 
         Caption         =   "Add new Packager"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLVEpsomIntermediaries 
         Caption         =   "Add new Broker (Organisation)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLVEpsomIntermediaries 
         Caption         =   "Add new Broker (Individual)"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLVEpsomIntermediaries 
         Caption         =   "Find"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTopLevel 
      Caption         =   "BaseRates"
      Index           =   18
      Visible         =   0   'False
      Begin VB.Menu mnuLVBaseRates 
         Caption         =   "View"
         Index           =   0
      End
      Begin VB.Menu mnuLVBaseRates 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLVBaseRates 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu mnuLVBaseRates 
         Caption         =   "Mark for Promotion"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmMain
' Description   :   Startup form for Supervisor
'
'History:
'
' Prog      Date        Description
' DJP       09/11/00    Phase 2 Task Management
' CL        24/04/01    Modification for inclusion of frmEditCurrencies
' DJP       26/06/01    SQL Server port
' DJP       06/08/01    Added code to disable currencies if tables don't exist.
' DJP       22/11/01    SYS2912 SQL Server locking problem.
' DJP       09/01/02    SYS2831 Support client variants
' STB       21/01/02    SYS2957 Supervisor Security Enhancement.
' STB       04/02/02    SYS1535 Certain items had delete enabled which
'                       shouldn't.
' SDS       12/02/02    SYS4030 Level of Security Access selected causing error whilst displaying menus
' SDS       12/02/02    SYS4031 Above AQR (SYS4030) was the source of error.
' SDS       12/02/02    SYS4032 Delete enabled for Task Management
' SDS       12/02/02    SYS2058 Missing items from File, New Menu
' SDS/DJP   22/02/02    SYS4085 Added Base Rate to TreeView
' DJP       22/02/02    SYS4142 Add TreeView/ListView inheritence and make EnableMenu public
' DJP       24/02/02    SYS4149 Security client variants.
' SDS       06/02/02    SYS4031 Added Batch Scheduler,Currencies to File, New Menu
' STB       08/03/02    SYS4238 Replaced BaseRate column with RateDifference for Base Rate Sets.
' GHun      26/04/02    SYS1619 General functional error left mouse double click allows edit.
' STB       08/05/02    SYS4531 The 'Options' menu item has been made invisible.
' STB       13/05/02    SYS4417 Added AllowableIncomeFactors.
' JR        27/05/02    SYS4537 Allow Client Projects to re-label TreeViewText
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'
'   AW     13/05/02    BM088        Income Multiples
'   AW     15/05/02    BM023        Extra columns for valuation fees
'   AW     21/05/02    BM087        Extra columns for Mig rates
'   AW     21/05/02    BM017        Extra columns for Redemption Fee Sets
'   MO     21/05/02    BMIDS00054   New menu items added to LVPRODUCTS
'   DB     04/11/02    BMIDS00720   Re-ordered the way that third parties are displayed.
'   DB     06/11/02    BMIDS00201   Amended rt click menu to disable and enable promote & view errors.
'   DB     31/01/03    SYS5724      Re-ordered the way that mortgage products are displayed.
'   DJP    22/02/03    BM0318       Added searching for Third Parties
'   DJP    24/02/03    BM0318       Added searching for Third Parties - remove selection on expand/collapse
'   DJP    24/02/03    BM0318       Rollback SYS5724
'   DJP    06/03/03    BM0282       Added searching for Mortgage Product, Conditions, Users
'   IK     07/03/03    BM0314       Added remove (DMS) Document Locks
'   BS     25/03/03     BM0282 Move listview lbltitle.Caption refresh
'   BS      26/03/03    BM0311  Change order of mortgage product columns
'   BS      20/05/03    BM0240  Change the About dialog to show App.Comments
'   JD      14/06/04    BMIDS765    CC076 add rental income interest rates.
'   JD      30/03/2005  BMIDS982    Changed screen text from Mig to HLC
'   JD      06/04/2005  BMIDS982    Changed screen text from Redemption Fee to ERC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Specific History:
'
' GHun      27/07/2005  MAR16   Remove File New menu option and move edit menu options under file
' GHun      16/08/2005  MAR45   Apply BBG1370 (New screen for print configuration)
' PJO       28/11/2005  MAR81   Solicitor Panel Maintenance
' TK        30/11/2005  MAR81   Solicitor Panel Maintenance
' RF        05/12/2005  MAR202  Complete GHun's changes to handle Packs
' JD        05/01/2006  MAR990  SetUpTreeValues. Make Valuer the same as LegalRep as they use the same stored procedure.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Specific History:
'
' PB        09/05/2006  EP521   Mis-spelling of Commission as "Comission"
' TW        09/10/2006  EP2_7   Added handling of Additional Borrowing Fee set and
'                               Credit Limit Increase Fee Set
' PB        17/10/2006  EP2_14  E2CR20 - Unit organisation changes
' TW        17/10/2006  EP2_15  Modifications for new classes
' TW        11/12/2006  EP2_20  WP1 - Loans & Products Supervisor Changes part 3
' TW        14/12/2006  EP2_518 Procuration Fees Support
' TW        18/12/2006  EP2_568 Add functionality to select which Omiga Users are returned
' TW        20/12/2006  EP2_615 Show/Hide "Intermediaries" Option on Navigation Tree
' TW        02/01/2007  EP2_640 Tidy up consistency of operation
' TW        15/01/2007  EP2_826 - Rationalisation of pop-up menus and actions to improve consistency and usability
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        29/01/2007  EP2_863 - Add new col to IntroducerFirm - "Trading As"
' TW        01/02/2007  EP2_1036 - Principal Firms/Network display and search (Follow-on)
' TW        03/02/2007  EP2_1101 - Promotion inconsistencies
' TW        14/02/2007  EP2_1366 - Error amending AR Broker records
' TW        19/02/2007  EP2_1036(2) - Remove the FSA Ref column from the List displays for Individual AR Brokers and Individual DA Brokers
' TW        18/05/2007  VR262 Show Description field on Base Rate in Supervisor main form
' TW        27/11/2007  DBM594 - Add new functionality "Payments for Completion"
'-------------------------------------------------------------------------------
 
Option Explicit
'MAR16 GHun Not used
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, _
'    ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'MAR16 End

' Private data
Private m_treeData As New Collection
Private m_nodeCurrent As MSComctlLib.node
Private Const TOOLBAR_COPY As Integer = 1
Private Const TOOLBAR_DELETE As Integer = 2
Private Const TREE_WIDTH As Single = 2800
Dim mbMoving As Boolean

Private Const MENU_EDIT As Integer = 0
Private Const MENU_COPY As Integer = 1
Private Const MENU_DELETE As Integer = 2

Private Const constRightGap As Single = 200
Private Const constTreeLabel As Integer = 0
Private Const constListViewLabel As Integer = 1
Private m_ReturnCode As MSGReturnCode
Private m_bUnloading As Boolean
Const sglSplitLimit As Single = 2800

'PB 17/10/2006 E2CR?? Begin
Private m_intDetail As LV_DETAIL
'PB End


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Called when the form is first loaded. Only happens once for
'                   this form.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bLoginSuccessful As Boolean
    
    m_bUnloading = False
    g_bFirstTimeExecuted = False
    SetReturnCode MSGFailure
        
    Set g_Timer = timerCancelForm
    g_Timer.Interval = 1
    g_clsErrorHandling.SetErrorForm frmDisplayError
    'BS BM0240 20/05/03
    'g_sVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    g_sVersion = "Version: " & App.Comments

    bRet = g_clsMainSupport.HandleStartup()
      
    If bRet Then
        
        On Error GoTo Failed
        
        ' Load at previous loaded position
        LoadPosition
        
        ' Initialise the listview control
        InitListView
        
        ' Setup the TreeView with all nodes required
        
        ' This may fail,
        Set g_clsVersion = New VersionControlTable
        
        ' The user may have to login:
        bLoginSuccessful = g_clsMainSupport.UserLogin(Me)
        
        Unload frmLogin
        
        'Create a security server to check for user access.
        If InitialiseSecurity And bLoginSuccessful Then
            ' Only build the parts of the treeview that are needed
            g_clsMainSupport.ShowSupervisorObjects
        Else
            ExitSupervisor
        End If
    Else
        m_bUnloading = True
        Unload Me
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Unload
' Description   :   Called when this form is closed. Close all other forms, and
'                   save state
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Dim nRet As Integer
    
    If m_bUnloading = False Then
        nRet = MsgBox("Exit Supervisor?", vbYesNo + vbQuestion)
            
        If nRet = vbYes Then
            ExitSupervisor
        Else
            Cancel = 1
        End If
    End If
End Sub
Public Sub ExitSupervisor()
    
    m_bUnloading = True
    g_clsFormProcessing.CancelForm Me
    Unload frmSplash
    
    'Release object references.
    Set g_clsSecurityMgr = Nothing
    
    'close all sub forms
    Dim i As Integer
    
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
    'End
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If frmMain.WindowState <> vbMinimized Then
        'If Me.Width < 10000 Then Me.Width = 10000
        SizeControls imgSplitter.Left
    End If
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If

    'SizeControls imgSplitter.Left

End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub SizeControls(X As Single)
    On Error Resume Next

    If X < TREE_WIDTH Then X = TREE_WIDTH
    If X > (Me.Width - TREE_WIDTH) Then X = Me.Width - TREE_WIDTH
    
    tvwDB.Width = X - 40
    
' TW 15/01/2007 EP2_826
    sbStatusBar.Panels(1).Width = tvwDB.Width
' TW 15/01/2007 EP2_826 End

    imgSplitter.Left = X
    lvListView.Left = X + 30
    lvListView.Width = Me.Width - (tvwDB.Width) - constRightGap
   
    lblTitle(constTreeLabel).Left = tvwDB.Left
    lblTitle(constTreeLabel).Width = tvwDB.Width ' + 10
    
    lblTitle(constListViewLabel).Left = lvListView.Left + 45
    
    ' The 70 below makes the listview title match up to the right edge. Not sure
    ' why it doesn't match using just the width...
    
    lblTitle(constListViewLabel).Width = lvListView.Width - 120
    lblTitle(constListViewLabel).Top = lblTitle(constTreeLabel).Top
    'set the top
' TW 15/01/2007 EP2_826
'    If tbToolBar.Visible Then
'        tvwDB.Top = tbToolBar.Height + picTitles.Height
'    Else
'        tvwDB.Top = picTitles.Height
'    End If
    tvwDB.Top = picTitles.Height
' TW 15/01/2007 EP2_826 End

    lvListView.Top = tvwDB.Top

    'set the height
    If sbStatusBar.Visible Then
        tvwDB.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
    Else
        tvwDB.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
    End If

    lvListView.Height = tvwDB.Height
    imgSplitter.Top = tvwDB.Top
    imgSplitter.Height = tvwDB.Height

    Dim vKey As Variant
    vKey = frmMain.GetSelectedTreeKey()

    If (Not IsNull(vKey)) Then
        SetColumnHeader CStr(vKey)
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   lvListView_DblClick
' Description   :   Called when the ListView is double clicked. Treat as
'                   though the default right click option was clicked
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvListView_DblClick()
    On Error GoTo Failed
    'SYS1619 call HandleListViewPopup to perform the default right click action (which is not always edit)
    'g_clsMainSupport.ListEdit
    HandleListViewPopup True
    'SYS1619 End
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvListView_DeletePressed()
    On Error GoTo Failed
    g_clsMainSupport.DeleteRecord
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub lvListView_GotFocus()
    On Error GoTo Failed

    g_clsMainSupport.SelectItem lvListView, SET_SELECTION

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

' Called when a selection from the listview is clicked on
Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
' TW 15/01/2007 EP2_826
'    EnableEdit
' TW 15/01/2007 EP2_826 End
    g_clsMainSupport.SelectItem lvListView
End Sub
Private Sub mnuFileExit_Click()
'    ExitSupervisor
' TW 15/01/2007 EP2_826
    m_bUnloading = True ' Prevent daft question !!
' TW 15/01/2007 EP2_826 End
    Unload frmMain
End Sub

Private Sub mnuFileNewBusinessGroup_Click()
    g_clsMainSupport.HandleTreeNew BUSINESS_GROUPS
End Sub

Private Sub mnuFileNewCondition_Click()
    g_clsMainSupport.HandleTreeNew CONDITIONS
End Sub

Private Sub mnuFileNewGlobalParameterBanded_Click()
    g_clsMainSupport.HandleTreeNew GLOBAL_PARAM_BANDED
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MENU CLICK FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuFileNewErrorMessage_Click()
    g_clsMainSupport.HandleTreeNew ERROR_MESSAGES
End Sub
Private Sub mnuFileNewComboEntry_Click()
    g_clsMainSupport.HandleTreeNew COMBOBOX_ENTRIES
End Sub
Private Sub mnuFileNewBatchScheduler_Click()
    g_clsMainSupport.HandleTreeNew BATCH_SCHEDULER
End Sub

Private Sub mnuFileNewIncomeFactor_Click()
    g_clsMainSupport.HandleTreeNew INCOME_FACTORS
End Sub

Private Sub mnuFileNewLender_Click()
    g_clsMainSupport.HandleTreeNew LENDERS
End Sub
Private Sub mnuFileNewGlobalParameterFixed_Click()
    g_clsMainSupport.HandleTreeNew GLOBAL_PARAM_FIXED
End Sub
Private Sub mnuFileNewOfflineAppLock_Click()
    'g_clsMainSupport.CreateLock Customer
    g_clsMainSupport.CreateLock Application
End Sub
Private Sub mnuFileNewOnlineAppLock_Click()
    g_clsMainSupport.CreateLock Application
End Sub
Private Sub mnuFileNewOrgCompetencies_Click()
    g_clsMainSupport.HandleTreeNew COMPETENCIES
End Sub
Private Sub mnuFileNewOrgCountry_Click()
    g_clsMainSupport.HandleTreeNew COUNTRIES
End Sub

Private Sub mnuFileNewOrgDepartment_Click()
    g_clsMainSupport.HandleTreeNew DEPARTMENTS
End Sub
Private Sub mnuFileNewOrgDistChannel_Click()
    g_clsMainSupport.HandleTreeNew DISTRIBUTION_CHANNELS
End Sub
Private Sub mnuFileNewOrgRegion_Click()
    g_clsMainSupport.HandleTreeNew REGIONS
End Sub
Private Sub mnuFileNewOrgUnit_Click()
    g_clsMainSupport.HandleTreeNew UNITS
End Sub

Private Sub mnuFileNewOrgUser_Click()
    g_clsMainSupport.HandleTreeNew USERS
End Sub
Private Sub mnuFileNewOrgWorkingHours_Click()
    g_clsMainSupport.HandleTreeNew WORKING_HOURS
End Sub

Private Sub mnuFileNewCurrencies_Click()
    g_clsMainSupport.HandleTreeNew CURRENCIES
End Sub

'MAR45 GHun
Private Sub mnuFileNewPrintingDocument_Click()
    g_clsMainSupport.HandleTreeNew PRINTING_DOCUMENT
End Sub
'MAR45 End

'MAR202 GHun
Private Sub mnuFileNewPrintingPack_Click()
    g_clsMainSupport.HandleTreeNew PRINTING_PACK
End Sub
'MAR202 End


Private Sub mnuFileNewPrintingTemplate_Click()
    g_clsMainSupport.HandleTreeNew PRINTING_TEMPLATE
End Sub

Private Sub mnuFileNewProductsBAndC_Click()
    g_clsMainSupport.HandleTreeNew BUILDINGS_AND_CONTENTS_PRODUCTS
End Sub
Private Sub mnuFileNewProductsLifeCoverRates_Click()
    g_clsMainSupport.HandleTreeNew LIFE_COVER_RATES
End Sub
Private Sub mnuFileNewProductsPayProtProducts_Click()
    g_clsMainSupport.HandleTreeNew PAYMENT_PROTECTION_PRODUCTS
End Sub
Private Sub mnuFileNewProductsPayProtRates_Click()
    g_clsMainSupport.HandleTreeNew PAYMENT_PROTECTION_RATES
End Sub
Private Sub mnuFileNewProductsProduct_Click()
    g_clsMainSupport.HandleTreeNew MORTGAGE_PRODUCTS
End Sub

Private Sub mnuFileNewQuestion_Click()
    g_clsMainSupport.HandleTreeNew ADDITIONAL_QUESTIONS
End Sub

Private Sub mnuFileNewRatesAdminFee_Click()
    g_clsMainSupport.HandleTreeNew ADMIN_FEES
End Sub
Private Sub mnuFileNewRatesBaseRate_Click()
    g_clsMainSupport.HandleTreeNew BASE_RATE
End Sub

Private Sub mnuFileNewRatesBaseRateSet_Click()
    g_clsMainSupport.HandleTreeNew BASE_RATES
End Sub

Private Sub mnuFileNewRatesInsuAdminFeeSet_Click()
    g_clsMainSupport.HandleTreeNew INSURANCE_ADMIN_FEESETS
End Sub

Private Sub mnuFileNewRatesProdSwitchFee_Click()
    g_clsMainSupport.HandleTreeNew PRODUCT_SWITCH_FEESETS
End Sub
Private Sub mnuFileNewRatesRentalIncomeRateSet_Click()
    g_clsMainSupport.HandleTreeNew RENTAL_INCOME_RATES
End Sub

Private Sub mnuFileNewRatesValuationFee_Click()
    g_clsMainSupport.HandleTreeNew VALUATION_FEES
End Sub
Private Sub mnuFileNewTPLegalRep_Click()
    g_clsMainSupport.HandleTreeNew LEGAL_REP_ADDRESS
End Sub
Private Sub mnuFileNewTPOther_Click()
    g_clsMainSupport.HandleTreeNew LOCAL_ADDRESS
End Sub
Private Sub mnuFileNewTPValuer_Click()
    g_clsMainSupport.HandleTreeNew VALUER_ADDRESS
End Sub
Private Sub mnuFileNewTaskManageTask_Click()
    g_clsMainSupport.HandleTreeNew TASK_MANAGEMENT_TASKS
End Sub
Private Sub mnuFileNewTaskManageStage_Click()
    g_clsMainSupport.HandleTreeNew TASK_MANAGEMENT_STAGES
End Sub
Private Sub mnuFileNewTaskManageActivity_Click()
    g_clsMainSupport.HandleTreeNew TASK_MANAGEMENT_ACTIVITIES
End Sub
Private Sub mnuFileNewIncomeMultiple_Click()
    g_clsMainSupport.HandleTreeNew INCOME_MULTIPLE
End Sub
Private Sub mnuFileNewRedemptionFee_Click()
    g_clsMainSupport.HandleTreeNew REDEM_FEE_SETS
End Sub
Private Sub mnuFileNewMPMigRate_Click()
    g_clsMainSupport.HandleTreeNew MP_MIG_RATE_SETS
End Sub
Private Sub mnuFileNewTaskManageCaseBusiness_Click()
    g_clsMainSupport.HandleTreeNew BUSINESS_GROUPS
End Sub
Private Sub mnuHelp_Click(Index As Integer)
    MsgBox g_sVersion, vbInformation, "Supervisor - About"
End Sub
Private Sub mnuLVAppProcessing_Click(Index As Integer)
    On Error GoTo Failed
        
    g_clsMainSupport.HandleAppProcessing Index
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub mnuLVBaseRates_Click(Index As Integer)
    Timer1.Interval = 1
    Timer1.Enabled = True
        
    Select Case Index
        Case LV_BASERATE_VIEW
            Timer1.Tag = "ListViewEdit"
        Case LV_BASERATE_DELETE
            Timer1.Tag = "ListViewDelete"
    
        Case LV_BASERATE_MARK_FOR_PROMOTION
            Timer1.Tag = "ListViewPromote"
    End Select
End Sub

Private Sub mnuLVBatch_Click(Index As Integer)
    
    On Error GoTo Failed
        
    g_clsMainSupport.HandleBatchProcessing Index
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub mnuLVIntermediaries_Click(Index As Integer)
    On Error GoTo Failed
    Dim clsIntermediary As Intermediary
    
    Set clsIntermediary = New Intermediary
    
    Select Case Index
    
        Case LV_INTERMEDIARY_MENU_FIND
            clsIntermediary.HandleIntermediarySearch
        Case LV_INTERMEDIARY_MENU_REFRESH
            clsIntermediary.HandleIntermediaryRefresh
    End Select
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub mnuLVIntermediariesAdd_Click(Index As Integer)
    On Error GoTo Failed
    
    g_clsMainSupport.HandleIntermediaries Index
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub mnuLVLocks_Click(Index As Integer)
    On Error GoTo Failed
    
    g_clsMainSupport.HandleLocks Index
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub mnuLVProducts_Click(Index As Integer)
    g_clsMainSupport.HandleProducts Index, lvListView
End Sub
Private Sub mnuToolsDatabaseOptions_Click()
    On Error GoTo Failed
    ' DJP SQL Server port - nothing to do with it really, but it's about time this happened:
    BeginWaitCursor
    
    g_clsMainSupport.CloseFindDialogs
    frmDatabaseOptions.Show vbModal, Me

    If frmDatabaseOptions.GetReturnCode() = MSGSuccess Then
        g_clsMainSupport.HandleUpdates
    End If

    Unload frmDatabaseOptions
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub mnuToolsManagementOptions_Click()
    On Error GoTo Failed
    g_clsMainSupport.CloseFindDialogs
    frmManageUpdates.Show vbModal, Me

    If frmManageUpdates.GetReturnCode() = MSGSuccess Then
        'g_clsMainSupport.HandleUpdates
    End If

    Unload frmManageUpdates
Failed:
Exit Sub
g_clsErrorHandling.DisplayError
End Sub

Private Sub mnuToolsOptions_Click()
    On Error GoTo Failed
    g_clsMainSupport.CloseFindDialogs
    
    frmOptions.Show vbModal, Me
    
    If frmOptions.GetReturnCode() = MSGSuccess Then
        g_clsMainSupport.HandleUpdates
    End If
    Unload frmOptions

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub mnuView_Click(Index As Integer)
    On Error GoTo Failed
    PopulateListView
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub AddTreeNodes()
    On Error GoTo Failed
    Dim TreeItem As TreeAccess
    Dim nCount As Long
    Dim nRelationship As Integer
    Dim nParent As Integer
    Dim sName As String
    Dim sKey As String 'JR SYS4537
    
    tvwDB.Nodes.Clear
    Set m_nodeCurrent = Nothing
    
    For nCount = 1 To m_treeData.Count
        Set TreeItem = m_treeData.Item(nCount)
        nParent = TreeItem.GetParent()
        nRelationship = TreeItem.GetRelationship()
        sName = TreeItem.GetName()
        sKey = TreeItem.GetKey()  'JR SYS4537
        
        If (nRelationship = -1) Then
            AddNode sKey, sName ' JR SYS4537
        Else
            AddNode sKey, sName, nRelationship, nParent 'JR SYS4537, added sName
        End If
    Next nCount
    
    ' Expand the first node
    tvwDB.Nodes(1).Expanded = True
    Exit Sub
Failed:
    Err.Raise 1, , "Error adding tree nodes: " & Err.DESCRIPTION
End Sub
Public Sub InitListView()
    lvListView.View = lvwReport
End Sub
Private Sub AddNode(ByRef sTitle As String, _
    ByRef sName As String, Optional vRelationship As Variant, Optional vParent As Variant)
    On Error GoTo Failed
    
    Dim mNode As node
    Dim nCount As Integer
    
    nCount = tvwDB.Nodes.Count
    
    If (IsMissing(vRelationship) Or IsMissing(vParent)) Then
        Set mNode = tvwDB.Nodes.Add(, , sTitle, sName, "closed")
    Else
        Set mNode = tvwDB.Nodes.Add(vParent, vRelationship, sTitle, sName, "closed")
    End If
    Exit Sub
Failed:
    Err.Raise 1, , "Error adding node " & sTitle & ", " & sName & ": " & Err.DESCRIPTION
End Sub
Friend Sub SetColumnHeader(sTag As String)
    Dim nWidth As Long
    Dim sColName As String
    Dim columnWidths As Collection
    Dim columnHeadings As Collection
    Dim TreeItem As TreeAccess
    Dim nCount As Long
    Dim nTotal As Long
    Dim dOnePercent As Double
    
    dOnePercent = frmMain.lvListView.Width / 100
    
    lvListView.ColumnHeaders.Clear
    lvListView.HideColumnHeaders = False
    
    Set TreeItem = m_treeData(sTag)
    Set columnWidths = TreeItem.GetColumnWidths
    Set columnHeadings = TreeItem.GetColumnHeadings
    
    nTotal = columnHeadings.Count
    
    For nCount = 1 To nTotal
        sColName = columnHeadings(nCount)
        nWidth = columnWidths(nCount) * dOnePercent
        SetFieldHeader sColName, nWidth
    Next nCount
End Sub
Public Sub SetFieldHeader(sName As String, nWidth As Long)
    Dim colX As ColumnHeader ' Declare variable.
    'Dim intX As Integer ' Counter variable.
    
    Set colX = lvListView.ColumnHeaders.Add()
    colX.Text = sName
    colX.Width = nWidth
End Sub
' This is needed because of a bug in VB6.0 If you display a popupmenu on a modal dialog
' that's been opened from somewhere else, the popup won't be displayed. If you display the
' modal dialog from the timer, it works.
Private Sub Timer1_Timer()
    On Error GoTo Failed
    Select Case Timer1.Tag
    
    Case "TreeView"
        Dim vKey As Variant
        vKey = frmMain.GetSelectedTreeKey()
        
        g_clsMainSupport.HandleTreeNew vKey
    
    Case "ListViewEdit"
        g_clsMainSupport.ListEdit
    
    Case "ListViewDelete"
        g_clsMainSupport.DeleteRecord
    
    Case "ListViewPromote"
        g_clsMainSupport.PromoteRecord
    End Select

    Timer1.Enabled = False
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    Timer1.Enabled = False

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupTreeValues
' Description   :   Sets up the treeview with all required nodes that need
'                   to be displayed. Tree nodes and menu items are also
'                   made invisible if the user is not allowed to use them.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Change history
' Prog      Date        Description
' SDS       11-0202  SYS4030
'AQR SYS4030 was fixed following the assumption that items shown in the Swap List
'form the Lowermost level in Tree view.
' JR        27/05/02    SYS4537, added SKey param to TreeItem.AddName
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetupTreeValues()
    
    Dim bAllow As Boolean
    Dim TreeItem As TreeAccess
    Dim nAddressIndex As Integer
    Dim nPanelIndex As Integer
    Dim nOptionsIndex As Integer
    Dim nRatesAndFeesIndex As Integer
    Dim nGlobalParamIndex As Integer
    Dim nLockIndex As Integer
    Dim nLockTypeIndex As Integer
    Dim nTaskManagementIndex As Integer
    Dim nCaseTrackingIndex As Integer
    Dim nIntermediariesIndex As Integer
    Dim nIntroducersIndex As Integer ' PB 18/10/2006 EP2_13
    Dim nFirmsIndex As Integer ' PB 18/10/2006 EP2_13
    Dim nIndividualsIndex As Integer ' PB 18/10/2006 EP2_13
    Dim nProcurationFeesIndex As Integer ' TW 14/12/2006 EP2_518
    Dim nBatchProcessingIndex As Integer ' TW 27/11/2007 DBM594
    Dim bParentAllowed As Boolean
    Dim clsTreeViewCS As TreeViewCS
    
    On Error GoTo Failed
    'Initialise variable to false
    bParentAllowed = False
    
    'Before we start adding/securing tree items, secure promotions.
    mnuToolsManagementOptions.Visible = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PROMOTIONS)
                
'   'TODO: If this is required then remove the comments.
'   'Before we start adding/securing tree items, secure db connections.
'    mnuToolsDatabaseOptions.Visible = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_DB_CONNECTIONS)
        
    'Initialise the tree control.
    tvwDB.LabelEdit = 1
    tvwDB.LineStyle = tvwRootLines
    tvwDB.Sorted = False
    Set m_treeData = New Collection
    
    ' First the Root node
    Set TreeItem = New TreeAccess
    TreeItem.AddName "Options", "Options"
    m_treeData.Add TreeItem, "Options"
    nOptionsIndex = m_treeData.Count()
    
    'COMBOBOX ENTRIES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_COMBOBOX_ENTRIES)
    
    'MAR16 GHun
    ''Show the File->New menu AFTER the first access check has taken place.
    'mnuFile.Visible = True
    'MAR16 End
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewComboEntry.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName COMBOBOX_ENTRIES, COMBOBOX_ENTRIES, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Group Name", 30
        TreeItem.AddHeading "Notes", 50
        m_treeData.Add TreeItem, COMBOBOX_ENTRIES
    End If
        
    'GLOBAL PARAMETERS

    mnuFileNewGlobalParameter.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName SYSTEM_PARAMETERS, SYSTEM_PARAMETERS, tvwChild, nOptionsIndex
' TW 02/01/2007 EP2_640
'    TreeItem.AddHeading "Name", 15
'    TreeItem.AddHeading "Data Type", 15
'    TreeItem.AddHeading "Value", 15
'    TreeItem.AddHeading "Active From", 15
'    TreeItem.AddHeading "Active To", 15
' TW 02/01/2007 EP2_640 End
    m_treeData.Add TreeItem, SYSTEM_PARAMETERS
    nGlobalParamIndex = m_treeData.Count
        
    'Global Parameters -> Fixed
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_GLOBAL_PARAM_FIXED)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewGlobalParameterFixed.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName GLOBAL_PARAM_FIXED, GLOBAL_PARAM_FIXED, tvwChild, nGlobalParamIndex
        TreeItem.AddHeading "Name", 25
        TreeItem.AddHeading "Description", 40
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "Amount", 12
        TreeItem.AddHeading "Percentage", 13
        TreeItem.AddHeading "Maximum", 10
        TreeItem.AddHeading "Boolean", 10
        TreeItem.AddHeading "String", 15
        m_treeData.Add TreeItem, GLOBAL_PARAM_FIXED
        bParentAllowed = True
    End If
    
    'Global Parameters -> Banded
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_GLOBAL_PARAM_BANDED)
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName GLOBAL_PARAM_BANDED, GLOBAL_PARAM_BANDED, tvwChild, nGlobalParamIndex
        TreeItem.AddHeading "Name", 20
        TreeItem.AddHeading "Description", 25
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "Band", 10
        TreeItem.AddHeading "Amount", 15
        TreeItem.AddHeading "Percentage", 15
        TreeItem.AddHeading "Maximum", 15
        TreeItem.AddHeading "Boolean", 15
        TreeItem.AddHeading "String", 15
        m_treeData.Add TreeItem, GLOBAL_PARAM_BANDED
        'Show/Hide the corresponding 'New' menu item.
        mnuFileNewGlobalParameterBanded.Visible = bAllow
        mnuFileNewGlobalParameter.Visible = True
        bParentAllowed = True
    Else
          If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewGlobalParameter.Visible = True
            mnuFileNewGlobalParameterBanded.Visible = bAllow
          Else
            m_treeData.Remove SYSTEM_PARAMETERS
          End If
    End If
    
    'RATES AND FEES
    bParentAllowed = False
    
    mnuFileNewRates.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName RATES_AND_FEES, RATES_AND_FEES, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, RATES_AND_FEES
    nRatesAndFeesIndex = m_treeData.Count()
    
' TW 09/10/2006 EP2_7
    'ID_ADDITIONAL_BORROWING_FEES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_ADDITIONAL_BORROWING_FEES)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName ADDITIONAL_BORROWING_FEES, ADDITIONAL_BORROWING_FEES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Nature Of Loan", 15
        TreeItem.AddHeading "Max LTV", 15
        TreeItem.AddHeading "Fee Amount", 15
        TreeItem.AddHeading "Fee Percentage", 15
        TreeItem.AddHeading "Min Fee Value", 15
        TreeItem.AddHeading "Max Fee Value", 15
        m_treeData.Add TreeItem, ADDITIONAL_BORROWING_FEES
        bParentAllowed = True
    End If
    
    'ID_CREDIT_LIMIT_INCREASE_FEES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_CREDIT_LIMIT_INCREASE_FEES)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName CREDIT_LIMIT_INCREASE_FEES, CREDIT_LIMIT_INCREASE_FEES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Nature Of Loan", 15
        TreeItem.AddHeading "Max LTV", 15
        TreeItem.AddHeading "Fee Amount", 15
        TreeItem.AddHeading "Fee Percentage", 15
        TreeItem.AddHeading "Min Fee Value", 15
        TreeItem.AddHeading "Max Fee Value", 15
        m_treeData.Add TreeItem, CREDIT_LIMIT_INCREASE_FEES
        bParentAllowed = True
    End If
' TW 09/10/2006 EP2_7 End


    'ADMIN FEES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_ADMIN_FEES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewRatesAdminFee.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName ADMIN_FEES, ADMIN_FEES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Type of Application", 15
        TreeItem.AddHeading "Location", 15
        TreeItem.AddHeading "Fee Amount", 15
        m_treeData.Add TreeItem, ADMIN_FEES
        bParentAllowed = True
    End If
    
    'VALUATION FEES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_VALUATION_FEES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewRatesValuationFee.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName VALUATION_FEES, VALUATION_FEES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Type of Application", 15
        TreeItem.AddHeading "Location", 15
        TreeItem.AddHeading "Maximum Value", 15
        TreeItem.AddHeading "Valuation Fee", 15
        '   AW  15/05/02    BM023
        TreeItem.AddHeading "Fee Percentage", 15
        TreeItem.AddHeading "Minimum Fee", 15
        TreeItem.AddHeading "Maximum Fee Value", 15
        
' TW 09/10/2006 EP2_7
        TreeItem.AddHeading "Type of Valuation", 15
' TW 09/10/2006 EP2_7 End
        m_treeData.Add TreeItem, VALUATION_FEES
        bParentAllowed = True
    End If
        
    '*=[MC]BMIDS763 - CC075
    '
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PRODUCT_SWITCH_FEESETS)
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewRatesAdminFee.Visible = bAllow
    mnuFileNewRatesProdSwitchFee.Visible = bAllow
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRODUCT_SWITCH_FEESETS, PRODUCT_SWITCH_FEESETS, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Type of Application", 15
        TreeItem.AddHeading "Location", 15
        TreeItem.AddHeading "Fee Amount", 15
        m_treeData.Add TreeItem, PRODUCT_SWITCH_FEESETS
        bParentAllowed = True
    End If
        
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_INSURANCE_ADMIN_FEESETS)
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewRatesAdminFee.Visible = bAllow
    mnuFileNewRatesInsuAdminFeeSet.Visible = bAllow
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName INSURANCE_ADMIN_FEESETS, INSURANCE_ADMIN_FEESETS, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Type of Application", 15
        TreeItem.AddHeading "Location", 15
        TreeItem.AddHeading "Fee Amount", 15
        m_treeData.Add TreeItem, INSURANCE_ADMIN_FEESETS
        bParentAllowed = True
    End If
        
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TT_FEESETS)
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewRatesAdminFee.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName TT_FEESETS, TT_FEESETS, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Type of Application", 15
        TreeItem.AddHeading "Location", 15
        TreeItem.AddHeading "Fee Amount", 15
        m_treeData.Add TreeItem, TT_FEESETS
        bParentAllowed = True
    End If
    
    '*=[MC]SECTION END - BMIDS763 - CC075
    
    'BMIDS765 JD Add Rental Income Rate Sets
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_RENTALINCOMERATES)
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewRatesRentalIncomeRateSet.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName RENTAL_INCOME_RATES, RENTAL_INCOME_RATES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Max Loan Amount", 15
        TreeItem.AddHeading "Max LTV", 15
        TreeItem.AddHeading "Rental Income Interest Rate", 15
        m_treeData.Add TreeItem, RENTAL_INCOME_RATES
        bParentAllowed = True
    End If

    'BASERATE SETS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_BASE_RATES)
    
    'Show/Hide the corresponding 'New' menu item.
    ' SYS4085
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName BASE_RATES, BASE_RATES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 25
        TreeItem.AddHeading "Start Date", 25
        TreeItem.AddHeading "Max Loan Amount", 20
        TreeItem.AddHeading "Max LTV", 20
        TreeItem.AddHeading "Rate Diff", 15
        m_treeData.Add TreeItem, BASE_RATES
        mnuFileNewRatesBaseRate.Visible = bAllow
        mnuFileNewRates.Visible = True
        bParentAllowed = True
    End If

    'BASE RATES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_BASE_RATE)

    If bAllow And g_clsMainSupport.DoesRateExist Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName BASE_RATE, BASE_RATE, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Rate ID", 15
' TW 18/05/2007 VR262
        TreeItem.AddHeading "Description", 20
' TW 18/05/2007 VR262 End
        TreeItem.AddHeading "StartDate", 15
        TreeItem.AddHeading "Interest Rate", 15
        TreeItem.AddHeading "Type", 15
        TreeItem.AddHeading "Applied Date", 20
        m_treeData.Add TreeItem, BASE_RATE
        mnuFileNewRatesBaseRateSet.Visible = bAllow
        bParentAllowed = True
    Else
          If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewRates.Visible = True
            mnuFileNewRatesBaseRate.Visible = bAllow
          Else
            m_treeData.Remove RATES_AND_FEES
          End If
    End If
 
    'AW 13/05/02    BM088
    'INCOME MULTIPLES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_INCOME_MULTIPLE)
    
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewIncomeMultiple.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName INCOME_MULTIPLE, INCOME_MULTIPLE, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Income Multiplier Code ", 20
        TreeItem.AddHeading "Description", 20
        TreeItem.AddHeading "Single Inc Mult", 15
        TreeItem.AddHeading "Joint Inc Mult", 15
        TreeItem.AddHeading "Highest Earner Inc Mult", 15
        TreeItem.AddHeading "Lowest Earner Inc Mult", 1
        m_treeData.Add TreeItem, INCOME_MULTIPLE
        bParentAllowed = True
    End If
        
    'AW 13/05/02    BM087
    'MPMIGRATE SETS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_MP_MIG_RATE_SETS)
    
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewMPMigRate.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName MP_MIG_RATE_SETS, MP_MIG_RATE_SETS, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "HLC Rate Set Code", 25
        TreeItem.AddHeading "Description", 25
        m_treeData.Add TreeItem, MP_MIG_RATE_SETS
        'mnuFileNewRatesBaseRate.Visible = bAllow
        'mnuFileNewRates.Visible = True
        bParentAllowed = True
    End If
    
    'AW 21/05/02    BM017
    'REDEMSETS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_REDEM_FEE_SETS)
    
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewRedemptionFee.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName REDEM_FEE_SETS, REDEM_FEE_SETS, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 25
        TreeItem.AddHeading "Description", 25
        m_treeData.Add TreeItem, REDEM_FEE_SETS
        'mnuFileNewRatesBaseRate.Visible = bAllow
        'mnuFileNewRates.Visible = True
        bParentAllowed = True
    End If
' TW 11/12/2006 EP2_20
    ' Transfer of Equity Fees
    
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TRANSFER_OF_EQUITY_FEES)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName TRANSFER_OF_EQUITY_FEES, TRANSFER_OF_EQUITY_FEES, tvwChild, nRatesAndFeesIndex
        TreeItem.AddHeading "Set Number", 20
        TreeItem.AddHeading "Start Date", 20
        TreeItem.AddHeading "Nature Of Loan", 15
        TreeItem.AddHeading "Fee Amount", 15
        m_treeData.Add TreeItem, TRANSFER_OF_EQUITY_FEES
        bParentAllowed = True
    End If
' TW 11/12/2006 EP2_20 End
    
    'BATCH SCHEDULER
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_BATCH_SCHEDULER) And g_clsMainSupport.DoesBatchExist()
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewBatchScheduler.Visible = bAllow
' TW 27/11/2007 DBM594
    Set TreeItem = New TreeAccess
    TreeItem.AddName BATCH_PROCESSING, BATCH_PROCESSING, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, BATCH_PROCESSING
    nBatchProcessingIndex = m_treeData.Count() ' TW 27/11/2007 DBM594
' TW 27/11/2007 DBM594 End
'PAYMENTS_FOR_COMPLETION
    
    If bAllow Then
        Set TreeItem = New TreeAccess
' TW 27/11/2007 DBM594
'        TreeItem.AddName BATCH_SCHEDULER, BATCH_SCHEDULER, tvwChild, nOptionsIndex
        TreeItem.AddName BATCH_SCHEDULER, BATCH_SCHEDULER, tvwChild, nBatchProcessingIndex
' TW 27/11/2007 DBM594 End
        TreeItem.AddHeading "Batch Number", 15
        TreeItem.AddHeading "Program Type", 20
        TreeItem.AddHeading "User ID", 15
        TreeItem.AddHeading "Batch Status", 20
        TreeItem.AddHeading "Batch Frequency", 15
        TreeItem.AddHeading "Batch Description", 30
        TreeItem.AddHeading "Execution Date/Time", 20
        m_treeData.Add TreeItem, BATCH_SCHEDULER
    End If
        
' TW 27/11/2007 DBM594
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PAYMENTS_FOR_COMPLETION, PAYMENTS_FOR_COMPLETION, tvwChild, nBatchProcessingIndex
       
        TreeItem.AddHeading "Application Number", 15
        TreeItem.AddHeading "Amount", 15
        TreeItem.AddHeading "Payment No.", 15
        TreeItem.AddHeading "Payment Type", 15
        TreeItem.AddHeading "Status", 15
        TreeItem.AddHeading "Application Type", 15
        TreeItem.AddHeading "Excluded?", 5
        TreeItem.AddHeading "Manual Completion?", 5

        m_treeData.Add TreeItem, PAYMENTS_FOR_COMPLETION
    End If
' TW 27/11/2007 DBM594 End
        
    'LENDERS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LENDERS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewLender.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LENDERS, LENDERS, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Lender Code", 20
        TreeItem.AddHeading "Lender Name", 20
        TreeItem.AddHeading "Lender Start Date", 15
        TreeItem.AddHeading "Lender End Date", 15
        m_treeData.Add TreeItem, LENDERS
    End If
            
    'PRODUCTS
    bParentAllowed = False

    mnuFileNewProducts.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName PRODUCTS, PRODUCTS, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, PRODUCTS
    nAddressIndex = m_treeData.Count

    'MORTGAGE PRODUCTS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_MORTGAGE_PRODUCTS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewProductsProduct.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName MORTGAGE_PRODUCTS, MORTGAGE_PRODUCTS, tvwChild, nAddressIndex
        'BS BM0311 26/03/03
        TreeItem.AddHeading "Product Code", 20
        TreeItem.AddHeading "Start Date", 10
        TreeItem.AddHeading "End Date", 10
        'BS BM0311 26/03/03
        'TreeItem.AddHeading "Product Code", 20
        TreeItem.AddHeading "Product Name", 30
        TreeItem.AddHeading "Lender Code", 20
        
        If Not g_clsVersion.DoesVersioningExist() Then
            TreeItem.AddHeading "Valid?", 20
        End If
        
        m_treeData.Add TreeItem, MORTGAGE_PRODUCTS
        bParentAllowed = True
    End If

    'LIFE COVER RATES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LIFE_COVER_RATES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewProductsLifeCoverRates.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LIFE_COVER_RATES, LIFE_COVER_RATES, tvwChild, nAddressIndex
        TreeItem.AddHeading "Life Cover Number", 15
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "Cover Type", 20
        TreeItem.AddHeading "Applicant Gender", 10
        TreeItem.AddHeading "Max Age", 10
        TreeItem.AddHeading "Max Term", 10
        TreeItem.AddHeading "Annual Rate", 10
        TreeItem.AddHeading "Smoker Rate", 10
        TreeItem.AddHeading "Poor Health Rate", 10
        m_treeData.Add TreeItem, LIFE_COVER_RATES
        bParentAllowed = True
    End If

    'BUILDINGS AND CONTENTS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_BUILDINGS_AND_CONTENTS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewProductsBAndC.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName BUILDINGS_AND_CONTENTS_PRODUCTS, BUILDINGS_AND_CONTENTS_PRODUCTS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Product Name", 25
        TreeItem.AddHeading "Product Number", 15
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "End Date", 15
        TreeItem.AddHeading "Valuables Limit", 10
        m_treeData.Add TreeItem, BUILDINGS_AND_CONTENTS_PRODUCTS
        bParentAllowed = True
    End If
    
    'PAYMENT PROTECTION
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PAYMENT_PROTECTION_RATES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewProductsPayProtRates.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PAYMENT_PROTECTION_RATES, PAYMENT_PROTECTION_RATES, tvwChild, nAddressIndex
        TreeItem.AddHeading "PP Rate Number", 15
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "Channel", 15
        TreeItem.AddHeading "Applicant Gender", 10
        TreeItem.AddHeading "Max Applicant Age", 15
        TreeItem.AddHeading "ASU Rate", 15
        TreeItem.AddHeading "U Rate", 15
        m_treeData.Add TreeItem, PAYMENT_PROTECTION_RATES
        bParentAllowed = True
    End If

    'PAYMENT PROTECTION PRODUCTS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PAYMENT_PROTECTION_PRODUCTS)
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PAYMENT_PROTECTION_PRODUCTS, PAYMENT_PROTECTION_PRODUCTS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Product Code", 15
        TreeItem.AddHeading "Start Date", 15
        TreeItem.AddHeading "End Date", 15
        TreeItem.AddHeading "Product Name", 20
        TreeItem.AddHeading "Max Applicant Age", 15
        TreeItem.AddHeading "Rate Set", 15
        m_treeData.Add TreeItem, PAYMENT_PROTECTION_PRODUCTS
        'Show/Hide the corresponding 'New' menu item.
        mnuFileNewProductsPayProtProducts.Visible = bAllow
        mnuFileNewProducts.Visible = True
        bParentAllowed = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewProducts.Visible = True
            mnuFileNewProductsPayProtProducts.Visible = bAllow
        Else
            m_treeData.Remove PRODUCTS
        End If
    End If

    'ORGANISATION
    bParentAllowed = False

    mnuFileNewOrg.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName ORGANISATION, ORGANISATION, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, ORGANISATION
    nAddressIndex = m_treeData.Count

    'COUNTRY ENTRIES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_COUNTRIES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgCountry.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName COUNTRIES, COUNTRIES, tvwChild, nAddressIndex
        TreeItem.AddHeading "Country Number", 30
        TreeItem.AddHeading "Country Name", 50
        m_treeData.Add TreeItem, COUNTRIES
        bParentAllowed = True
    End If
    
    'DISTRIBUTION CHANNELS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_DISTRIBUTION_CHANNELS)
        
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgDistChannel.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName DISTRIBUTION_CHANNELS, DISTRIBUTION_CHANNELS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Channel Id", 25
        TreeItem.AddHeading "Channel Name", 75
        m_treeData.Add TreeItem, DISTRIBUTION_CHANNELS
        bParentAllowed = True
    End If
    
    'DEPARTMENTS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_DEPARTMENTS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgDepartment.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName DEPARTMENTS, DEPARTMENTS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Department Id", 15
        TreeItem.AddHeading "Department Name", 30
        TreeItem.AddHeading "Channel", 15
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        m_treeData.Add TreeItem, DEPARTMENTS
        bParentAllowed = True
    End If
    
    'REGIONS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_REGIONS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgRegion.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName REGIONS, REGIONS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Region Id", 15
        TreeItem.AddHeading "Region Type", 30
        TreeItem.AddHeading "Region Name", 30
        m_treeData.Add TreeItem, REGIONS
        bParentAllowed = True
    End If
    
    'UNITS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_UNITS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgUnit.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName UNITS, UNITS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Unit Id", 15
        TreeItem.AddHeading "Unit Name", 30
        TreeItem.AddHeading "Department Id", 30
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        m_treeData.Add TreeItem, UNITS
        bParentAllowed = True
    End If
    
    'COMPETENCIES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_COMPETENCIES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgCompetencies.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName COMPETENCIES, COMPETENCIES, tvwChild, nAddressIndex
        TreeItem.AddHeading "Competency Type", 15
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        TreeItem.AddHeading "Funds Release Mandate", 15
        TreeItem.AddHeading "Loan Amount Mandate", 15
        TreeItem.AddHeading "LTV Mandate", 15
        TreeItem.AddHeading "Risk Assessment Mandate", 15
        m_treeData.Add TreeItem, COMPETENCIES
        bParentAllowed = True
    End If
    
    'WORKING HOURS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_WORKING_HOURS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOrgWorkingHours.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName WORKING_HOURS, WORKING_HOURS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Working Hours Type", 15
        TreeItem.AddHeading "Working Hours Type Description", 30
        TreeItem.AddHeading "Bank Holiday Indicator", 15
        m_treeData.Add TreeItem, WORKING_HOURS
        bParentAllowed = True
    End If
    
    'USERS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_USERS)
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName USERS, USERS, tvwChild, nAddressIndex
        TreeItem.AddHeading "User ID", 15
        TreeItem.AddHeading "User Name", 25
        TreeItem.AddHeading "Access Type", 20
        TreeItem.AddHeading "Active From", 15
        TreeItem.AddHeading "Active To", 15
        m_treeData.Add TreeItem, USERS
        'Show/Hide the corresponding 'New' menu item.
        mnuFileNewOrgUser.Visible = bAllow
        mnuFileNewOrg.Visible = True
        bParentAllowed = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewOrg.Visible = True
            mnuFileNewOrgUser.Visible = bAllow
        Else
            m_treeData.Remove ORGANISATION
        End If
    End If
    
    'THIRD PARTIES
    bParentAllowed = False
    
    mnuFileNewTP.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName NAMES_AND_ADDRESSES, NAMES_AND_ADDRESSES, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, NAMES_AND_ADDRESSES
    nAddressIndex = m_treeData.Count
    
    'PANEL

    mnuFileNewTPPanel.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName PANEL_ADDRESS, PANEL_ADDRESS, tvwChild, nAddressIndex
    m_treeData.Add TreeItem, PANEL_ADDRESS
    nPanelIndex = m_treeData.Count

    'LEGAL REP
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LEGAL_REP_ADDRESS)
            
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewTPLegalRep.Visible = bAllow
                
    If bAllow Then
        'DB BMIDS00720 - company name should be first
        Set TreeItem = New TreeAccess
        TreeItem.AddName LEGAL_REP_ADDRESS, LEGAL_REP_ADDRESS, tvwChild, nPanelIndex
        TreeItem.AddHeading "Company Name", 30
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        TreeItem.AddHeading "Name and Address Type", 20
        TreeItem.AddHeading "Panel ID", 15
        TreeItem.AddHeading "Legal Rep Status", 10 ' TK 29/11/2005 MAR81
        TreeItem.AddHeading "Panel Userid", 0
        m_treeData.Add TreeItem, LEGAL_REP_ADDRESS
        mnuFileNewTPPanel.Visible = True
        bParentAllowed = True
    End If

    'VALUER
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_VALUER_ADDRESS)
    
    If bAllow Then
        'DB BMIDS00720 - company name should be first
        Set TreeItem = New TreeAccess
        TreeItem.AddName VALUER_ADDRESS, VALUER_ADDRESS, tvwChild, nPanelIndex
        TreeItem.AddHeading "Company Name", 30
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        TreeItem.AddHeading "Name and Address Type", 20
        TreeItem.AddHeading "Panel ID", 15
        TreeItem.AddHeading "", 0 ' JD MAR990 to match legal rep and eliminate an error
        TreeItem.AddHeading "Panel Userid", 0 ' JD MAR990
        m_treeData.Add TreeItem, VALUER_ADDRESS
        'Show/Hide the corresponding 'New' menu item.
        mnuFileNewTPValuer.Visible = bAllow
        mnuFileNewTPPanel.Visible = True
        mnuFileNewTP.Visible = True
        bParentAllowed = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
             mnuFileNewTPValuer.Visible = bAllow
            'mnuFileNewTPPanel.Visible = bAllow
            mnuFileNewTP.Visible = True
        
        Else
            'mnuFileNewTP.Visible = False
            'mnuFileNewTPPanel.Visible = False
            m_treeData.Remove PANEL_ADDRESS
        End If
    End If
 
    'LOCAL
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCAL_ADDRESS)
    
    'Show/Hide the corresponding 'New' menu item.
    
    If bAllow Then
        'DB BMIDS00720 - company name should be first
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCAL_ADDRESS, LOCAL_ADDRESS, tvwChild, nAddressIndex
        TreeItem.AddHeading "Company Name", 30
        TreeItem.AddHeading "Active From", 10
        TreeItem.AddHeading "Active To", 10
        TreeItem.AddHeading "Name and Address Type", 25
        
        ' DJP BM0318 Add Town
        TreeItem.AddHeading "Town ", 25
        mnuFileNewTPOther.Visible = bAllow
        m_treeData.Add TreeItem, LOCAL_ADDRESS
        mnuFileNewTP.Visible = True
        bParentAllowed = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewTP.Visible = True
            mnuFileNewTPPanel.Visible = True
            mnuFileNewTPOther.Visible = bAllow
        Else
            'mnuFileNewTP.Visible = False
            'mnuFileNewTPPanel.Visible = False
            m_treeData.Remove NAMES_AND_ADDRESSES
        End If
    End If
        
    'INCOME FACTORS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_INCOME_FACTORS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewIncomeFactor.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName INCOME_FACTORS, INCOME_FACTORS, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Lender Name", 15
        TreeItem.AddHeading "Employment Status", 15
        TreeItem.AddHeading "Income Group", 15
        TreeItem.AddHeading "Income Type", 15
        TreeItem.AddHeading "Factor", 15
        m_treeData.Add TreeItem, INCOME_FACTORS
    End If
    
    'ERROR MESSAGES
    
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_ERROR_MESSAGES)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewErrorMessage.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName ERROR_MESSAGES, ERROR_MESSAGES, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Error Code ", 15
        TreeItem.AddHeading "Message Type", 15
        TreeItem.AddHeading "Message Text", 60
        m_treeData.Add TreeItem, ERROR_MESSAGES
    End If
    
    'LOG LOCKS
    bParentAllowed = False
        
    FileNewAppLock.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCK_MAINTENANCE, LOCK_MAINTENANCE, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, LOCK_MAINTENANCE
    nLockIndex = m_treeData.Count()

    'LOCKS ONLINE
 
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_ONLINE, LOCKS_ONLINE, tvwChild, nLockIndex
    m_treeData.Add TreeItem, LOCKS_ONLINE
    nLockTypeIndex = m_treeData.Count()

    'ONLINE APPLICATION
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCKS_ONLINE_APPLICATION)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCKS_ONLINE_APPLICATION, LOCKS_ONLINE_APPLICATION, tvwChild, nLockTypeIndex
        TreeItem.AddHeading "Application Number", 20
        TreeItem.AddHeading "Lock Date", 15
        TreeItem.AddHeading "User ID", 10
        TreeItem.AddHeading "Unit ID", 10
        TreeItem.AddHeading "Machine ID", 10
        m_treeData.Add TreeItem, LOCKS_ONLINE_APPLICATION
        'Show/Hide the corresponding 'New' menu item.
        mnuFileNewOnlineAppLock.Visible = bAllow
        bParentAllowed = True
    End If

    'ONLINE CUSTOMER
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCKS_ONLINE_CUSTOMER)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCKS_ONLINE_CUSTOMER, LOCKS_ONLINE_CUSTOMER, tvwChild, nLockTypeIndex
        TreeItem.AddHeading "Customer Number", 20
        TreeItem.AddHeading "Lock Date", 15
        TreeItem.AddHeading "User ID", 10
        TreeItem.AddHeading "Unit ID", 10
        TreeItem.AddHeading "Machine ID", 10
        m_treeData.Add TreeItem, LOCKS_ONLINE_CUSTOMER
        FileNewAppLock.Visible = True
        bParentAllowed = True
    Else
      If bParentAllowed Then
        ' Only make the last child invisible, if a prior child has made the parent visible
        ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
        FileNewAppLock.Visible = True
        FileNewAppLock.Visible = bAllow
      Else
        m_treeData.Remove LOCKS_ONLINE
      End If
    End If

    'LOCKS OFFLINE
    Set TreeItem = New TreeAccess
    TreeItem.AddName LOCKS_OFFLINE, LOCKS_OFFLINE, tvwChild, nLockIndex
    m_treeData.Add TreeItem, LOCKS_OFFLINE
    nLockTypeIndex = m_treeData.Count()

    'OFFLINE APPLICATION
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCKS_OFFLINE_APPLICATION)
            
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewOfflineAppLock.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCKS_OFFLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION, tvwChild, nLockTypeIndex
        TreeItem.AddHeading "Application Number", 20
        TreeItem.AddHeading "Lock Date", 15
        TreeItem.AddHeading "User ID", 10
        TreeItem.AddHeading "Unit ID", 10
        TreeItem.AddHeading "Machine ID", 10
        m_treeData.Add TreeItem, LOCKS_OFFLINE_APPLICATION
    End If

    'ONLINE CUSTOMER
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCKS_OFFLINE_CUSTOMER)
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCKS_OFFLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER, tvwChild, nLockTypeIndex
        TreeItem.AddHeading "Customer Number", 20
        TreeItem.AddHeading "Lock Date", 15
        TreeItem.AddHeading "User ID", 10
        TreeItem.AddHeading "Unit ID", 10
        TreeItem.AddHeading "Machine ID", 10
        m_treeData.Add TreeItem, LOCKS_OFFLINE_CUSTOMER
        FileNewAppLock.Visible = True
        bParentAllowed = True
    Else
      If bParentAllowed Then
        ' Only make the last child invisible, if a prior child has made the parent visible
        ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
        FileNewAppLock.Visible = True
        'FileNewAppLock.Visible = True
      Else
        m_treeData.Remove LOCKS_OFFLINE
        m_treeData.Remove LOCK_MAINTENANCE
      End If
    End If
 
    ' ik_bm0314

    'LOCKS DOCUMENTS
    nLockTypeIndex = m_treeData.Count()
            
    ' ik_frig
    ' bAllow = True
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_LOCKS_DOCS)
            
    'Show/Hide the corresponding 'New' menu item.
    ' mnuFileNewOfflineAppLock.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOCKS_DOCS, LOCKS_DOCS, tvwChild, nLockIndex
        TreeItem.AddHeading "Application Number", 15
        TreeItem.AddHeading "Lock Date", 12
        ' ik_wip_20030310
        ' TreeItem.AddHeading "User ID", 8
        ' TreeItem.AddHeading "Unit ID", 8
        TreeItem.AddHeading "Locked By", 8
        ' ik_wip_20030310_ends
        TreeItem.AddHeading "Document", 32
        TreeItem.AddHeading "Version", 10
        m_treeData.Add TreeItem, LOCKS_DOCS
    
        nLockTypeIndex = m_treeData.Count()
    End If
 
    ' ik_bm0314_ends
        
    'TASK MANAGEMENT
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TASK_MANAGEMENT) And _
        g_clsMainSupport.DoesTaskManagementExist()
    
 '   If bAllow Then
    mnuFileNewTaskManage.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName TASK_MANAGEMENT, TASK_MANAGEMENT, tvwChild, nOptionsIndex
    m_treeData.Add TreeItem, TASK_MANAGEMENT
    nTaskManagementIndex = m_treeData.Count()

    'TASKS
    bParentAllowed = False
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TASK_MANAGEMENT_TASKS)
    mnuFileNewTaskManageTask.Visible = bAllow
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_TASKS, TASK_MANAGEMENT_TASKS, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Task ID", 15
        TreeItem.AddHeading "Task", 40
        TreeItem.AddHeading "Type", 20
        TreeItem.AddHeading "Owner Type", 20
        m_treeData.Add TreeItem, TASK_MANAGEMENT_TASKS
        bParentAllowed = True
    End If
        
    'STAGES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TASK_MANAGEMENT_STAGES)
    mnuFileNewTaskManageStage.Visible = bAllow
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_STAGES, TASK_MANAGEMENT_STAGES, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Stage ID", 15
        TreeItem.AddHeading "Stage", 40
        TreeItem.AddHeading "Authority Level", 20
        TreeItem.AddHeading "Exception?", 20
        m_treeData.Add TreeItem, TASK_MANAGEMENT_STAGES
        bParentAllowed = True
    End If
    
    'ACTIVITIES
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_TASK_MANAGEMENT_ACTIVITIES)
    mnuFileNewTaskManageActivity.Visible = bAllow
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName TASK_MANAGEMENT_ACTIVITIES, TASK_MANAGEMENT_ACTIVITIES, tvwChild, nTaskManagementIndex
        TreeItem.AddHeading "Activity ID", 15
        TreeItem.AddHeading "Activity", 40
        TreeItem.AddHeading "Desciption", 40
        m_treeData.Add TreeItem, TASK_MANAGEMENT_ACTIVITIES
        bParentAllowed = True
    End If
                            
    'CASE TRACKING
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_CASE_TRACKING)
            
    'If bAllow Then
    Set TreeItem = New TreeAccess
    TreeItem.AddName CASE_TRACKING, CASE_TRACKING, tvwChild, nTaskManagementIndex
    m_treeData.Add TreeItem, CASE_TRACKING
    nCaseTrackingIndex = m_treeData.Count()
            
    'BUSINESS GROUPS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_BUSINESS_GROUPS) And g_clsMainSupport.DoesBusinessGroupsExist()
            
    'Show/Hide the corresponding 'New' menu item.
    'mnuFileNewBusinessGroup.Visible = bAllow
            
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName BUSINESS_GROUPS, BUSINESS_GROUPS, tvwChild, nCaseTrackingIndex
        TreeItem.AddHeading "Group Name", 25
        TreeItem.AddHeading "Business Area", 15
        m_treeData.Add TreeItem, BUSINESS_GROUPS
        mnuFileNewTaskManage.Visible = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewTaskManage.Visible = True
            'mnuFileNewRatesBaseRate.Visible = bAllow
            mnuFileNewTaskManageCase.Visible = bAllow
            m_treeData.Remove CASE_TRACKING
        Else
            m_treeData.Remove CASE_TRACKING
            m_treeData.Remove TASK_MANAGEMENT
        End If
    End If
    'End If
'    Else
'        'APPLICATION PROCESSING
'        bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_APPLICATION_PROCESSING)
'
'        'Show/Hide the corresponding 'New' menu item.
'        mnuFileNewBusinessGroup.Visible = bAllow
'
'        If bAllow Then
'            Set TreeItem = New TreeAccess
'            TreeItem.AddName APPLICATION_PROCESSING, tvwChild, nOptionsIndex
'            TreeItem.AddHeading "Application Number", 20
'            TreeItem.AddHeading "Surname", 20
'            TreeItem.AddHeading "ForeName", 15
'            TreeItem.AddHeading "Address", 45
'            m_treeData.Add TreeItem, APPLICATION_PROCESSING
'        End If
'   End If

' PB 18/10/2006 EP2_13 - New organisation features
    'INTRODUCERS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_INTRODUCERS)
    
    If bAllow Then
        
      ' PB 03/11/2006 EP2_13 Revised design
        
        ' Introducers
        Set TreeItem = New TreeAccess
        TreeItem.AddName INTRODUCERS, INTRODUCERS, tvwChild, nOptionsIndex
' TW 02/01/2007 EP2_640
'        TreeItem.AddHeading " ", 20 ' blank
' TW 02/01/2007 EP2_640 End
        m_treeData.Add TreeItem, INTRODUCERS
        nIntroducersIndex = m_treeData.Count()
        
        Set TreeItem = New TreeAccess
        TreeItem.AddName ASSOCIATIONS, ASSOCIATIONS, tvwChild, nIntroducersIndex
        TreeItem.AddHeading "id", 0
        TreeItem.AddHeading "Name", 30
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "Description", 90
        TreeItem.AddHeading "Address", 50
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, ASSOCIATIONS
       
        Set TreeItem = New TreeAccess
        TreeItem.AddName CLUBS, CLUBS, tvwChild, nIntroducersIndex
        TreeItem.AddHeading "id", 0
        TreeItem.AddHeading "Name", 30
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "Description", 90
        TreeItem.AddHeading "Address", 50
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, CLUBS
        
        ' Firms
        Set TreeItem = New TreeAccess
        TreeItem.AddName FIRMS, FIRMS, tvwChild, nIntroducersIndex
' TW 02/01/2007 EP2_640
'        TreeItem.AddHeading " ", 20 ' blank
' TW 02/01/2007 EP2_640 End
        m_treeData.Add TreeItem, FIRMS
        nFirmsIndex = m_treeData.Count()
        
        ' Principals
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRINCIPALS, PRINCIPALS, tvwChild, nFirmsIndex
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "id", 5
'        TreeItem.AddHeading "Name", 20
'        TreeItem.AddHeading "Description", 90
        TreeItem.AddHeading "id", 0
        TreeItem.AddHeading "Unit Id", 5
        TreeItem.AddHeading "FSA Ref", 5
        TreeItem.AddHeading "Firm Name", 30
        TreeItem.AddHeading "Address", 50
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, PRINCIPALS
        
        ' Packagers
        Set TreeItem = New TreeAccess
        TreeItem.AddName PACKAGERS, PACKAGERS, tvwChild, nFirmsIndex
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "id", 5
'        TreeItem.AddHeading "Name", 20
'        TreeItem.AddHeading "Description", 90
'        TreeItem.AddHeading "id", 5
        TreeItem.AddHeading "id", 0
        TreeItem.AddHeading "Unit Id", 5
' TW 01/02/2007 EP2_1036
'        TreeItem.AddHeading "FSA Ref", 5
' TW 01/02/2007 EP2_1036 End
        TreeItem.AddHeading "Firm Name", 30
        TreeItem.AddHeading "Address", 50
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, PACKAGERS
        
        ' ARs
        Set TreeItem = New TreeAccess
        TreeItem.AddName ARFIRMS, ARFIRMS, tvwChild, nFirmsIndex
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "id", 5
'        TreeItem.AddHeading "Name", 20
'        TreeItem.AddHeading "Description", 90
        TreeItem.AddHeading "id", 0
        TreeItem.AddHeading "Unit Id", 5
        TreeItem.AddHeading "FSA Ref", 5
        TreeItem.AddHeading "Firm Name", 30
        TreeItem.AddHeading "Address", 50
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, ARFIRMS
        
        ' Individuals
        Set TreeItem = New TreeAccess
        TreeItem.AddName INDIVIDUALS, INDIVIDUALS, tvwChild, nIntroducersIndex
' TW 02/01/2007 EP2_640
'        TreeItem.AddHeading " ", 20 ' blank
' TW 02/01/2007 EP2_640 End
        m_treeData.Add TreeItem, INDIVIDUALS
        nFirmsIndex = m_treeData.Count()
        
        ' Packagers
        Set TreeItem = New TreeAccess
        TreeItem.AddName PACKAGERS, INDIVIDUALS & "_" & PACKAGERS, tvwChild, nFirmsIndex
        TreeItem.AddHeading "id", 5
        TreeItem.AddHeading "Name", 30
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "Description", 10
        TreeItem.AddHeading "Address", 50
'        TreeItem.AddHeading "Building Name/No", 20
'        TreeItem.AddHeading "Street", 25
'        TreeItem.AddHeading "Postcode", 10
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, INDIVIDUALS & "_" & PACKAGERS
        
        ' AR Brokers
        Set TreeItem = New TreeAccess
        TreeItem.AddName ARBROKERS, ARBROKERS, tvwChild, nFirmsIndex
        TreeItem.AddHeading "id", 5
' TW 19/02/2007 EP2_1036(2)
'' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "FSA Ref", 5
'' TW 16/01/2007 EP2_859 End
' TW 19/02/2007 EP2_1036(2) End
        TreeItem.AddHeading "Name", 30
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "Description", 10
        TreeItem.AddHeading "Address", 50
'        TreeItem.AddHeading "Building Name/No", 20
'        TreeItem.AddHeading "Street", 25
'        TreeItem.AddHeading "Postcode", 10
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, ARBROKERS
        
        ' DA Brokers
        Set TreeItem = New TreeAccess
        TreeItem.AddName DABROKERS, DABROKERS, tvwChild, nFirmsIndex
        TreeItem.AddHeading "id", 5
' TW 19/02/2007 EP2_1036(2)
'' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "FSA Ref", 5
'' TW 16/01/2007 EP2_859 End
' TW 19/02/2007 EP2_1036(2) End
        TreeItem.AddHeading "Name", 30
' TW 16/01/2007 EP2_859
'        TreeItem.AddHeading "Description", 10
        TreeItem.AddHeading "Address", 50
'        TreeItem.AddHeading "Building Name/No", 20
'        TreeItem.AddHeading "Street", 25
'        TreeItem.AddHeading "Postcode", 10
' TW 16/01/2007 EP2_859 End
        m_treeData.Add TreeItem, DABROKERS
        
        
' TW 14/12/2006 EP2_518
        ' Procuration Fees
        Set TreeItem = New TreeAccess
        TreeItem.AddName PROCURATION_FEES, PROCURATION_FEES, tvwChild, nIntroducersIndex
' TW 02/01/2007 EP2_640
'        TreeItem.AddHeading " ", 20 ' blank
' TW 02/01/2007 EP2_640 End
        m_treeData.Add TreeItem, PROCURATION_FEES
        nProcurationFeesIndex = m_treeData.Count()

        ' Default Procuration Fees
        Set TreeItem = New TreeAccess
        TreeItem.AddName DEFAULT_PROCURATION_FEES, DEFAULT_PROCURATION_FEES, tvwChild, nProcurationFeesIndex
        TreeItem.AddHeading "Submission Route", 30
        TreeItem.AddHeading "Product Scheme", 30
        TreeItem.AddHeading "Percentage", 10
        m_treeData.Add TreeItem, DEFAULT_PROCURATION_FEES

        ' Loan Amount Adjustments
        Set TreeItem = New TreeAccess
        TreeItem.AddName LOAN_AMOUNT_ADJUSTMENTS, LOAN_AMOUNT_ADJUSTMENTS, tvwChild, nProcurationFeesIndex
        TreeItem.AddHeading "Product Category", 30
        TreeItem.AddHeading "Maximum Loan", 10
        TreeItem.AddHeading "Percentage", 10
        m_treeData.Add TreeItem, LOAN_AMOUNT_ADJUSTMENTS

        ' LTV Amount Adjustments
        Set TreeItem = New TreeAccess
        TreeItem.AddName LTV_AMOUNT_ADJUSTMENTS, LTV_AMOUNT_ADJUSTMENTS, tvwChild, nProcurationFeesIndex
        TreeItem.AddHeading "Product Category", 30
        TreeItem.AddHeading "Maximum LTV", 10
        TreeItem.AddHeading "Percentage", 10
        m_treeData.Add TreeItem, LTV_AMOUNT_ADJUSTMENTS
' TW 14/12/2006 EP2_518 End

      
    End If
    ' PB 18/10/2006 EP2_13 End

    'QUESTIONS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_ADDITIONAL_QUESTIONS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewQuestion.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName ADDITIONAL_QUESTIONS, ADDITIONAL_QUESTIONS, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Reference", 30
        TreeItem.AddHeading "Question", 50
        TreeItem.AddHeading "Type", 30
        TreeItem.AddHeading "Details Required", 50
        TreeItem.AddHeading "Deleted", 30
        m_treeData.Add TreeItem, ADDITIONAL_QUESTIONS
    End If
    
    'CONDITIONS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_CONDITIONS)
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewCondition.Visible = bAllow
        
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName CONDITIONS, CONDITIONS, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Reference", 30
        TreeItem.AddHeading "Name", 50
        TreeItem.AddHeading "Description", 50
        TreeItem.AddHeading "Type", 30
        TreeItem.AddHeading "Editable", 20
        TreeItem.AddHeading "Free Format", 20
        TreeItem.AddHeading "Rule Ref", 20
        TreeItem.AddHeading "Deleted", 20
        m_treeData.Add TreeItem, CONDITIONS
    End If
    
    ' RF MAR202
    SetUpTreeValuesForPrinting nOptionsIndex
    bParentAllowed = False
    
    'CURRENCY CALCULATOR
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_CURRENCIES) And g_clsMainSupport.DoesCurrencyExist()
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewCurrencies.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName CURRENCIES, CURRENCIES, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Currency Code", 10
        TreeItem.AddHeading "Currency Name", 10
        TreeItem.AddHeading "Conversion Rate", 10
        TreeItem.AddHeading "Precision Rounding", 10
        TreeItem.AddHeading "Precision Direction", 10
        TreeItem.AddHeading "Base Currency", 10
        m_treeData.Add TreeItem, CURRENCIES
    End If
    
    'INTERMEDIARY
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_INTERMEDIARIES) And g_clsMainSupport.DoesIntermediaryExist()
' TW 20/12/2006 EP2_615
    If bAllow Then
' Show/Hide "Intermediaries" Option on Navigation Tree depending on value of 'SupervisorShowIntermediaries' Global Parameter
' True = reveal; False = hide
        On Error Resume Next 'Ignore error if parameter does not exist
        bAllow = g_clsGlobalParameter.FindBoolean("SupervisorShowIntermediaries")
        If Err <> 0 Then
            bAllow = False
        End If
        On Error GoTo Failed
        
    End If
' TW 20/12/2006 EP2_615 End
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName INTERMEDIARIES, INTERMEDIARIES, tvwChild, nOptionsIndex
        TreeItem.AddHeading "Name", 20
        m_treeData.Add TreeItem, INTERMEDIARIES
        nIntermediariesIndex = m_treeData.Count()

' EP15 pct
'        Set TreeItem = New TreeAccess
'        TreeItem.AddName LEADAGENT, LEADAGENT, tvwChild, nIntermediariesIndex
'        TreeItem.AddHeading "Company Name", 20
'        TreeItem.AddHeading "Commission Number", 20 ' PB EP521
'        TreeItem.AddHeading "Active From", 15
'        TreeItem.AddHeading "Active To", 15
'        TreeItem.AddHeading "Panel ID", 15
'        m_treeData.Add TreeItem, LEADAGENT
'
'        Set TreeItem = New TreeAccess
'        TreeItem.AddName INDIVIDUAL, INDIVIDUAL, tvwChild, nIntermediariesIndex
'        TreeItem.AddHeading "Title", 10
'        TreeItem.AddHeading "Forename", 20
'        TreeItem.AddHeading "Surname", 20
'        TreeItem.AddHeading "Active From", 15
'        TreeItem.AddHeading "Active To", 15
'        TreeItem.AddHeading "Panel ID", 15
'        m_treeData.Add TreeItem, INDIVIDUAL

        Set TreeItem = New TreeAccess
        TreeItem.AddName PACKAGER, PACKAGER, tvwChild, nIntermediariesIndex
        TreeItem.AddHeading "Company Name", 20
        TreeItem.AddHeading "Commission Number", 20 ' PB EP521
        TreeItem.AddHeading "Active From", 15
        TreeItem.AddHeading "Active To", 15
        TreeItem.AddHeading "Panel ID", 15
        m_treeData.Add TreeItem, PACKAGER
        
        Set TreeItem = New TreeAccess
        TreeItem.AddName BROKER, BROKER, tvwChild, nIntermediariesIndex
        TreeItem.AddHeading "Company Name", 20
        TreeItem.AddHeading "Commission Number", 20 ' PB EP521
        TreeItem.AddHeading "Active From", 15
        TreeItem.AddHeading "Active To", 15
        TreeItem.AddHeading "Panel ID", 15
        m_treeData.Add TreeItem, BROKER
   End If
' EP15 pct End



    ' SYS4142 Check for any client specific TreeView items
    Set clsTreeViewCS = New TreeViewCS
    clsTreeViewCS.SetupTreeValues m_treeData
    
    AddTreeNodes
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "Failed to set up tree values: " & Err.DESCRIPTION
    Exit Sub
    Resume
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   timerCancelForm_Timer
' Description   :   Can't think of a better way of doing this. Closes the form
'                   returned from GetCurrentForm. This is so that a from that loads
'                   then decides it wants to close due to an error, can close itself.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub timerCancelForm_Timer()
    Dim frmForm As Form
    Set frmForm = g_clsFormProcessing.GetCurrentForm()
    timerCancelForm.Enabled = False
    If Not frmForm Is Nothing Then
        
        Unload frmForm

    End If

End Sub
' Change the image of the node to "closed"
Private Sub tvwDB_Collapse(ByVal node As MSComctlLib.node)
    On Error GoTo Failed
    ' DJP BM318, when expanding a node, remove any existing selected items
    Set tvwDB.SelectedItem = Nothing
    
    node.Image = "closed"
    node.Expanded = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' Change the image of the node to "open"
Private Sub tvwDB_Expand(ByVal node As MSComctlLib.node)
    On Error GoTo Failed
    node.Image = "open"
    
    ' DJP BM318, when expanding a node, remove any existing selected items
    Set tvwDB.SelectedItem = Nothing
    
    HandleExpandedTreeNode node
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub tvwDB_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    If KeyAscii = vbKeyReturn Then
        HandleTreeClick tvwDB.SelectedItem
    End If
    Exit Sub
Failed:
        g_clsErrorHandling.DisplayError
End Sub
Private Function HandleTreeClick(tvNode As node) As Boolean
' TW 15/01/2007 EP2_826 - Rewritten to simplify
' Note. A popup menu is no longer shown on left click

Dim bPopulated As Boolean
Dim bLoaded As Boolean
Dim nCount As Long
Dim vKey As Variant

    On Error GoTo Failed
    SetTreeStatus tvNode
    bPopulated = False
    
' If any Search dialogs are open, close them all
    g_clsMainSupport.CloseFindDialogs

' Clear the listview
    lvListView.ColumnHeaders.Clear
    lvListView.ListItems.Clear
    
' Reset the ListView label
    lblTitle(constListViewLabel).FontBold = False
    lblTitle(constListViewLabel).ForeColor = vbButtonText
    lblTitle(constListViewLabel) = ""
    
    If Not tvNode Is Nothing Then
        If tvNode.children > 0 Then
            Exit Function
        End If
        vKey = GetSelectedTreeKey()
        
        If Not IsNull(vKey) Then
            lblTitle(constListViewLabel) = vKey
            Select Case (CStr(vKey))
                    
                Case LOCAL_ADDRESS, VALUER_ADDRESS, LEGAL_REP_ADDRESS, MORTGAGE_PRODUCTS, USERS, _
                     LOCKS_ONLINE_CUSTOMER, LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION, LOCKS_OFFLINE_CUSTOMER, LOCKS_DOCS, _
                     PRINCIPALS, PACKAGERS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS, _
                     UNITS, _
                     APPLICATION_PROCESSING
                     
' Don't auto populate the listview when any of the above are clicked.
                    LoadListViewHeaders
                    
                Case Else
' Otherwise, auto populate the listview
                    bPopulated = True
                    PopulateListView
            End Select
        End If
    End If
    
    HandleTreeClick = bPopulated
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   tvwDB_MouseUp
' Description   :   When the user right clicks on the treeview, need to show the menu popup.
'                   Otherwise, populate the listview
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tvwDB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tvNode As node
    Dim bShowPopup As Boolean
    On Error GoTo Failed
    
    Set tvNode = tvwDB.SelectedItem()
    
    Select Case Button
    
    Case RIGHT_MOUSE
' TW 17/10/2006 EP2_15
' I changed this code because the right click was
' causing the listview to always be populated. This caused problems when > 1000 records
' would have been displayed, when a MsgBox would be shown. Now, the popup will be displayed
' if appropriate, otherwise the listview will be populated, if less than 1000 records.

'        EnableTVPopup
'        bShowPopup = HandleTreeClick(tvNode)
'        If bShowPopup Then
'            HandleTreeViewPopup tvNode.Key
'        End If
    
        bShowPopup = EnableTVPopup
        If Not bShowPopup Then
            bShowPopup = HandleTreeClick(tvNode)
        End If
        If bShowPopup Then
            HandleTreeViewPopup tvNode.Key
        End If
' TW 17/10/2006 EP2_15 End
    
    Case LEFT_MOUSE
        ' DJP BM0318 In some cases, we may want to just display a find dialog, so this method
        ' will do it (and if not, just populate the listview as normal)
        HandleTreeClick tvNode
    
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   lvListView_MouseUp
' Description   :   When the user right clicks on the listview, need to show the menu popup.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Failed
    If (Button = RIGHT_MOUSE) Then
        HandleListViewPopup
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   HandleTreeViewPopup
' Description   :   When the user right clicks on the TreeView, we need to decide which
'                   popup menu to show.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleTreeViewPopup(sKey As String)
' TW 15/01/2007 EP2_826 - Rewritten to simplify

Dim PopUpLeft As Integer
Dim SpaceWidth As Integer
Dim N As node

    On Error GoTo Failed
    PopUpLeft = tvwDB.Indentation * 2
    SpaceWidth = Me.TextWidth(" ")
    Set N = tvwDB.SelectedItem.Parent
    Do
        If N Is Nothing Then
            Exit Do
        End If
        PopUpLeft = PopUpLeft + tvwDB.Indentation + SpaceWidth
        Set N = N.Parent
    Loop
    PopUpLeft = PopUpLeft + Me.TextWidth(tvwDB.SelectedItem.Text) + SpaceWidth
    If PopUpLeft > tvwDB.Width Then
        PopUpLeft = tvwDB.Width
    End If
    
    mnuTVOptions(TV_OPTIONS_RETRIEVE_OMIGA_USERS).Visible = False
    mnuTVOptions(TV_OPTIONS_RETRIEVE_INTRO_USERS).Visible = False
    mnuTVOptions(TV_OPTIONS_RETRIEVE_OMIGA_UNITS).Visible = False
    mnuTVOptions(TV_OPTIONS_RETRIEVE_INTRO_UNITS).Visible = False
    
    mnuLVLocks.Item(LV_FIND_LOCK).Visible = True
    mnuLVLocks.Item(LV_ALL_LOCK).Visible = True
    
    mnuTVOptions(TV_OPTIONS_NEW).Enabled = True
    mnuTVOptions(TV_OPTIONS_RETRIEVE_ALL).Enabled = False
    mnuTVOptions(TV_OPTIONS_FIND).Enabled = False
    
    Select Case sKey
        Case LENDERS, STAMP_DUTY_RATES, MIG_RATE_SETS, ADDITIONAL_BORROWING_FEES, CREDIT_LIMIT_INCREASE_FEES, ADMIN_FEES, VALUATION_FEES, INCOME_MULTIPLE, _
             BASE_RATES, MP_MIG_RATE_SETS, REDEM_FEE_SETS, COMBOBOX_ENTRIES, ERROR_MESSAGES, _
             LENDER_ADDRESS, GLOBAL_PARAM_FIXED, GLOBAL_PARAM_BANDED, _
             DISTRIBUTION_CHANNELS, DEPARTMENTS, REGIONS, COMPETENCIES, WORKING_HOURS, _
             BUILDINGS_AND_CONTENTS_PRODUCTS, PAYMENT_PROTECTION_RATES, PAYMENT_PROTECTION_PRODUCTS, COUNTRIES, _
             LIFE_COVER_RATES, TASK_MANAGEMENT_TASKS, TASK_MANAGEMENT_STAGES, TASK_MANAGEMENT_ACTIVITIES, ADDITIONAL_QUESTIONS, _
             BATCH_SCHEDULER, BASE_RATE, PRINTING_TEMPLATE, PRINTING_DOCUMENT, PRINTING_PACK, BUSINESS_GROUPS, CURRENCIES, INCOME_FACTORS, PRODUCT_SWITCH_FEESETS, _
             INSURANCE_ADMIN_FEESETS, TT_FEESETS, RENTAL_INCOME_RATES, TRANSFER_OF_EQUITY_FEES, _
             DEFAULT_PROCURATION_FEES, LOAN_AMOUNT_ADJUSTMENTS, LTV_AMOUNT_ADJUSTMENTS, _
             CONDITIONS

            PopupMenu mnuTopLevel(TV_OPTIONS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuTVOptions(TV_OPTIONS_NEW)
        
        Case CLUBS, ASSOCIATIONS, PACKAGERS, PRINCIPALS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS
            mnuTVOptions(TV_OPTIONS_FIND).Enabled = True
            mnuTVOptions(TV_OPTIONS_RETRIEVE_ALL).Enabled = True
            
            PopupMenu mnuTopLevel(TV_OPTIONS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuTVOptions(TV_OPTIONS_NEW)
        
        Case UNITS
            mnuTVOptions(TV_OPTIONS_RETRIEVE_OMIGA_UNITS).Visible = True
            mnuTVOptions(TV_OPTIONS_RETRIEVE_INTRO_UNITS).Visible = True
            mnuTVOptions(TV_OPTIONS_RETRIEVE_ALL).Enabled = True
            mnuTVOptions(TV_OPTIONS_FIND).Enabled = True
            
            PopupMenu mnuTopLevel(TV_OPTIONS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuTVOptions(TV_OPTIONS_NEW)
        
        Case USERS
            mnuTVOptions(TV_OPTIONS_RETRIEVE_OMIGA_USERS).Visible = True
            mnuTVOptions(TV_OPTIONS_RETRIEVE_INTRO_USERS).Visible = True
            mnuTVOptions(TV_OPTIONS_RETRIEVE_ALL).Enabled = True
            mnuTVOptions(TV_OPTIONS_FIND).Enabled = True
            
            PopupMenu mnuTopLevel(TV_OPTIONS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuTVOptions(TV_OPTIONS_NEW)
        
        Case LOCAL_ADDRESS, LEGAL_REP_ADDRESS, VALUER_ADDRESS, MORTGAGE_PRODUCTS
            PopupMenu mnuTopLevel(LV_THIRD_PARTY_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVThirdParties(LV_THIRD_PARTY_MENU_ADD)
        
        Case APPLICATION_PROCESSING
            mnuTVOptions(TV_OPTIONS_NEW).Enabled = False
        
            PopupMenu mnuTopLevel(TV_OPTIONS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuTVOptions(TV_OPTIONS_FIND)
        
        Case LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION
            ' Need to display both Create Application and Find Application
            mnuLVLocks.Item(LV_CREATE_LOCK).Visible = True
            mnuLVLocks.Item(LV_REMOVE_LOCK).Visible = False
            
            PopupMenu mnuTopLevel(LV_LOCK_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVLocks(LV_FIND_LOCK)
        
        Case LOCKS_ONLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER, LOCKS_DOCS
            ' Don't display Create Customer, just find
            mnuLVLocks.Item(LV_CREATE_LOCK).Visible = False
            mnuLVLocks.Item(LV_REMOVE_LOCK).Visible = False
            
            PopupMenu mnuTopLevel(LV_LOCK_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVLocks(LV_FIND_LOCK)
        
        Case Else
            Dim clsTreeViewCS As TreeViewCS
            Set clsTreeViewCS = New TreeViewCS

            clsTreeViewCS.HandleTreeViewPopup sKey

            
    End Select
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetSelectedTreeKey() As Variant
    Dim mNode As node
    Dim clsTreeItem As TreeItem
    ' Which node was selected?
       
    Set mNode = tvwDB.SelectedItem
    

    If (Not mNode Is Nothing) Then
        If TypeOf mNode.Tag Is TreeItem Then
            Set clsTreeItem = mNode.Tag
            GetSelectedTreeKey = clsTreeItem.GetType
        Else
            GetSelectedTreeKey = mNode.Key
        End If
    Else
        GetSelectedTreeKey = Null
    End If
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableMenu
' Description   : Enables an item(s) on a menu control.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnableMenu(ByRef mnu As Object, Optional ByVal bEnable As Boolean = True, Optional ByVal nEnable As Variant, Optional ByVal bExclusive As Boolean = True)
    
    Dim nCount As Integer
    Dim mnuItem As Menu
    
    On Error GoTo Failed
    
    'Iterate through each menu item in the supplied menu.
    For nCount = 0 To mnu.Count - 1
        'Get a reference to the item so we can work on a local reference.
        Set mnuItem = mnu(nCount)
        
        'If a specific item was supplied,
        If Not IsMissing(nEnable) Then
            If nCount = nEnable Then
                'If this item is the specific item then set its enabled state
                'to whatever was specified.
                mnuItem.Enabled = bEnable
            ElseIf bExclusive Then
                'If it isn't and the exclusive flag is set, toggle the enabled
                'state of the menu item.
                mnuItem.Enabled = Not bEnable
            End If
        Else
            'If no specific menu item was given, just set the enabled state of
            'the current item to whatever was specified.
            mnuItem.Enabled = bEnable
        End If
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   HandleListViewPopup
' Description   :   When the user right clicks on the ListView, we need to
'                   decide which popup menu to show. If vblnPerformDefaultAction
'                   is true then call the click event of the default menu option,
'                   otherwise display the popup menu
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleListViewPopup(Optional ByVal vblnPerformDefaultAction As Boolean = False)
' TW 15/01/2007 EP2_826 - Rewritten to simplify
    
Dim vKey As Variant
Dim objSelection As Object
Dim bMultiSelect As Boolean
Dim bProductValid As Boolean
Dim bEnableDelete As Boolean
Dim bActivateValid As Boolean
Dim PopUpLeft As Integer

    On Error GoTo Failed
    PopUpLeft = lvListView.Left + lvListView.ColumnHeaders(1).Width + 60
        
    bEnableDelete = False
    
    ' Which node was selected?
    vKey = GetSelectedTreeKey()
    
    Set objSelection = lvListView.SelectedItem

    If (Not objSelection Is Nothing) Then
        If (Not IsNull(vKey)) Then
            'Set the multi-select flag if more than one item is selected.
            bMultiSelect = (GetSelectedCount() > 1)
            
            'Set default menu settings
            EnableMenu mnuLVEdit, True
            If bMultiSelect Then
                bEnableDelete = True
                EnableMenu mnuLVEdit, False, LV_EDIT
            End If
            
            Select Case vKey
                'Common items which can be deleted.
                Case ADMIN_FEES, _
                    VALUATION_FEES, MP_MIG_RATE_SETS, REDEM_FEE_SETS, _
                    INCOME_MULTIPLE, GLOBAL_PARAM_FIXED, _
                    GLOBAL_PARAM_BANDED, COMBOBOX_ENTRIES, ERROR_MESSAGES, _
                    PAYMENT_PROTECTION_RATES, COUNTRIES, _
                    PAYMENT_PROTECTION_PRODUCTS, LIFE_COVER_RATES, _
                    BUILDINGS_AND_CONTENTS_PRODUCTS, _
                    TASK_MANAGEMENT_TASKS, TASK_MANAGEMENT_ACTIVITIES, _
                    TASK_MANAGEMENT_STAGES, ADDITIONAL_QUESTIONS, _
                    CONDITIONS, PRINTING_TEMPLATE, PRINTING_DOCUMENT, _
                    BUSINESS_GROUPS, CURRENCIES, INCOME_FACTORS, PRODUCT_SWITCH_FEESETS, _
                    INSURANCE_ADMIN_FEESETS, TT_FEESETS, RENTAL_INCOME_RATES, _
                    CREDIT_LIMIT_INCREASE_FEES, ADDITIONAL_BORROWING_FEES, TRANSFER_OF_EQUITY_FEES, _
                    DEFAULT_PROCURATION_FEES, LOAN_AMOUNT_ADJUSTMENTS, LTV_AMOUNT_ADJUSTMENTS, _
                    USERS, _
                    LEGAL_REP_ADDRESS, VALUER_ADDRESS, LOCAL_ADDRESS
                    
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVEdit_Click LV_EDIT
                    Else
                        PopupMenu mnuTopLevel(LV_EDIT_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVEdit(LV_EDIT)
                    End If
                
                Case BASE_RATES, BASE_RATE
                ' Delete enabled for Base Rate Set, disabled for Base Rate
                    EnableMenu mnuLVBaseRates, (vKey = BASE_RATES), LV_BASERATE_DELETE, False
                    If vblnPerformDefaultAction Then
                        mnuLVBaseRates_Click LV_BASERATE_VIEW
                    Else
                        PopupMenu mnuTopLevel(LV_BASERATE_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVBaseRates(LV_BASERATE_VIEW)
                    End If
                    
                'Common items which cannot be deleted.
                Case DISTRIBUTION_CHANNELS, REGIONS, COMPETENCIES, _
                    WORKING_HOURS, UNITS, DEPARTMENTS, _
                    INTERMEDIARIES, LEADAGENT, INTERMEDIARY_COMPANY, _
                    ADMIN_CENTRE, INDIVIDUAL, PACKAGER, BROKER, _
                    CLUBS, ASSOCIATIONS, PRINCIPALS, PACKAGERS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS, _
                    LENDERS, PRINTING_PACK
                    
                  ' Edit and Mark for Promotion allowed, Delete not allowed
                    EnableMenu mnuLVEdit, False, LV_DELETE, False
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVEdit_Click LV_EDIT
                    Else
                        PopupMenu mnuTopLevel(LV_EDIT_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVEdit(LV_EDIT)
                    End If
                    
                Case MORTGAGE_PRODUCTS
                    If bMultiSelect Then
                        ' Just enable delete and promote
                        EnableMenu mnuLVProducts, True, LV_PRODUCTS_DELETE
                        EnableMenu mnuLVProducts, True, LV_PRODUCTS_PROMOTE, False
                    Else
                        ' Is the product Valid?
                        bProductValid = g_clsMainSupport.IsProductValid()
                        
                        'If product valid then enable promote and disable view errors
                        If bProductValid = True Then
                            EnableMenu mnuLVProducts, bProductValid, LV_PRODUCTS_PROMOTE, False
                            EnableMenu mnuLVProducts, Not bProductValid, LV_PRODUCTS_VIEW_ERRORS, False
                        Else
                            EnableMenu mnuLVProducts, bProductValid, LV_PRODUCTS_VIEW_ERRORS, False
                            EnableMenu mnuLVProducts, bProductValid, LV_PRODUCTS_PROMOTE, True
                        End If
                        'DB END
                    End If
                                            
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVProducts_Click LV_PRODUCTS_EDIT
                    Else
                        PopupMenu mnuTopLevel(LV_PRODUCTS_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVProducts(LV_PRODUCTS_EDIT)
                    End If
                    
                Case BATCH_SCHEDULER
                    EnableMenu mnuLVBatch, True
                    
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVBatch_Click LV_BATCH_PROCESS_VIEW
                    Else
                        PopupMenu mnuTopLevel(LV_BATCH_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVBatch(LV_BATCH_PROCESS_VIEW)
                    End If
                    
' TW 27/11/2007 DBM594
                Case PAYMENTS_FOR_COMPLETION
                  ' Edit allowed, Mark for Promotion and Delete not allowed
                    EnableMenu mnuLVEdit, True
                    mnuLVEdit(LV_DELETE).Enabled = False
                    mnuLVEdit(LV_MARK_FOR_PROMOTION).Enabled = False
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVEdit_Click LV_EDIT
                    Else
                        PopupMenu mnuTopLevel(LV_EDIT_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVEdit(LV_EDIT)
                    End If
' TW 27/11/2007 DBM594 End
                    
                    
                Case APPLICATION_PROCESSING
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVAppProcessing_Click LV_CANCEL_APPLICATION
                    Else
                        PopupMenu mnuTopLevel(LV_APPLICATION_PROCESSING), vbPopupMenuRightButton, PopUpLeft, , mnuLVAppProcessing(LV_CANCEL_APPLICATION)
                    End If
                
                Case _
                    LOCKS_ONLINE_APPLICATION, _
                    LOCKS_OFFLINE_APPLICATION, _
                    LOCKS_ONLINE_CUSTOMER, _
                    LOCKS_OFFLINE_CUSTOMER, _
                    LOCKS_DOCS
                    bEnableDelete = True
    
                    mnuLVLocks.Item(LV_REMOVE_LOCK).Visible = True
                    mnuLVLocks.Item(LV_FIND_LOCK).Visible = False
                    mnuLVLocks.Item(LV_CREATE_LOCK).Visible = False
                    mnuLVLocks.Item(LV_FIND_LOCK).Visible = False
                    mnuLVLocks.Item(LV_ALL_LOCK).Visible = False
                    
                    'Click the default right click option if required, otherwise popup the menu
                    If vblnPerformDefaultAction Then
                        mnuLVLocks_Click LV_REMOVE_LOCK
                    Else
                        PopupMenu mnuTopLevel(LV_LOCK_MENU), vbPopupMenuRightButton, PopUpLeft, , mnuLVLocks(LV_REMOVE_LOCK)
                    End If
                    
                Case Else
                    Dim clsListViewCS As ListViewCS
                    Set clsListViewCS = New ListViewCS
                    clsListViewCS.HandleListViewPopup vKey
            
            End Select
        
        Else
            MsgBox ("No item selected")
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub LoadPosition()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ListView Menu events
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Main Menu events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuLVLenders_Click(Index As Integer)
    On Error GoTo Failed
    g_clsMainSupport.HandleLenders Index
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub mnuLVEdit_Click(Index As Integer)
    Timer1.Interval = 1
    Timer1.Enabled = True
        
    Select Case Index
        Case LV_EDIT
            Timer1.Tag = "ListViewEdit"
        Case LV_DELETE
            Timer1.Tag = "ListViewDelete"
    
        Case LV_MARK_FOR_PROMOTION
            Timer1.Tag = "ListViewPromote"
    End Select
End Sub
' Called when the right mouse button is clicked on the TreeView.
Private Sub mnuTVOptions_Click(Index As Integer)
    On Error GoTo Failed
    Select Case Index
    
    Case TV_OPTIONS_NEW
        Timer1.Interval = 1
        Timer1.Enabled = True
        Timer1.Tag = "TreeView"
        
    Case TV_OPTIONS_FIND
        g_clsMainSupport.HandleTreeFind
    
    'PB 13/10/2006
    Case TV_OPTIONS_RETRIEVE_ALL
        ' retrieve all
        PopulateListView False, LV_DETAIL_ALL
    
    Case TV_OPTIONS_RETRIEVE_OMIGA_UNITS
        ' retrieve omiga units
        PopulateListView False, LV_DETAIL_OMIGA_UNITS
    
    Case TV_OPTIONS_RETRIEVE_INTRO_UNITS
        ' guess what?
        PopulateListView False, LV_DETAIL_INTRO_UNITS
    
    'PB 13/10/2006 End
' TW 18/12/2006 EP2_568
    Case TV_OPTIONS_RETRIEVE_OMIGA_USERS
        PopulateListView False, LV_USERS_OMIGA_USERS
    Case TV_OPTIONS_RETRIEVE_INTRO_USERS
        PopulateListView False, LV_USERS_INTRO_USERS
' TW 18/12/2006 EP2_568 End
    
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub mnuLVUsers_Click(Index As Integer)
    On Error GoTo Failed
    Select Case Index
    Case LV_USER_EDIT
        g_clsMainSupport.HandleUsers LV_USER_EDIT, UserDetails
    
    Case LV_USER_DELETE
        g_clsMainSupport.HandleUsers LV_USER_DELETE
    
    Case LV_USER_QUALIFICATIONS
        g_clsMainSupport.HandleUsers LV_USER_EDIT, Qualifications
    
    Case LV_USER_COMPETENCY_HISTORY
        g_clsMainSupport.HandleUsers LV_USER_EDIT, CompetencyHistory
    
    Case LV_USER_USER_ROLES
        g_clsMainSupport.HandleUsers LV_USER_EDIT, UserRoles
    
    Case LV_USER_PROMOTE
        g_clsMainSupport.PromoteRecord
    End Select
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub mnuLVAddress_Click(Index As Integer)
    On Error GoTo Failed
    Select Case Index
    Case LV_ADDRESS_EDIT
        g_clsMainSupport.HandleEditThirtParty
    Case LV_ADDRESS_ACTIVATE ' TK 30/11/2005 MAR81
        g_clsMainSupport.HandleActivateThirtParty
    Case LV_ADDRESS_BANK_DETAILS
        g_clsMainSupport.HandleEditThirtParty TabBankAccounts
    Case LV_ADDRESS_CONTACT_DETAILS
        g_clsMainSupport.HandleEditThirtParty TabContactDetails
    Case LV_ADDRESS_LEGAL_REP
        g_clsMainSupport.HandleEditThirtParty TabLegalRepDetails
    Case LV_ADDRESS_VALUER
        g_clsMainSupport.HandleEditThirtParty TabValuerDetails
    Case LV_ADDRESS_DELETE
        g_clsMainSupport.DeleteRecord
    
    Case LV_ADDRESS_PROMOTE
        g_clsMainSupport.PromoteRecord
    
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Public Function LoadListViewHeaders() As Boolean
    On Error GoTo Failed
    Dim tvNode As node
    Dim sKey As String
    Dim sCurrentKey As String
    Dim bLoaded As Boolean
    Dim bRet As Boolean
    
    bLoaded = False
    Set tvNode = tvwDB.SelectedItem()

    If Not tvNode Is Nothing Then
        sbStatusBar.Panels("Status").Text = tvNode.Text
        'BS BM0282 25/03/03
        'lblTitle(constListViewLabel).Caption = tvNode.Text

        sKey = tvNode.Key

        bRet = DoesNodeHaveHeader(tvNode, sKey)
        sCurrentKey = lvListView.SaveColumnDetails()

        If sCurrentKey <> sKey And bRet = True Or lvListView.ColumnHeaders.Count = 0 Then
            'BS BM0282 25/03/03
            'If a different key has been clicked then ok to refresh the title here, otherwise we
            'have to wait to see what the user does to work out if the title should be refreshed
            lblTitle(constListViewLabel).Caption = tvNode.Text

            lvListView.ListItems.Clear
            If Not bRet Then
                SetColumnHeader tvNode.Text
            Else
                SetColumnHeader sKey
            End If
            lvListView.LoadColumnDetails sKey
            bLoaded = True
        Else
            bLoaded = False
        End If
    End If
    
    
    LoadListViewHeaders = bLoaded
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateListView
' Description   :   When the user clicks on a node on the treeview we need to
'                   populate the listview with the items required. Methods that obtain
'                   the data is retrieved from the main form support class, g_clsMainSupport.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub PopulateListView(Optional bUseExistingSearch As Boolean = False, Optional intDetail As LV_DETAIL = LV_DETAIL_NONE)
    On Error GoTo Failed
    Dim nCols As Long
    Dim nThisCount As Long
    Dim cLine As Collection
    Dim tvNode As node
    Dim sKey As String
    Dim bMultiSelect As Boolean
    Dim clsTreeItem As TreeItem
    'Dim objTag As Object
    On Error GoTo Failed
                
    bMultiSelect = True
    
    BeginWaitCursor
    
    'PB 17/10/2006 Begin
    ' We need to remember this value in case of a page refresh
    If intDetail = LV_DETAIL_NONE Then
        intDetail = m_intDetail
    End If
    m_intDetail = intDetail
    'PB End
    '
    Set tvNode = tvwDB.SelectedItem()

    If Not tvNode Is Nothing Then
        
        sKey = tvNode.Key
        ' Default to single select
        lvListView.MultiSelect = False
        lvListView.ListItems.Clear
        g_clsMainSupport.SetRecordSet Nothing
            
        'Is there a tree item class held in the tag of this tree node?
        If TypeOf tvNode.Tag Is TreeItem Then
            Set clsTreeItem = tvNode.Tag
            sKey = clsTreeItem.GetType
        Else
            sKey = tvNode.Key
            Set clsTreeItem = New TreeItem
            clsTreeItem.SetType sKey
        End If
        
        On Error GoTo Failed
        
'        lvListView.SaveColumnDetails
       'SetColumnHeader tvNode.Text
 '      lvListView.LoadColumnDetails sKey
        LoadListViewHeaders
        ' Obtain the data
        nCols = lvListView.ColumnHeaders.Count
        Set cLine = New Collection
        
        Select Case sKey
        
        Case LENDERS
            'bMultiSelect = False
            g_clsMainSupport.GetLenders lvListView
        Case MORTGAGE_PRODUCTS
            g_clsMainSupport.GetProducts lvListView
            
' TW 09/10/2006 EP2_7
        Case ADDITIONAL_BORROWING_FEES
            g_clsMainSupport.GetData New AdditionalBorrowingFeeTable, lvListView, POPULATE_FIRST_BAND
        Case CREDIT_LIMIT_INCREASE_FEES
            g_clsMainSupport.GetData New CreditLimitIncreaseFeeTable, lvListView, POPULATE_FIRST_BAND
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case TRANSFER_OF_EQUITY_FEES
            g_clsMainSupport.GetData New TransferOfEquityFeeTable, lvListView, POPULATE_FIRST_BAND
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
        Case DEFAULT_PROCURATION_FEES
            g_clsMainSupport.GetDefaultProcurationFees lvListView
        Case LOAN_AMOUNT_ADJUSTMENTS
            g_clsMainSupport.GetLoanAmountAdjustments lvListView
        Case LTV_AMOUNT_ADJUSTMENTS
            g_clsMainSupport.GetLTVAmountAdjustments lvListView
' TW 14/12/2006 EP2_518 End
        Case ADMIN_FEES
            'g_clsMainSupport.GetAdminFees lvListView
            g_clsMainSupport.GetData New AdminFeeTable, lvListView, POPULATE_FIRST_BAND
        Case VALUATION_FEES
            'g_clsMainSupport.GetValuationFees lvListView
            g_clsMainSupport.GetData New ValuationFeeTable, lvListView, POPULATE_FIRST_BAND
        '*=[MC]BMIDS763 - CC075 - FEESETS
        Case PRODUCT_SWITCH_FEESETS
            g_clsMainSupport.GetData New ProductSwitchFeeTable, lvListView, POPULATE_FIRST_BAND
        Case INSURANCE_ADMIN_FEESETS
            g_clsMainSupport.GetData New InsuranceAdminFeeBand, lvListView, POPULATE_FIRST_BAND
        Case TT_FEESETS
            g_clsMainSupport.GetData New TTFeeBand, lvListView, POPULATE_FIRST_BAND
        '*=[MC] BMIDS763 -CC075 FEE SETS CHANGE END
        ' BMIDS765 JD
        Case RENTAL_INCOME_RATES
            g_clsMainSupport.GetData New RentalIncomeRateSetBandTable, lvListView, POPULATE_FIRST_BAND
        Case BASE_RATES
            'g_clsMainSupport.GetBaseRates lvListView
            g_clsMainSupport.GetData New BaseRateTable, lvListView, POPULATE_FIRST_BAND
        '   AW  13/05/02    BM088
        Case INCOME_MULTIPLE
            g_clsMainSupport.GetData New IncomeMultipleSetTable, lvListView, POPULATE_ALL
        '   AW  16/05/02    BM087
        Case MP_MIG_RATE_SETS
            g_clsMainSupport.GetData New MPMigRateSetTable, lvListView, POPULATE_ALL
        '   AW  16/05/02    BM017
        Case REDEM_FEE_SETS
            g_clsMainSupport.GetData New RedemptionFeeSetTable, lvListView, POPULATE_ALL
        Case LEGAL_REP_ADDRESS
            g_clsMainSupport.GetPanelAddress lvListView, ThirdPartyLegalRep
        Case VALUER_ADDRESS
            g_clsMainSupport.GetPanelAddress lvListView, ThirdPartyValuer
        Case LOCAL_ADDRESS
            g_clsMainSupport.GetLocalAddress lvListView
        Case GLOBAL_PARAM_FIXED
            g_clsMainSupport.GetData New FixedParametersTable, lvListView
        Case GLOBAL_PARAM_BANDED
            g_clsMainSupport.GetData New BandedGlobalParametersTable, lvListView, POPULATE_FIRST_BAND
        Case COMBOBOX_ENTRIES
            g_clsMainSupport.GetData New ComboValueGroupTable, lvListView
        Case COUNTRIES
            g_clsMainSupport.GetData New CountryTable, lvListView
        Case ERROR_MESSAGES
            g_clsMainSupport.GetData New ErrorMessageTable, lvListView
        Case DISTRIBUTION_CHANNELS
            g_clsMainSupport.GetData New DistributionChannelTable, lvListView
        Case APPLICATION_PROCESSING
            bMultiSelect = False
            g_clsMainSupport.GetActiveApplications lvListView, bUseExistingSearch
        Case DEPARTMENTS
            g_clsMainSupport.GetData New DepartmentTable, lvListView
        Case REGIONS
            g_clsMainSupport.GetRegions lvListView
        Case UNITS
            ' MUST provide detail level if units
            Select Case intDetail
                Case LV_DETAIL_ALL
                    ' Show all units, i.e. Omiga and Introducer
                    g_clsMainSupport.GetData New UnitTable, lvListView
                Case LV_DETAIL_OMIGA_UNITS
                    ' Show only omiga units
                    g_clsMainSupport.GetData New OmigaUnitTable, lvListView
                Case LV_DETAIL_INTRO_UNITS
                    ' Show only intro units
                    g_clsMainSupport.GetData New IntroUnitTable, lvListView
                Case Else ' default = same as LV_DETAIL_ALL
                    g_clsMainSupport.GetData New UnitTable, lvListView
            End Select
        Case COMPETENCIES
            g_clsMainSupport.GetCompetencies lvListView
        Case WORKING_HOURS
            g_clsMainSupport.GetData New WorkingHourTypeTable, lvListView
' TW 18/12/2006 EP2_568
        Case USERS
'            g_clsMainSupport.GetOmigaUsers lvListView
            g_clsMainSupport.GetOmigaUsers lvListView, intDetail
' TW 18/12/2006 EP2_568 End
        Case LIFE_COVER_RATES
            g_clsMainSupport.GetLifeCoverRates lvListView
        Case BUILDINGS_AND_CONTENTS_PRODUCTS
            g_clsMainSupport.GetData New BuildingAndContentsTable, lvListView
        Case PAYMENT_PROTECTION_RATES
            g_clsMainSupport.GetPaymentProtectionRates lvListView
        Case PAYMENT_PROTECTION_PRODUCTS
            g_clsMainSupport.GetData New PayProtProductsTable, lvListView
        Case LOCKS_ONLINE_APPLICATION
            g_clsMainSupport.GetApplicationLocks lvListView, OnLine
        Case LOCKS_ONLINE_CUSTOMER
            g_clsMainSupport.GetCustomerLocks lvListView, OnLine
        Case LOCKS_OFFLINE_APPLICATION
            g_clsMainSupport.GetApplicationLocks lvListView, Offline
        Case LOCKS_OFFLINE_CUSTOMER
            g_clsMainSupport.GetCustomerLocks lvListView, Offline
        ' ik_bm0314
        Case LOCKS_DOCS
            g_clsMainSupport.GetDocumentLocks frmMain.lvListView
        ' --- DJP Phase 2 Task Management Start ---
        Case TASK_MANAGEMENT_STAGES
             g_clsMainSupport.GetStages lvListView

        Case TASK_MANAGEMENT_TASKS
             g_clsMainSupport.GetTasks lvListView
        
        Case TASK_MANAGEMENT_ACTIVITIES
             g_clsMainSupport.GetActivities lvListView
        ' --- DJP Phase 2 Task Management End ---
        Case BUSINESS_GROUPS
            g_clsMainSupport.GetBusinessGroups lvListView
        Case ADDITIONAL_QUESTIONS
            g_clsMainSupport.GetQuestions lvListView
        
        Case CONDITIONS
            g_clsMainSupport.GetConditions lvListView
        Case CURRENCIES
             g_clsMainSupport.GetCurrencies lvListView
        Case BATCH_SCHEDULER
            g_clsMainSupport.GetBatch lvListView
' TW 27/11/2007 DBM594
        Case PAYMENTS_FOR_COMPLETION
            g_clsMainSupport.GetPaymentsForCompletion lvListView
' TW 27/11/2007 DBM594 End
        Case BASE_RATE
            g_clsMainSupport.GetRate lvListView
        Case PRINTING_TEMPLATE
            g_clsMainSupport.GetPrintTemplates lvListView
        'MAR45 GHun
        Case PRINTING_DOCUMENT
            g_clsMainSupport.GetPrintDocuments lvListView
        'MAR45 End
        'MAR202 GHun
        Case PRINTING_PACK
            g_clsMainSupport.GetPrintPacks lvListView
        'MAR202 End
        Case LEADAGENT, INTERMEDIARY_COMPANY, ADMIN_CENTRE, BROKER, PACKAGER ' EP15 pct
            g_clsMainSupport.GetIntermediaries lvListView, clsTreeItem, tvNode
        Case INDIVIDUAL
            g_clsMainSupport.GetIntermediaries lvListView, clsTreeItem, tvNode
        Case INCOME_FACTORS
            g_clsMainSupport.GetIncomeFactors lvListView
' TW 17/10/2006 EP2_15
        'PB 18/10/2006 E2_13 Begin
        Case ASSOCIATIONS
' TW 16/01/2007 EP2_859
'            g_clsMainSupport.GetData clsTableAccess:=New MortgageClubNetAssocTable, lvListView:=lvListView, enumClassOption:=OPTION_CLUB_ASSOCIATION_ASSOCIATION
            g_clsMainSupport.GetAssociations lvListView
' TW 16/01/2007 EP2_859 End
        Case CLUBS
' TW 16/01/2007 EP2_859
'            g_clsMainSupport.GetData clsTableAccess:=New MortgageClubNetAssocTable, lvListView:=lvListView, enumClassOption:=OPTION_CLUB_ASSOCIATION_CLUB
            g_clsMainSupport.GetClubs lvListView
' TW 16/01/2007 EP2_859 End
        Case PRINCIPALS
' TW 16/01/2007 EP2_859
'            g_clsMainSupport.GetData clsTableAccess:=New PrincipalFirmTable, lvListView:=lvListView, enumClassOption:=OPTION_PRINCIPALFIRM_FIRMBROKER
            g_clsMainSupport.GetPrincipals lvListView
' TW 16/01/2007 EP2_859 End
        Case PACKAGERS
' TW 16/01/2007 EP2_859
'            g_clsMainSupport.GetData clsTableAccess:=New PrincipalFirmTable, lvListView:=lvListView, enumClassOption:=OPTION_PRINCIPALFIRM_FIRMPACKAGER
            g_clsMainSupport.GetFirmPackagers lvListView
' TW 16/01/2007 EP2_859 End
        Case ARFIRMS
' TW 16/01/2007 EP2_859
'            g_clsMainSupport.GetData clsTableAccess:=New ARFirmTable, lvListView:=lvListView, enumClassOption:=OPTION_NONE
            g_clsMainSupport.GetARFirms lvListView
' TW 16/01/2007 EP2_859 End
        Case INDIVIDUALS & "_" & PACKAGERS
            g_clsMainSupport.GetIndividualPackagers lvListView
        Case ARBROKERS
            g_clsMainSupport.GetIndividualARBrokers lvListView
        Case DABROKERS
            g_clsMainSupport.GetIndividualDABrokers lvListView
        
        'PB 18/10/2006 E2_13 End
' TW 17/10/2006 EP2_15 End
        Case Else
            ' DJP SYS4142
            Dim bSuccess As Boolean
            Dim clsListView As ListViewCS
            
            Set clsListView = New ListViewCS
            bSuccess = clsListView.PopulateListView(sKey, lvListView)
            
            If Not bSuccess Then
                lvListView.ColumnHeaders.Clear
                lblTitle(constListViewLabel).Caption = ""
            End If
        End Select
        
        frmMain.lvListView.MultiSelect = bMultiSelect
    End If
    
' TW 15/01/2007 EP2_826
    nThisCount = lvListView.ListItems.Count
    Select Case nThisCount
        Case 0
            lblTitle(constListViewLabel).FontBold = True
            lblTitle(constListViewLabel).ForeColor = vbRed
            lblTitle(constListViewLabel).Caption = tvwDB.SelectedItem() & ": No records returned"
        Case 1
            lblTitle(constListViewLabel).Caption = tvwDB.SelectedItem() & ": 1 record"
        Case Else
            lblTitle(constListViewLabel).Caption = tvwDB.SelectedItem() & ": " & nThisCount & " records"
    End Select
' TW 15/01/2007 EP2_826 End
    
    g_clsDataAccess.CloseConnection
    EndWaitCursor

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetTreeStatus(Optional node As MSComctlLib.node = Nothing)
    If Not m_nodeCurrent Is Nothing Then
        m_nodeCurrent.Image = "closed"
    End If
    
    If node Is Nothing Then
        Set node = tvwDB.SelectedItem
    End If
    
    If Not node Is Nothing Then
        If node.Expanded = False And node.children = False Then
            node.Image = "open"
            Set m_nodeCurrent = node
        ElseIf node.children = True Then
            node.Expanded = True
            node.Image = "open"
        End If
    End If
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Public Function GetSelectedCount() As Long
    Dim nCount As Long
    Dim lstItem As ListItem
    
    nCount = 0
    
    For Each lstItem In frmMain.lvListView.ListItems
        If lstItem.Selected = True Then
            nCount = nCount + 1
        End If
    Next
    GetSelectedCount = nCount
End Function

Private Function EnableTVPopup() As Boolean
    Dim bRet As Boolean
    Dim sKey As String
    Dim tvNode As node
    
    On Error GoTo Failed
    
    Set tvNode = tvwDB.SelectedItem()
    
    bRet = True
    
    sKey = tvNode.Key

    Select Case sKey

        Case Else
            bRet = HandleTreeViewPopupExtra(sKey)

    End Select
    
    
    EnableTVPopup = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub HandleExpandedTreeNode(node As MSComctlLib.node)
    On Error GoTo Failed
    Dim sKey As String
    Dim clsTreeItem As TreeItem
    'Dim bHeader As Boolean
    Dim clsIntermediary As Intermediary
    
    Set clsIntermediary = New Intermediary
    
    If TypeOf node.Tag Is TreeItem Then
        Set clsTreeItem = node.Tag
        sKey = clsTreeItem.GetType
        
    Else
        sKey = node.Key
    End If
    
    
    Select Case sKey

        Case INTERMEDIARIES
            g_clsMainSupport.PopulateIntermediaries tvwDB, node, sKey
        Case LEADAGENT, ADMIN_CENTRE, INTERMEDIARY_COMPANY, INDIVIDUAL, BROKER, PACKAGER ' EP15 pct
            g_clsMainSupport.PopulateSingleIntermediary tvwDB, node, clsTreeItem
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Private Function DoesNodeHaveHeader(node As MSComctlLib.node, sKey As String) As Boolean
    On Error Resume Next
    
    Dim bRet As Boolean
    'Dim col As Collection
    Dim clsTreeItem As TreeItem
        
    'Gets the headings for the parent node
    bRet = True
    
    If TypeOf node.Tag Is TreeItem Then
        Set clsTreeItem = node.Tag
        'If InStr(1, node.Key, clsTreeItem.GetType) = 0 Then
            sKey = clsTreeItem.GetType
            If sKey <> INDIVIDUAL Or node.Text = INTERMEDIARY_COMPANY Or node.Text = ADMIN_CENTRE Then
                sKey = LEADAGENT
                sKey = PACKAGER
            End If
            bRet = True
        'Else
'            bRet = False
       'End If
    End If
    
    DoesNodeHaveHeader = bRet
End Function

Private Function HandleTreeViewPopupExtra(sKey As String) As Boolean
    ' pct DEBUG - On Error GoTo Failed
    
    Dim sTitle As String
    Dim bParentSet As Boolean
    Dim node As MSComctlLib.node
    Dim bRet As Boolean
    Dim bFoundParent As Boolean
    Dim nGuidEnd As Integer
    Dim clsTreeItem As TreeItem
    Dim bHeaderNode As Boolean
    
    Set clsTreeItem = New TreeItem
    On Error GoTo 0
    Set node = tvwDB.Nodes(sKey)
    sTitle = node.Text

    If TypeOf node.Tag Is TreeItem Then
        Set clsTreeItem = node.Tag
        bHeaderNode = clsTreeItem.GetIsHeader
    End If
    
    bParentSet = False
    
    bRet = True
    
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = True
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False ' pct EP15 16/03/2005 - New line
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = True ' pct EP15 16/03/2005 - New line
    mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = True ' pct EP15 16/03/2005 - New line
    mnuLVIntermediaries(LV_INTERMEDIARY_MENU_FIND).Enabled = False
    mnuLVIntermediaries(LV_INTERMEDIARY_MENU_REFRESH).Enabled = False
    
    ' pct - Traverse the tree looking for the Type of Intermediary.
    ' pct - This determines which options will be displayed in the pop-up menu.
    ' pct - Set the bFoundParent flag when we have got the info we need.
    While bFoundParent = False And sTitle <> LV_OPTIONS

        Select Case sTitle
            Case INTERMEDIARIES
                bFoundParent = True
                
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                mnuLVIntermediaries(LV_INTERMEDIARY_MENU_FIND).Enabled = True
                mnuLVIntermediaries(LV_INTERMEDIARY_MENU_REFRESH).Enabled = True

' EP15 pct
            Case PACKAGER
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = False
                If bHeaderNode Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = True
                End If
                bFoundParent = True
                    
                If node.Key = PACKAGER And Not bParentSet Then
                    bRet = True
                Else
                    bRet = False
                End If
                
            Case BROKER
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = False
                If bHeaderNode Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = True
                End If
                bFoundParent = True
                    
                If node.Key = BROKER And Not bParentSet Then
                    bRet = True
                Else
                    bRet = False
                End If
' EP15 pct end

            Case LEADAGENT
                If node.Parent.Key = INTERMEDIARIES And Not bParentSet Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = False ' EP15 pct
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = False ' EP15 pct
                Else
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                    bFoundParent = True
                End If
                
            Case INTERMEDIARY_COMPANY
                
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = False ' EP15 pct
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = False ' EP15 pct
                
                If Not bHeaderNode Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                Else
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
                End If
                
                bFoundParent = True

            Case ADMIN_CENTRE
                
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = False ' EP15 pct
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = False ' EP15 pct
                If Not bHeaderNode Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                Else
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = False
                End If
                
                bFoundParent = True
            Case INDIVIDUAL
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_LEADAGENT).Visible = False
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_PACKAGER).Visible = False ' EP15 pct
                mnuLVIntermediariesAdd(LV_INTERMEDIARY_BROKER).Visible = False ' EP15 pct
                If bHeaderNode Then
                    mnuLVIntermediariesAdd(LV_INTERMEDIARY_INDIVIDUAL).Visible = True
                End If
                bFoundParent = True
                    
                If node.Key = INDIVIDUAL And Not bParentSet Then
                    bRet = True
                Else
                    bRet = False
                End If
            Case PRINCIPALS
               'bFoundParent = True
                
                'mnuLVIntermediariesAdd(LV_INTERMEDIARY_COMPANY).Visible = False
                'mnuLVIntermediariesAdd(LV_INTERMEDIARY_ADMIN).Visible = False
                'mnuLVIntermediaries(LV_INTERMEDIARY_MENU_FIND).Enabled = True
                'mnuLVIntermediaries(LV_INTERMEDIARY_MENU_REFRESH).Enabled = True
            
        End Select

        Set node = node.Parent ' pct - Step back one node in the tree
        sTitle = node.Text
        bParentSet = True
    Wend

    If bRet And bFoundParent Or bHeaderNode Then
        'Need to get the parent Key. But the key is made up of the parent GUID and the type
            nGuidEnd = InStr(1, sKey, node.Child.Text)
    
            If nGuidEnd > 0 Then
                clsTreeItem.SetTag Mid(sKey, 1, (nGuidEnd - 1))
            Else
                clsTreeItem.SetTag sKey
            End If

            'clsTreeItem.SetType sTitle

            g_clsMainSupport.GetIntermediary.SetTreeItem clsTreeItem

        PopupMenu mnuTopLevel(LV_INTERMEDIARY_MENU), vbPopupMenuRightButton
    End If

    HandleTreeViewPopupExtra = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
Public Sub ClearCurrentTreeitem()
    Set m_nodeCurrent = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseSecurity
' Description   : Creates a global security server and constructs a virtual
'                 table of securable items.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function InitialiseSecurity() As Boolean
    Dim bRet As Boolean
    Dim iManagement As ISecurityManager
    Dim clsSecurityManagerCS As SecurityManagerCS
    On Error GoTo Failed

    'Create the global security object.
    Set g_clsSecurityMgr = New SecurityManager

    'Inialise the security component so it can connect to a database.
    g_clsSecurityMgr.Initialise g_clsDataAccess.GetActiveConnection
        
    'Cast a management interface onto it, to add securable items (build a virtual table).
    Set iManagement = g_clsSecurityMgr
    
    With iManagement
    
        'Add promotions to the virtual table.
        .AddSecurityResource ID_PROMOTIONS, PROMOTIONS
        
    '    'Add db connections to the virtual table.
    '    .AddSecurityResource ID_DB_CONNECTIONS,DB_CONNECTIONS
        
        'Add combobox entries to the virtual table.
        .AddSecurityResource ID_COMBOBOX_ENTRIES, COMBOBOX_ENTRIES
    
        'Add parameters to the virtual table.
        '.AddSecurityResource ID_SYSTEM_PARAMETERS, SYSTEM_PARAMETERS
        .AddSecurityResource ID_GLOBAL_PARAM_FIXED, "Global Fixed Parameters" 'GLOBAL_PARAM_FIXED
        .AddSecurityResource ID_GLOBAL_PARAM_BANDED, "Global Banded Parameters" 'GLOBAL_PARAM_BANDED
    
        'Add rates and fees to the virtual table.
        '.AddSecurityResource ID_RATES_AND_FEES, RATES_AND_FEES
        
' TW 09/10/2006 EP2_7
        .AddSecurityResource ID_ADDITIONAL_BORROWING_FEES, "Additional Borrowing Fee Sets" 'ID_ADDITIONAL_BORROWING_FEES
        .AddSecurityResource ID_CREDIT_LIMIT_INCREASE_FEES, "Credit Limit Increase Fee Sets" 'ID_CREDIT_LIMIT_INCREASE_FEES
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        .AddSecurityResource ID_TRANSFER_OF_EQUITY_FEES, "Transfer of Equity Fee Sets" 'ID_CREDIT_LIMIT_INCREASE_FEES
' TW 11/12/2006 EP2_20 End

        .AddSecurityResource ID_ADMIN_FEES, "Rates Admin FeeSets" 'ADMIN_FEES
        .AddSecurityResource ID_VALUATION_FEES, "Rates Valuation FeeSets" ' VALUATION_FEES
        .AddSecurityResource ID_BASE_RATES, "Rates BaseRates" ' BASE_RATES
        .AddSecurityResource ID_BASE_RATE, "Rates BaseRateSets" ' BASE_RATE
        .AddSecurityResource ID_INCOME_MULTIPLE, "Rates Income Multiples" ' INCOME_MULTIPLE
        .AddSecurityResource ID_MP_MIG_RATE_SETS, "Rates Mortgage Product HLC Rates" ' MP_MIG_RATE_SETS
        .AddSecurityResource ID_REDEM_FEE_SETS, "Rates ERC Sets" ' ID_REDEM_FEE_SETS  'JD BMIDS982
        .AddSecurityResource ID_RENTALINCOMERATES, "Rental Income Interest Rate Sets"  'JD BMIDS765
            
        'Add batch to the virtual table.
        .AddSecurityResource ID_BATCH_SCHEDULER, BATCH_SCHEDULER
    
        'Add lenders to the virtual table.
        .AddSecurityResource ID_LENDERS, LENDERS
        
        'Add products to the virtual table.
        '.AddSecurityResource ID_PRODUCTS, PRODUCTS
        .AddSecurityResource ID_MORTGAGE_PRODUCTS, "Products MortgageProducts" ' MORTGAGE_PRODUCTS
        .AddSecurityResource ID_LIFE_COVER_RATES, "Products LifeCoverRates" 'LIFE_COVER_RATES
        .AddSecurityResource ID_BUILDINGS_AND_CONTENTS, "Products Building and Contents Products" ' BUILDINGS_AND_CONTENTS_PRODUCTS
        .AddSecurityResource ID_PAYMENT_PROTECTION_RATES, "Products Payment Protection Rates" 'PAYMENT_PROTECTION_RATES
        .AddSecurityResource ID_PAYMENT_PROTECTION_PRODUCTS, "Products Payment Protection Products" ' PAYMENT_PROTECTION_PRODUCTS
        
        'Add organisation to the virtual table.
        '.AddSecurityResource ID_ORGANISATION, ORGANISATION
        .AddSecurityResource ID_COUNTRIES, "Organisation Countries" ' COUNTRIES
        .AddSecurityResource ID_DISTRIBUTION_CHANNELS, "Organisation DistributionChannels" ' DISTRIBUTION_CHANNELS
        .AddSecurityResource ID_DEPARTMENTS, "Organisation Departments" ' DEPARTMENTS
        .AddSecurityResource ID_REGIONS, "Organisation Regions" ' REGIONS
        .AddSecurityResource ID_UNITS, "Organisation Units" ' UNITS
        .AddSecurityResource ID_COMPETENCIES, "Organisation Competencies" ' COMPETENCIES
        .AddSecurityResource ID_WORKING_HOURS, "Organisation Working Hours" ' WORKING_HOURS
        .AddSecurityResource ID_USERS, "Organisation Users" ' USERS
        
        'Add Third Parties to the virtual table.
        '.AddSecurityResource ID_NAMES_AND_ADDRESSES, NAMES_AND_ADDRESSES
        '.AddSecurityResource ID_PANEL_ADDRESS, "Third Parties Panel" ' PANEL_ADDRESS
        .AddSecurityResource ID_LEGAL_REP_ADDRESS, "Third PartiesPanel LegalRep" 'LEGAL_REP_ADDRESS
        .AddSecurityResource ID_VALUER_ADDRESS, "Third PartiesPanel Valuers" 'VALUER_ADDRESS
        .AddSecurityResource ID_LOCAL_ADDRESS, "Third Parties Other" 'LOCAL_ADDRESS
    
        'Add Income Factors to the virtual table.
        .AddSecurityResource ID_INCOME_FACTORS, INCOME_FACTORS
    
        'Add Error Messages to the virtual table.
        .AddSecurityResource ID_ERROR_MESSAGES, ERROR_MESSAGES
        
        'Add locks to the virtual table.
        '.AddSecurityResource ID_LOCK_MAINTENANCE, LOCK_MAINTENANCE
        '.AddSecurityResource ID_LOCKS_ONLINE , LOCKS_ONLINE
        .AddSecurityResource ID_LOCKS_ONLINE_APPLICATION, "LocksOnline Application" 'LOCKS_ONLINE_APPLICATION
        .AddSecurityResource ID_LOCKS_ONLINE_CUSTOMER, "LocksOnline Customer" ' LOCKS_ONLINE_CUSTOMER
        '.AddSecurityResource ID_LOCKS_OFFLINE, LOCKS_OFFLINE
        .AddSecurityResource ID_LOCKS_OFFLINE_APPLICATION, "LocksOffline Application" ' LOCKS_OFFLINE_APPLICATION
        .AddSecurityResource ID_LOCKS_OFFLINE_CUSTOMER, "LocksOffline Customer" ' LOCKS_OFFLINE_CUSTOMER
        ' ik_bm0314
        .AddSecurityResource ID_LOCKS_DOCS, "Locks Documents" ' LOCKS_DOCS
        
        'Add task management to the virtual table.
        '.AddSecurityResource ID_APPLICATION_PROCESSING, APPLICATION_PROCESSING
        '.AddSecurityResource ID_TASK_MANAGEMENT, TASK_MANAGEMENT
        .AddSecurityResource ID_TASK_MANAGEMENT_TASKS, "Task Management Tasks" 'TASK_MANAGEMENT_TASKS
        .AddSecurityResource ID_TASK_MANAGEMENT_STAGES, "Task Management Stages" 'TASK_MANAGEMENT_STAGES
        .AddSecurityResource ID_TASK_MANAGEMENT_ACTIVITIES, "Task Management Activities" ' TASK_MANAGEMENT_ACTIVITIES
        '.AddSecurityResource ID_CASE_TRACKING, "Task Management caseTracking" 'CASE_TRACKING
        .AddSecurityResource ID_BUSINESS_GROUPS, "Task Management CaseTracking BusinessGroups " 'BUSINESS_GROUPS
        
        'Add questions to the virtual table.
        .AddSecurityResource ID_ADDITIONAL_QUESTIONS, ADDITIONAL_QUESTIONS
        
        'Add conditions to the virtual table.
        .AddSecurityResource ID_CONDITIONS, CONDITIONS
        
        'Add printing templates to the virtual table.
        .AddSecurityResource ID_PRINTING_TEMPLATE, PRINTING_TEMPLATE
        .AddSecurityResource ID_PRINTING_DOCUMENT, PRINTING_DOCUMENT 'MAR45 GHun
        .AddSecurityResource ID_PRINTING_PACK, PRINTING_PACK     'MAR202 GHun
        
        'Add currencies to the virtual table.
        .AddSecurityResource ID_CURRENCIES, CURRENCIES
        
        'Add intermediaries to the virtual table.
        .AddSecurityResource ID_INTERMEDIARIES, INTERMEDIARIES
        
        'Add introducers to the virtual table. 'EP2_13 PB
        .AddSecurityResource ID_INTRODUCERS, INTRODUCERS
        
    End With

    ' SYS4149 Do any client specific security
    Set clsSecurityManagerCS = New SecurityManagerCS
    bRet = clsSecurityManagerCS.AddSecurity(iManagement)

    'Return a success indicator.
    InitialiseSecurity = bRet

    Exit Function
    
Failed:
    g_clsErrorHandling.DisplayError
    InitialiseSecurity = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : mnuLVThirdParties_Click
' Description   : Event handler for Third Party menu click.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuLVThirdParties_Click(Index As Integer)
    On Error GoTo Failed
    
    Select Case Index
        Case LV_THIRD_PARTY_MENU_ADD
            Dim vKey As Variant
            vKey = frmMain.GetSelectedTreeKey()
            
            g_clsMainSupport.HandleTreeNew vKey
        
        Case LV_THIRD_PARTY_MENU_FIND
            ' popup the find menu
            Dim tvNode As node
            Set tvNode = tvwDB.SelectedItem()
            g_clsMainSupport.HandleTreeFind
            'HandleTreeClick tvNode
        
        Case LV_THIRD_PARTY_MENU_ALL
            'BS BM0282 25/03/03
            'Reset title as the listview will be refreshed
            lblTitle(constListViewLabel).Caption = tvwDB.SelectedItem()
            
            ' Populate all records
        PopulateListView
     
     End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub SetUpTreeValuesForPrinting(vnOptionsIndex As Integer)
' Added for MAR202
On Error GoTo Failed
    
    Dim bAllow As Boolean
    Dim bParentAllowed As Boolean
    Dim TreeItem As TreeAccess
    Dim nPrintingIndex As Integer   'MAR45 GHun

    'MAR45 GHun
    'PRINTING
    bParentAllowed = False
        
    mnuFileNewPrinting.Visible = False
    Set TreeItem = New TreeAccess
    TreeItem.AddName PRINTING, PRINTING, tvwChild, vnOptionsIndex
    m_treeData.Add TreeItem, PRINTING
    nPrintingIndex = m_treeData.Count()
    'MAR45 End
    
    'PRINTING TEMPLATE
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PRINTING_TEMPLATE) And _
        g_clsMainSupport.DoesPrintTemplateExist()
    
    'Show/Hide the corresponding 'New' menu item.
    mnuFileNewPrintingTemplate.Visible = bAllow
    
    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRINTING_TEMPLATE, PRINTING_TEMPLATE, tvwChild, nPrintingIndex 'MAR45 GHun
        TreeItem.AddHeading "TemplateId", 20
        TreeItem.AddHeading "DPS TemplateId", 20
        TreeItem.AddHeading "Document Group", 20
        TreeItem.AddHeading "Template Name", 20
        TreeItem.AddHeading "Description", 30
        TreeItem.AddHeading "Minumum Role", 20
        m_treeData.Add TreeItem, PRINTING_TEMPLATE
        bParentAllowed = True   'MAR45 GHun
    End If
    
    'MAR45 GHun
    'PRINTING DOCUMENT LOCATIONS
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PRINTING_DOCUMENT)

    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRINTING_DOCUMENT, PRINTING_DOCUMENT, tvwChild, nPrintingIndex
        TreeItem.AddHeading "DPS TemplateId", 5
        TreeItem.AddHeading "Template Name", 20
        TreeItem.AddHeading "Description", 30
        TreeItem.AddHeading "File Location", 40
        m_treeData.Add TreeItem, PRINTING_DOCUMENT
        bParentAllowed = True
        'Show/Hide the corresponding menu item.
        mnuFileNewPrinting.Visible = bAllow
        mnuFileNewPrintingDocument.Visible = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewPrinting.Visible = True
            mnuFileNewPrintingDocument.Visible = bAllow
        'Else
        '    m_treeData.Remove PRINTING
        End If
    End If
    If bParentAllowed Then
        mnuFileNewPrinting.Visible = True
    End If
            
    'MAR202 GHun Start
    'bParentAllowed = False
    'MAR45 End
    
    bAllow = g_clsSecurityMgr.HasAccess(g_sSupervisorUser, ID_PRINTING_PACK)

    If bAllow Then
        Set TreeItem = New TreeAccess
        TreeItem.AddName PRINTING_PACK, PRINTING_PACK, tvwChild, nPrintingIndex
        TreeItem.AddHeading "Pack Id", 10
        TreeItem.AddHeading "Pack Name", 30
        TreeItem.AddHeading "Pack Description", 60
        m_treeData.Add TreeItem, PRINTING_PACK
        bParentAllowed = True
        'Show/Hide the corresponding menu item.
        mnuFileNewPrinting.Visible = bAllow
        mnuFileNewPrintingPack.Visible = True
    Else
        If bParentAllowed Then
            ' Only make the last child invisible, if a prior child has made the parent visible
            ' Otherwise do nothing. Trying to set ALL children invisible causes an error!
            mnuFileNewPrinting.Visible = True
            mnuFileNewPrintingPack.Visible = bAllow
        Else
            m_treeData.Remove PRINTING
        End If
    End If
    If bParentAllowed Then
        mnuFileNewPrinting.Visible = True
    End If
            
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, _
        "Error in setting up tree values for printing: " & Err.DESCRIPTION
End Sub

'Sub EditAssociation()
'    '
'    ' Get selected row
'    '
'    Dim strAssociationId As String
'
'End Sub
