VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"
Begin VB.Form frmProductMaintenance 
   Caption         =   "Product Maintenance"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frmProductMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8535
      TabIndex        =   2
      Top             =   7035
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   7215
      TabIndex        =   1
      Top             =   7035
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabProdMaint 
      Height          =   6810
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   12012
      _Version        =   393216
      Tabs            =   11
      Tab             =   6
      TabHeight       =   520
      TabCaption(0)   =   "P of L Eligibility"
      TabPicture(0)   =   "frmProductMaint.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Channel Eligibility"
      TabPicture(1)   =   "frmProductMaint.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other Fees"
      TabPicture(2)   =   "frmProductMaint.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Base Rates"
      TabPicture(3)   =   "frmProductMaint.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Administration Fees"
      TabPicture(4)   =   "frmProductMaint.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Valuation Fees"
      TabPicture(5)   =   "frmProductMaint.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Interest Rates"
      TabPicture(6)   =   "frmProductMaint.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Frame7"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Incentives"
      TabPicture(7)   =   "frmProductMaint.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Type of Application Eligibility"
      TabPicture(8)   =   "frmProductMaint.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame10"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Special Groups"
      TabPicture(9)   =   "frmProductMaint.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame11"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Employment Eligibility"
      TabPicture(10)  =   "frmProductMaint.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame12"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame12 
         Caption         =   "Eligibility"
         Height          =   5055
         Left            =   -74820
         TabIndex        =   22
         Top             =   1440
         Width           =   8655
         Begin MSGOCX.MSGEditBox txtProductCode 
            Height          =   315
            Index           =   7
            Left            =   1800
            TabIndex        =   23
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
         End
         Begin VB.Label lblPurposeOfLoan 
            Caption         =   "Product Code"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Special Groups"
         Height          =   4215
         Left            =   -74730
         TabIndex        =   21
         Top             =   1665
         Width           =   8235
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4155
         Left            =   -74700
         TabIndex        =   18
         Tag             =   "Tab1"
         Top             =   1560
         Width           =   8895
         Begin MSGOCX.MSGEditBox txtProductCode 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   19
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
         End
         Begin VB.Label lblPurposeOfLoan 
            Caption         =   "Product Code"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   20
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   5235
         Left            =   -74880
         TabIndex        =   12
         Tag             =   "Tab8"
         Top             =   1320
         Width           =   8955
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   5175
         Left            =   180
         TabIndex        =   9
         Tag             =   "Tab7"
         Top             =   1260
         Width           =   8835
         Begin VB.Frame frameFeeSetBands 
            Caption         =   "Interest Rates"
            Height          =   4695
            Left            =   240
            TabIndex        =   15
            Top             =   60
            Width           =   8235
            Begin MSGOCX.MSGEditBox txtProductCode 
               Height          =   315
               Index           =   5
               Left            =   1980
               TabIndex        =   16
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
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
            End
            Begin VB.Label lblIncentives 
               Caption         =   "Product Code"
               Height          =   255
               Index           =   2
               Left            =   660
               TabIndex        =   17
               Top             =   420
               Width           =   1215
            End
         End
         Begin MSGOCX.MSGEditBox txtProductCode 
            Height          =   315
            Index           =   2
            Left            =   2040
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
         End
         Begin VB.Label lblInterestRates 
            Caption         =   "Product Code"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   11
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   5295
         Left            =   -74940
         TabIndex        =   8
         Top             =   1260
         Width           =   9015
         Begin MSGOCX.MSGEditBox txtProductCode 
            Height          =   315
            Index           =   4
            Left            =   1920
            TabIndex        =   13
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
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
         End
         Begin VB.Label lblChanEligibility 
            Caption         =   "Product Code"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   14
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   5295
         Left            =   -74940
         TabIndex        =   7
         Tag             =   "Tab2"
         Top             =   1260
         Width           =   9015
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5295
         Left            =   -74820
         TabIndex        =   6
         Tag             =   "Tab1"
         Top             =   1260
         Width           =   8895
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5295
         Left            =   -74940
         TabIndex        =   5
         Tag             =   "Tab6"
         Top             =   1260
         Width           =   9015
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   5295
         Left            =   -74940
         TabIndex        =   4
         Tag             =   "Tab5"
         Top             =   1260
         Width           =   9015
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5295
         Left            =   -74820
         TabIndex        =   3
         Tag             =   "Tab4"
         Top             =   1260
         Width           =   8895
      End
   End
End
Attribute VB_Name = "frmProductMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Option Explicit
'Private m_colForms As Collection
'
'Private Const PURP_LOAN_PRODUCT_CODE = 0
'Private Const CHANNEL_PRODUCT_CODE = 1
'Private Const INTEREST_RATE_PRODUCT_CODE = 2
'Private Const INCENTIVE_PRODUCT_CODE = 3
'Private Const OTHER_FEE_PRODUCT_CODE = 4
'Private Const INTEREST_RATES_PRODUCT_CODE = 5
'Private Const TYPE_OF_APP_PRODUCT_CODE = 6
'Private Const EMP_ELIG_PRODUCT_CODE = 7
'Private Const OTHER_FEES_DETAILS = 0
'Private Const BASE_RATE_DETAILS = 0
'
'' Interest Rate buttons
'Private Const INT_RATE_ADD = 0
'Private Const INT_RATE_EDIT = 1
'Private Const INT_RATE_DELETE = 2
'
'' Incentives Buttons
'Private Const INCENTIVES_ADD = 0
'Private Const INCENTIVES_EDIT = 1
'Private Const INCENTIVES_DELETE = 2
'Private m_bIsEdit As Boolean
'Private m_sProductCode As String
'Private m_sProductStartDate As String
'Private m_vOrgID As Variant
'Public Sub SetProductCode(sProductCode As String, sProductStartDate As String, vOrgID As Variant)
'    m_sProductCode = sProductCode
'    m_sProductStartDate = sProductStartDate
'    m_vOrgID = vOrgID
'End Sub
'
'Public Function IsEdit() As Boolean
'    IsEdit = m_bIsEdit
'End Function
'Public Sub SetIsEdit(Optional bEdit As Boolean = True)
'    m_bIsEdit = bEdit
'End Sub
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdAdminFeeSelect_Click()
'    m_colForms(maintTabAdminFees).SelectAdminFee
'End Sub
'
'Private Sub cmdBaseRateCurrentSet_Click()
'    m_colForms(maintTabBaseRates).CurrentSet
'End Sub
'
'Private Sub cmdBaseRateSelect_Click()
'    m_colForms(maintTabBaseRates).SelectBaseRate
'End Sub
'
'Private Sub cmdIncentivesAdd_Click()
'    m_colForms(maintTabIncentives).Add
'End Sub
'
'Private Sub cmdIncentivesDelete_Click()
'    m_colForms(maintTabIncentives).Delete
'End Sub
'
'Private Sub cmdIncentivesEdit_Click()
'    m_colForms(maintTabIncentives).Edit
'End Sub
'
'Private Sub cmdInterestRateAdd_Click()
'    m_colForms(maintTabInterestRates).Add
'End Sub
'
'Private Sub cmdInterestRateDelete_Click()
'    m_colForms(maintTabInterestRates).Delete
'End Sub
'
'Private Sub cmdInterestRateEdit_Click()
'    m_colForms(maintTabInterestRates).Edit
'End Sub
'
'Private Sub cmdValuationFeeSelect_Click()
'    m_colForms(maintTabValuationFees).SelectValuationFee
'End Sub

'Private Sub cmdInterestRates_Click(Index As Integer)
'    Select Case Index
'    Case INT_RATE_ADD
'        frmMaintIntRates.Show vbModal, Me
'    Case INT_RATE_EDIT
'        frmMaintIntRates.Show vbModal, Me
'    Case INT_RATE_DELETE
'        MsgBox "Are you sure?", vbYesNo
'    End Select
'End Sub
'Private Sub lvAdminFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    cmdAdminFeeCurrentSet.Enabled = True
'    cmdAdminFeeSelect.Enabled = True
'End Sub
'
'Private Sub lvBaseRates_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    cmdBaseRateCurrentSet.Enabled = True
'    cmdBaseRateSelect.Enabled = True
'End Sub
'
'Private Sub lvIncentives_DblClick()
'    m_colForms(maintTabIncentives).Edit
'End Sub
'Private Sub lvIncentives_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    cmdIncentivesEdit.Enabled = True
'    cmdIncentivesDelete.Enabled = True
'End Sub
'Private Sub lvInterestRates_DblClick()
'    m_colForms(maintTabInterestRates).Edit
'End Sub
'
'Private Sub lvInterestRates_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    cmdInterestRateEdit.Enabled = True
'    cmdInterestRateDelete.Enabled = True
'End Sub
'Private Sub lvValuationFees_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    cmdValuationFeeCurrentSet.Enabled = True
'    cmdValuationFeeSelect.Enabled = True
'End Sub
'Private Sub MSGSwapOtherFees_ExtraClick(nIndex As Integer)
'    Select Case nIndex
'        Case OTHER_FEES_DETAILS
'            MsgBox ("Other Fees Details pressed")
'        End Select
'End Sub
'
'Private Sub MSGSwapBaseRate_ExtraClick(nIndex As Integer)
'    Select Case nIndex
'        Case BASE_RATE_DETAILS
'            MsgBox ("Base Rate Details pressed")
'        End Select
'End Sub
'
''Private Sub MSGSwapPurpOfLoan_SecondItemClicked(colRow As Collection, objExtra As Object)
''    m_colForms(maintTabPurposeOfLoan).PurposeOfLoanClicked colRow, objExtra
''End Sub
'
'Private Sub SSTabProdMaint_Click(PreviousTab As Integer)
'    g_clsFormProcessing.SetTabstops Me
'    InitialiseFields
'End Sub

'Private Sub Form_Activate()
'    On Error GoTo Failed
'    InitialiseFields
'    Exit Sub
'Failed:
'    g_clsErrAssist.DisplayError
'End Sub
'Private Sub Form_Load()
'    On Error GoTo Failed
'    Dim bRet As Boolean
'    SSTabProdMaint_Click 0
'    SetupListViewFields
'
'    InitialiseForms
'
'    g_clsFormProcessing.SetParent Me
'    InitialiseFields
'    PopulateScreenFields
'    SetActiveTab maintTabPurposeOfLoan - 1
'
'    EndWaitCursor
'    Exit Sub
'Failed:
'    g_clsErrAssist.DisplayError
'    g_clsFormProcessing.CancelForm Me
'End Sub
'Private Sub InitialiseForms()
'    On Error GoTo Failed
'    Dim nTabCount As Integer
'    Dim nThisTab As Integer
'
'    Set m_colForms = New Collection
'
'    m_colForms.Add New MortProdPurpOfLoan
'    m_colForms.Add New MortProdChannelElig
'    m_colForms.Add New MortProdOtherFee
'    m_colForms.Add New MortProdBaseRates
'    m_colForms.Add New MortProdAdminFees
'    m_colForms.Add New MortProdValuationFee
'    m_colForms.Add New MortProdIntRates
'    m_colForms.Add New MortProdIncentive
'    m_colForms.Add New MortProdTypeofAppElig
'    m_colForms.Add New MortProdSpecialGroup
'    m_colForms.Add New MortProdEmpElig
'
'    m_colForms(maintTabPurposeOfLoan).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabPurposeOfLoan).Initialise m_bIsEdit
'
'    m_colForms(maintTabChannelEligibility).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabChannelEligibility).Initialise m_bIsEdit
'
'    m_colForms(maintTabOtherFees).SetProductCode m_sProductCode, m_sProductStartDate, m_vOrgID
'    m_colForms(maintTabOtherFees).Initialise m_bIsEdit
'
'    m_colForms(maintTabBaseRates).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabBaseRates).Initialise m_bIsEdit
'
'    m_colForms(maintTabAdminFees).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabAdminFees).Initialise m_bIsEdit
'
'    m_colForms(maintTabValuationFees).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabValuationFees).Initialise m_bIsEdit
'
'    m_colForms(maintTabIncentives).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabIncentives).Initialise m_bIsEdit
'
'    m_colForms(maintTabInterestRates).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabInterestRates).Initialise m_bIsEdit
'
'    m_colForms(maintTabTypeOfAppElig).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabTypeOfAppElig).Initialise m_bIsEdit
'
'    m_colForms(maintTabSpecialGroup).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabSpecialGroup).Initialise m_bIsEdit
'
'    m_colForms(maintTabEmpElig).SetProductCode m_sProductCode, m_sProductStartDate
'    m_colForms(maintTabEmpElig).Initialise m_bIsEdit
'
'    Exit Sub
'Failed:
'    g_clsErrAssist.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'Public Sub PopulateScreenFields()
'    On Error GoTo Failed
'    Dim nTabCount As Integer
'    Dim nThisTab As Integer
'
'    txtProductCode(PURP_LOAN_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(CHANNEL_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(INTEREST_RATE_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(INCENTIVE_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(OTHER_FEE_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(INTEREST_RATES_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(TYPE_OF_APP_PRODUCT_CODE).Text = m_sProductCode
'    txtProductCode(EMP_ELIG_PRODUCT_CODE).Text = m_sProductCode
'
'    nTabCount = m_colForms.Count
'
'    For nThisTab = 1 To nTabCount
'        m_colForms(nThisTab).SetScreenFields
'    Next
'
'    Exit Sub
'Failed:
'    g_clsErrAssist.RaiseError Err.Number, "PopulateScreenFields" + Err.DESCRIPTION
'End Sub
'Private Sub InitialiseFields()
'    On Error GoTo Failed
'
'    cmdBaseRateCurrentSet.Enabled = False
'    cmdBaseRateSelect.Enabled = False
'
'    cmdAdminFeeSelect.Enabled = False
'    cmdValuationFeeSelect.Enabled = False
'
'    cmdIncentivesDelete.Enabled = False
'    cmdIncentivesEdit.Enabled = False
'
'    cmdInterestRateEdit.Enabled = False
'    cmdInterestRateDelete.Enabled = False
'
'    Exit Sub
'Failed:
'    g_clsErrAssist.RaiseError Err.Number, "InitialiseFields:" + Err.DESCRIPTION
'End Sub
'
'Private Sub cmdOK_Click()
'    On Error GoTo Failed
'    Dim bRet As Boolean
'
'    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
'
'    If bRet = True Then
'        bRet = ValidateScreenData()
'
'        If bRet = True Then
'            SaveScreenData
'            Unload Me
'        End If
'    End If
'    Exit Sub
'Failed:
'    g_clsErrAssist.DisplayError
'End Sub
'Private Function ValidateScreenData() As Boolean
'    Dim bRet As Boolean
'    Dim nCount As Integer
'    Dim nThisForm As Integer
'
'    bRet = True
'
'    nThisForm = 1
'    nCount = m_colForms.Count
'
'    While nThisForm <= nCount And bRet = True
'        bRet = m_colForms(nThisForm).ValidateScreenData()
'
'        If bRet = True Then
'            nThisForm = nThisForm + 1
'        Else
'            SSTabProdMaint.Tab = nThisForm - 1
'        End If
'    Wend
'
'    ValidateScreenData = bRet
'
'End Function
'Public Sub SaveScreenData()
'    On Error GoTo Failed
'    Dim nTabCount As Integer
'    Dim nThisTab As Integer
'
'    nTabCount = 1 'm_colForms.Count
'
'    nThisTab = 1
'
'    m_colForms(maintTabPurposeOfLoan).SaveScreenData
'    m_colForms(maintTabChannelEligibility).SaveScreenData
'    m_colForms(maintTabOtherFees).SaveScreenData
'    m_colForms(maintTabBaseRates).SaveScreenData
'    m_colForms(maintTabAdminFees).SaveScreenData
'    m_colForms(maintTabValuationFees).SaveScreenData
'    m_colForms(maintTabTypeOfAppElig).SaveScreenData
'    m_colForms(maintTabSpecialGroup).SaveScreenData
'    m_colForms(maintTabEmpElig).SaveScreenData
'
'    Exit Sub
'Failed:
'    g_clsErrAssist.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'
'Friend Sub SetActiveTab(activeTab As ProdMaintTabs)
'    SSTabProdMaint.Tab = activeTab
'
'End Sub
'' Need to setup the tabs ListView fields.
'Private Sub SetupListViewFields()
'    SetupPurposeOfLoan
'    SetupTypeOfApp
'    SetupChannelElig
'    SetupOtherFees
'    SetupBaseRates
'    SetupAdminFees
'    SetupValuationFees
'    SetupEmpElig
''    SetupIncentives
'End Sub
'Private Sub SetupValuationFees()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
''        colListLine.Add GetTypeOfValuationText()
''    colListLine.Add GetLocationText()
''    colListLine.Add Me.GetMaximumValue()
''    colListLine.Add GetAmount()
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Fee Set No"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Start Date"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Type of Valuation"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Location"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 10
'    lvHeaders.sName = "Max Value"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Fee Amount"
'    headers.Add lvHeaders
'
'    lvValuationFees.AddHeadings headers
'
'End Sub
'Private Sub SetupAdminFees()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Set No."
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Start Date"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Type of App"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Location"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Fee Amount"
'
'    headers.Add lvHeaders
'
'    lvAdminFees.AddHeadings headers
'
'End Sub
'Private Sub SetupBaseRates()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Set Number"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Start Date"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Max. Loan Amt."
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 25
'    lvHeaders.sName = "Rate"
'    headers.Add lvHeaders
'
'    lvBaseRates.AddHeadings headers
'End Sub
'Private Sub SetupTypeOfApp()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Purpose of Loan"
'    headers.Add lvHeaders
'    MSGSwapTypeOfAppEligibility.SetFirstColumnHeaders headers
'
'    Set headers = New Collection
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Eligible Type of Application"
'    headers.Add lvHeaders
'    MSGSwapTypeOfAppEligibility.SetSecondColumnHeaders headers
'End Sub
'
'Private Sub SetupPurposeOfLoan()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    ' Purpose of Loan
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Purpose of Loan"
'    headers.Add lvHeaders
'    MSGSwapPurpOfLoan.SetFirstColumnHeaders headers
'
'    Set headers = New Collection
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Purpose of Loan Eligibility"
'    headers.Add lvHeaders
'    MSGSwapPurpOfLoan.SetSecondColumnHeaders headers
'
'    ' Type of Buyer
'    Set headers = New Collection
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Available Type of Buyer"
'    headers.Add lvHeaders
'    MSGSwapTypeOfBuyer.SetFirstColumnHeaders headers
'
'    Set headers = New Collection
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Selected Type of Buyer"
'    headers.Add lvHeaders
'    MSGSwapTypeOfBuyer.SetSecondColumnHeaders headers
'
'
'End Sub
'
'Private Sub SetupChannelElig()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Sales Channel List"
'    headers.Add lvHeaders
'    MSGSwapChannel.SetFirstColumnHeaders headers
'
'    Set headers = New Collection
'    lvHeaders.sName = "Sales Channel Eligibility"
'
'    headers.Add lvHeaders
'    MSGSwapChannel.SetSecondColumnHeaders headers
'
'End Sub
'Private Sub SetupEmpElig()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'    lvHeaders.nWidth = 100
'    lvHeaders.sName = "Employment Eligibility"
'    headers.Add lvHeaders
'    MSGSwapEmpElig.SetFirstColumnHeaders headers
'
'    Set headers = New Collection
'    lvHeaders.sName = "Eligible Product Status"
'
'    headers.Add lvHeaders
'    MSGSwapEmpElig.SetSecondColumnHeaders headers
'End Sub
'
'Private Sub SetupOtherFees()
'    Dim headers As New Collection
'    Dim lvHeaders As listViewAccess
'    Dim line As Collection
'
'
'    lvHeaders.nWidth = 50
'    lvHeaders.sName = "Fee Name"
'    headers.Add lvHeaders
'
'    lvHeaders.nWidth = 50
'    lvHeaders.sName = "Start Date"
'    headers.Add lvHeaders
'
'
'    MSGSwapOtherFees.SetFirstColumnHeaders headers
'    MSGSwapOtherFees.SetSecondColumnHeaders headers
'
'    MSGSwapOtherFees.SetFirstTitle "Available Other Fees"
'    MSGSwapOtherFees.SetSecondTitle "Selected Other Fees"
'End Sub
'
' DJP Admin
Private Sub lvBaseRates_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
