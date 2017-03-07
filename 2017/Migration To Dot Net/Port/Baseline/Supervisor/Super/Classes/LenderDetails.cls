VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderDetails
' Description   : Manages all form processing for the main Lender details tab on the Lender Details
'                 form.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port, use Supervisor version of SQLAssist instead
' DJP       21/10/01    SYS2831 - change Lender to support client variants
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       03/12/01    SYS2912 SQL Server locking problem, tidy up.
' STB       07/01/01    SYS3581 - SetAddState allocates a keys collection to enable promotions.
' MC        02/06/2004  BMIDS763    BMIDS763 - CC075 - TT Fee Field removed, Product Switch Fee Check box added
'                       to lender details tab sheet
' JD        30/03/2005  BMIDS982 Changed screen text from MIG to HLC
' TW        09/10/2006  EP2_7 Added code for Additional Borrowing Fee and Credit Limit Increase Fee

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Text Edit constants
Private Const LENDER_CODE = 0
Private Const LENDER_NAME = 1
Private Const START_DATE = 2
Private Const END_DATE = 3
Private Const MAX_LOAN_COMPONENTS = 4
Private Const MIG_START_LTV = 5
Private Const IGNORE_MIG_PREMIUM = 6
Private Const MIN_MIG_PREMIUM = 7
Private Const TT_FEE = 8
Private Const SEALING_FEE = 9
Private Const DEEDS_RELEASE_FEE = 10
Private Const ARRANGEMENT_FEE_PERCENT = 11
Private Const ARRANGEMENT_FEE_PERCENT_MAX = 12
Private Const ARRANGEMENT_FEE = 13
Private Const MIG_THRESHOLD = 14

Private Const NON_GUARANTEED_INCOME_MULT = 16
Private Const GUARANTEED_INCOME_MULT = 17
Private Const SINGLE_INCOME_MULTIPLE = 18
Private Const SECOND_EARNER_MULTIPLE = 19
Private Const HIGHEST_EARNER_MULTIPLE = 20
Private Const JOINT_INCOME_MULTIPLE = 21

' CheckBoxes
Private Const ADD_PORTING_FEE = 0
Private Const ADD_ADMIN_FEE = 1
Private Const ADD_VALUATION_FEE = 2
Private Const ADD_REINSPECTION_FEE = 3
Private Const ADD_ARRANGEMENT_FEE = 4
Private Const ADD_MIG_FEE = 5
Private Const ADD_TT_FEE = 6
Private Const ADD_SEALING_FEE = 7
Private Const ADD_DEEDS_RELEASE_FEE = 8
Private Const ADD_PRODUCTSWITCH_FEE = 10
' TW 09/10/2006 EP2_7
Private Const ADD_ADDITIONALBORROWING_FEE = 11
Private Const ADD_CREDITLIMITINCREASE_FEE = 12
' TW 09/10/2006 EP2_7 End

' Labels
Private Const ORGANISATION_TYPE = 1

' Private data
Private m_sMIGRateKey As String
Private m_clsLenderAddressDir As MortLenderDirTable
Private m_clsNameAndAddressDir As NameAndAddressDirTable
Private m_clsLenderDetailsTable As MortgageLendersTable

Private m_bIsEdit As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is instantiated. Create module data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsLenderDetailsTable = New MortgageLendersTable
    Set m_clsLenderAddressDir = New MortLenderDirTable
    Set m_clsNameAndAddressDir = New NameAndAddressDirTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Called when this class is instantiated. Create module data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    On Error GoTo Failed
    Dim nRecordCount As Integer
    Dim sFunctionName As String
    Dim vDirectoryGUID As Variant
    Dim clsTableAccess As TableAccess
    
    sFunctionName = "LenderDetails.SetEditState"
    
    On Error GoTo Failed
    Set clsTableAccess = m_clsLenderDetailsTable
    
    ' Get Lender table data
    clsTableAccess.GetTableData
    clsTableAccess.ValidateData
    nRecordCount = clsTableAccess.RecordCount()
    
    If nRecordCount > 0 Then
        Dim colValues As New Collection
        Set colValues = clsTableAccess.GetKeyMatchValues()

        ' Now Lender Address Dir data
        Set clsTableAccess = m_clsLenderAddressDir
        clsTableAccess.SetKeyMatchValues colValues
        clsTableAccess.GetTableData
        clsTableAccess.ValidateData
        
        nRecordCount = clsTableAccess.RecordCount()
        
        If nRecordCount > 0 Then
            ' Need the Directory GUID
            ' DJP SQL Server port
            vDirectoryGUID = m_clsLenderAddressDir.GetDirectoryGUID()

            Set clsTableAccess = m_clsNameAndAddressDir
            Set colValues = New Collection
            ' DJP SQL Server port
            colValues.Add vDirectoryGUID
            clsTableAccess.SetKeyMatchValues colValues

            clsTableAccess.GetTableData

            clsTableAccess.ValidateData
            nRecordCount = clsTableAccess.RecordCount()

            If nRecordCount = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, sFunctionName & " Incorrect number Name and Address records returned: " & nRecordCount
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName & " Incorrect number of Lender Name and Address Directory records found: " & nRecordCount
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName & " Incorrect number of Lender records returned: " & nRecordCount
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Called when this class is instantiated. Create module data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    On Error GoTo Failed
    Dim vOrgID As Variant
    Dim sNextNumber As String
    Dim colKeys As Collection
    Dim vDirectoryGUID As Variant
    
    ' First, create a record on the Lender table
    g_clsFormProcessing.CreateNewRecord m_clsLenderDetailsTable

    ' Next, the Lender Address table
    g_clsFormProcessing.CreateNewRecord m_clsLenderAddressDir

    ' And Name and Address Directory
    g_clsFormProcessing.CreateNewRecord m_clsNameAndAddressDir
    vDirectoryGUID = m_clsNameAndAddressDir.SetDirectoryGUID()

    ' Set the Directory GUID for the Lender Address table.
    m_clsLenderAddressDir.SetDirectoryGUID vDirectoryGUID
    
    ' Now set the Address and ContactDetails GUIDs
    m_clsNameAndAddressDir.SetAddressGUID
    m_clsNameAndAddressDir.SetContactDetailsGUID
    
    ' The set the OrganisationID
    vOrgID = m_clsLenderDetailsTable.SetOrganisationID
    m_clsLenderAddressDir.SetOrganisationID vOrgID
    
    'Create a keys collection to assign against the table object.
    Set colKeys = New Collection
    
    'Add the OrganisationID (GUID) into the collection.
    colKeys.Add m_clsLenderAddressDir.GetOrganisationID()
    
    'Assign the keys collection to the table now - otherwise promting will fail.
    TableAccess(m_clsLenderDetailsTable).SetKeyMatchValues colKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to retrieve all data from the  database, setup fields etc.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    
    m_bIsEdit = bIsEdit
    
    ' Only do this if we are editing a record.
    If bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    g_clsFormProcessing.PopulateCombo "OrganisationType", frmLenderDetails.cboOrganisationType
    g_clsFormProcessing.PopulateCombo "ThirdPartyType", frmLenderDetails.cboAddressType
    
    ' Default radio buttons
    frmLenderDetails.optAddToLoan(OPT_YES).Value = True
    frmLenderDetails.optAddToLoan(OPT_YES).Enabled = False
    frmLenderDetails.optAddToLoan(OPT_NO).Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLenderTable
' Description   : Returns the MortgageLender table class. This can be used for any derived classes
'                 to do their own read/write operations.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLenderTable() As MortgageLendersTable
    On Error GoTo Failed
        
    If m_clsLenderDetailsTable Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "LenderDetails - Lender table is empty"
    End If
    
    Set GetLenderTable = m_clsLenderDetailsTable
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all screen fields from the data retrieved from the database. This
'                 method will be called externally to this class to populate the LenderDetails tab
'                 of the Lender Details form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    Dim vval As Variant

    ' NameAndAddressDirectory
    vval = m_clsNameAndAddressDir.GetNameAndAddressType
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAddressType, vval, SET_CONTROL_VALUE
    
    frmLenderDetails.txtLenderDetails(LENDER_CODE).Text = m_clsLenderDetailsTable.GetLenderCode()
    frmLenderDetails.txtLenderDetails(LENDER_NAME).Text = m_clsLenderDetailsTable.GetLenderName()
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(START_DATE), m_clsLenderDetailsTable.GetStartDate(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(END_DATE), m_clsLenderDetailsTable.GetEndDate(), SET_CONTROL_VALUE
    frmLenderDetails.txtLenderDetails(MAX_LOAN_COMPONENTS).Text = m_clsLenderDetailsTable.GetMaxNoLoans()
    frmLenderDetails.txtLenderDetails(MIG_THRESHOLD).Text = m_clsLenderDetailsTable.GetMIGThreshold()
    frmLenderDetails.txtLenderDetails(MIG_START_LTV).Text = m_clsLenderDetailsTable.GetMIGStartLTV()
    frmLenderDetails.txtLenderDetails(IGNORE_MIG_PREMIUM).Text = m_clsLenderDetailsTable.GetIgnoreMIGPremium()
    
    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddIPT(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), m_clsLenderDetailsTable.GetAddIPTToMIGPremium(), SET_CONTROL_VALUE
    
    frmLenderDetails.txtLenderDetails(TT_FEE).Text = m_clsLenderDetailsTable.GetTTFee()
    frmLenderDetails.txtLenderDetails(SEALING_FEE).Text = m_clsLenderDetailsTable.GetSealingFee()
    frmLenderDetails.txtLenderDetails(DEEDS_RELEASE_FEE).Text = m_clsLenderDetailsTable.GetDeedsReleaseFee()
    frmLenderDetails.txtLenderDetails(MIN_MIG_PREMIUM).Text = m_clsLenderDetailsTable.GetMinMIGPremium()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE).Text = m_clsLenderDetailsTable.GetArrangementFeeAmount()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT).Text = m_clsLenderDetailsTable.GetArrangementFeePercent()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX).Text = m_clsLenderDetailsTable.GetArrangementFeePercentMax()

    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_PORTING_FEE), m_clsLenderDetailsTable.GetAllowPortingFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ADMIN_FEE), m_clsLenderDetailsTable.GetAllowAdminFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_VALUATION_FEE), m_clsLenderDetailsTable.GetAllowValuationFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_REINSPECTION_FEE), m_clsLenderDetailsTable.GetAllowReinspectionFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ARRANGEMENT_FEE), m_clsLenderDetailsTable.GetAllowArrangementFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_MIG_FEE), m_clsLenderDetailsTable.GetAllowMIGFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_TT_FEE), m_clsLenderDetailsTable.GetAllowTTFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_SEALING_FEE), m_clsLenderDetailsTable.GetAllowSealingFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_DEEDS_RELEASE_FEE), m_clsLenderDetailsTable.GetAllowDeedsReleaseFeeAdded(), SET_CONTROL_VALUE
    '*=[MC]BMIDS763
    Call g_clsFormProcessing.HandleCheckBox(frmLenderDetails.chkLenderDetails(ADD_PRODUCTSWITCH_FEE), m_clsLenderDetailsTable.GetProductSwitchFeeAdded(), SET_CONTROL_VALUE)
    '*=SECTION END

' TW 09/10/2006 EP2_7
    Call g_clsFormProcessing.HandleCheckBox(frmLenderDetails.chkLenderDetails(ADD_ADDITIONALBORROWING_FEE), m_clsLenderDetailsTable.GetAllowFurtherLendingFeeAdded(), SET_CONTROL_VALUE)
    Call g_clsFormProcessing.HandleCheckBox(frmLenderDetails.chkLenderDetails(ADD_CREDITLIMITINCREASE_FEE), m_clsLenderDetailsTable.GetCreditLimitIncreaseFeeAdded(), SET_CONTROL_VALUE)
' TW 09/10/2006 EP2_7 End
    
    vval = m_clsNameAndAddressDir.GetOrganisationType(frmLenderDetails.cboOrganisationType, frmLenderDetails.lblLenderDetails(ORGANISATION_TYPE))
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboOrganisationType, vval, SET_CONTROL_VALUE
    
    ' Income Multiple changes
    frmLenderDetails.txtLenderDetails(NON_GUARANTEED_INCOME_MULT).Text = m_clsLenderDetailsTable.GetNonGuaranteedIncomeMultiple(frmLenderDetails.txtLenderDetails(NON_GUARANTEED_INCOME_MULT), frmLenderDetails.lblLenderDetails(NON_GUARANTEED_INCOME_MULT))
    frmLenderDetails.txtLenderDetails(GUARANTEED_INCOME_MULT).Text = m_clsLenderDetailsTable.GetGuaranteedIncomeMultiple(frmLenderDetails.txtLenderDetails(GUARANTEED_INCOME_MULT), frmLenderDetails.lblLenderDetails(GUARANTEED_INCOME_MULT))
    frmLenderDetails.txtLenderDetails(SINGLE_INCOME_MULTIPLE).Text = m_clsLenderDetailsTable.GetSingleIncomeMultiple(frmLenderDetails.txtLenderDetails(SINGLE_INCOME_MULTIPLE), frmLenderDetails.lblLenderDetails(SINGLE_INCOME_MULTIPLE))
    frmLenderDetails.txtLenderDetails(SECOND_EARNER_MULTIPLE).Text = m_clsLenderDetailsTable.GetSecondEarnerIncomeMultiple(frmLenderDetails.txtLenderDetails(SECOND_EARNER_MULTIPLE), frmLenderDetails.lblLenderDetails(SECOND_EARNER_MULTIPLE))
    frmLenderDetails.txtLenderDetails(HIGHEST_EARNER_MULTIPLE).Text = m_clsLenderDetailsTable.GetHighestEarnerIncomeMultiple(frmLenderDetails.txtLenderDetails(HIGHEST_EARNER_MULTIPLE), frmLenderDetails.lblLenderDetails(HIGHEST_EARNER_MULTIPLE))
    frmLenderDetails.txtLenderDetails(JOINT_INCOME_MULTIPLE).Text = m_clsLenderDetailsTable.GetJointIncomeMultiple(frmLenderDetails.txtLenderDetails(JOINT_INCOME_MULTIPLE), frmLenderDetails.lblLenderDetails(JOINT_INCOME_MULTIPLE))
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all screen fields to the database. This method will
'                 be called externally to this class to save LenderDetails tab
'                 of the Lender Details form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim sLenderName As String
    
    sLenderName = frmLenderDetails.txtLenderDetails(LENDER_NAME).Text
    ' Lender Name and Address Dir
    m_clsLenderAddressDir.SetMainMortgageLenderInd "1"
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAddressType, vTmp, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetNameAndAddressType vTmp


    m_clsLenderDetailsTable.SetLenderName sLenderName
    m_clsNameAndAddressDir.SetCompanyName sLenderName
    m_clsLenderDetailsTable.SetLenderCode frmLenderDetails.txtLenderDetails(LENDER_CODE).Text
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(START_DATE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetStartDate vTmp
    m_clsNameAndAddressDir.SetActiveFrom vTmp
    
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetEndDate vTmp
    m_clsNameAndAddressDir.SetActiveTo vTmp
    
    m_clsLenderDetailsTable.SetMaxNoLoans frmLenderDetails.txtLenderDetails(MAX_LOAN_COMPONENTS).Text
    m_clsLenderDetailsTable.SetMIGStartLTV frmLenderDetails.txtLenderDetails(MIG_START_LTV).Text
    m_clsLenderDetailsTable.SetMIGThreshold frmLenderDetails.txtLenderDetails(MIG_THRESHOLD).Text
    
    m_clsLenderDetailsTable.SetIgnoreMIGPremium frmLenderDetails.txtLenderDetails(IGNORE_MIG_PREMIUM).Text
    
    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddIPT(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAddIPTToMIGPremium vTmp
    
    m_clsLenderDetailsTable.SetTTFee frmLenderDetails.txtLenderDetails(TT_FEE).Text
    m_clsLenderDetailsTable.SetSealingFee frmLenderDetails.txtLenderDetails(SEALING_FEE).Text
    m_clsLenderDetailsTable.SetDeedsReleaseFee frmLenderDetails.txtLenderDetails(DEEDS_RELEASE_FEE).Text
    m_clsLenderDetailsTable.SetMinMIGPremium frmLenderDetails.txtLenderDetails(MIN_MIG_PREMIUM).Text

    m_clsLenderDetailsTable.SetArrangementFeeAmount frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE).Text
    m_clsLenderDetailsTable.SetArrangementFeePercent frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT).Text
    m_clsLenderDetailsTable.SetArrangementFeePercentMax frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX).Text

    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddToLoan(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAddCostsToLoan (vTmp)
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_PORTING_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowPortingFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ADMIN_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowAdminFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_VALUATION_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowValuationFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ARRANGEMENT_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowArrangementFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_REINSPECTION_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowReinspectionFeeAdded vTmp
    '*=[MC]BMIDS763 - CC075 FEESETS
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_PRODUCTSWITCH_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetProductSwitchFeeAdded vTmp
    
    '*=SECTION END
' TW 09/10/2006 EP2_7
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ADDITIONALBORROWING_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowFurtherLendingFeeAdded vTmp
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_CREDITLIMITINCREASE_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetCreditLimitIncreaseFeeAdded vTmp
' TW 09/10/2006 EP2_7 End
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_MIG_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowMIGFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_TT_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowTTFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_SEALING_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowSealingFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_DEEDS_RELEASE_FEE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowDeedsReleaseFeeAdded vTmp
    
    ' Here...
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboOrganisationType, vTmp, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetOrganisationType vTmp
    
    ' Income Multiple changes
    If frmLenderDetails.txtLenderDetails(NON_GUARANTEED_INCOME_MULT).Enabled Then
        m_clsLenderDetailsTable.SetNonGuaranteedIncomeMultiple frmLenderDetails.txtLenderDetails(NON_GUARANTEED_INCOME_MULT).Text
    End If
    
    If frmLenderDetails.txtLenderDetails(GUARANTEED_INCOME_MULT).Enabled Then
        m_clsLenderDetailsTable.SetGuaranteedIncomeMultiple frmLenderDetails.txtLenderDetails(GUARANTEED_INCOME_MULT).Text
    End If
    
    If frmLenderDetails.txtLenderDetails(SINGLE_INCOME_MULTIPLE).Enabled Then
        m_clsLenderDetailsTable.SetSingleIncomeMultiple frmLenderDetails.txtLenderDetails(SINGLE_INCOME_MULTIPLE).Text
    End If
    
    If frmLenderDetails.txtLenderDetails(SECOND_EARNER_MULTIPLE).Enabled Then
        m_clsLenderDetailsTable.SetSecondEarnerIncomeMultiple frmLenderDetails.txtLenderDetails(SECOND_EARNER_MULTIPLE).Text
    End If
    
    If frmLenderDetails.txtLenderDetails(HIGHEST_EARNER_MULTIPLE).Enabled Then
        m_clsLenderDetailsTable.SetHighestEarnerIncomeMultiple frmLenderDetails.txtLenderDetails(HIGHEST_EARNER_MULTIPLE).Text
    End If
    
    If frmLenderDetails.txtLenderDetails(JOINT_INCOME_MULTIPLE).Enabled Then
        m_clsLenderDetailsTable.SetJointIncomeMultiple frmLenderDetails.txtLenderDetails(JOINT_INCOME_MULTIPLE).Text
    End If
    
    SaveRateSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveRateSet()
    On Error GoTo Failed
        
    If Len(m_sMIGRateKey) > 0 Then
        m_clsLenderDetailsTable.SetRateSet m_sMIGRateKey
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates all tables used by this tab. This method will be called externally to
'                 this class to commit all data that has been added to any of the table classes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed

    TableAccess(m_clsNameAndAddressDir).Update
    TableAccess(m_clsLenderDetailsTable).Update
    TableAccess(m_clsLenderAddressDir).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub InitTab()

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates that all data entered on the Lender Details tab is valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    
    ' First, the End Date must not be earlier than the startdate
    ValidateDates
    
    ' Arrangement Fees
    ValidateArrangementFee
    
    ' Max Loan components
    ValidateLoanComponents
    
    ' MIG Threshold
    ValidateMIGThreshold
               
    ValidateScreenData = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateLoanComponents
' Description   : Validates that a loan component has been entered and is at least 1.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateLoanComponents()
   
    On Error GoTo Failed
    Dim sMax
    sMax = frmLenderDetails.txtLenderDetails(MAX_LOAN_COMPONENTS).Text
    
    If Len(sMax) > 0 Then
        If CInt(sMax) < 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Loan Components must be 1 or above"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateArrangementFee
' Description   : Fee and Fee Percent cannot be entered at the same time, and Fee and Perent max
'                 cannot be entered at the same time.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateArrangementFee()
    On Error GoTo Failed
    Dim sFee As String
    Dim sFeePercent As String
    Dim sFeePercentMax As String
    
    sFee = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE).Text
    sFeePercent = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT).Text
    sFeePercentMax = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX).Text
    
    If Len(sFee) > 0 And Len(sFeePercent) > 0 Then
        g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE)
        g_clsErrorHandling.RaiseError errGeneralError, "Only one of Arrangement Fee and Arrangement Fee Percent can be entered"
    End If
    
    If Len(sFee) > 0 And Len(sFeePercentMax) > 0 Then
        g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX)
        g_clsErrorHandling.RaiseError errGeneralError, "Percentage Max cannot be entered when Arrangement fee is entered"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateDates
' Description   : Validates start and end dates - if both have been entererd, the End Date must
'                 be later or equal to the start date.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateDates()
    On Error GoTo Failed
    Dim endDate As Date
    Dim startDate As Date
    Dim bEndValid As Boolean
    Dim bStartValid As Boolean
    
    bStartValid = frmLenderDetails.txtLenderDetails(START_DATE).ValidateData(False)
    bEndValid = frmLenderDetails.txtLenderDetails(END_DATE).ValidateData(False)
    
    If bStartValid = True And bEndValid = True Then
        startDate = CDate(frmLenderDetails.txtLenderDetails(START_DATE).Text)
        endDate = CDate(frmLenderDetails.txtLenderDetails(END_DATE).Text)
    
        If endDate < startDate Then
            g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(END_DATE)
            g_clsErrorHandling.RaiseError errGeneralError, "End Date must be newer or equal to Start Date"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateMIGThreshold
' Description   : Validates the MIG Threshold. MIG Threshold must be less than or equal to MIG Start LTV
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateMIGThreshold()
    On Error GoTo Failed
    
    Dim dMIGThreshold As Double
    Dim dMIGStartLTV As Double
    Dim sTmp As String
    
    With frmLenderDetails.txtLenderDetails
        sTmp = .Item(MIG_START_LTV).Text
    
        If Len(sTmp) > 0 Then
            dMIGStartLTV = CDbl(sTmp)
        End If
    
        If dMIGStartLTV <> 0 Then
            sTmp = .Item(MIG_THRESHOLD).Text
            If Len(sTmp) > 0 Then
                dMIGThreshold = CDbl(sTmp)
            
                If dMIGThreshold > dMIGStartLTV Then
                    .Item(MIG_THRESHOLD).SetFocus
                    g_clsErrorHandling.RaiseError errGeneralError, "HLC Threshold must be less than or equal to HLC Start LTV" 'JD BMIDS982
                End If
            End If
        End If
    End With

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''' Public Access methods '''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the edit state of this tab - edit or add.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the edit state of this tab - edit or add.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsLenderDetailsTable = clsTableAccess
End Sub
Public Function GetContactDetailsGUID() As Variant
    GetContactDetailsGUID = m_clsNameAndAddressDir.GetContactDetailsGUID()
End Function
Public Function GetFeeSetKey() As Variant
    GetFeeSetKey = m_clsLenderDetailsTable.GetLegalFeeSetKey()
End Function
Public Function GetRateSet() As String
    GetRateSet = m_clsLenderDetailsTable.GetRateSet()
End Function
Public Function GetAddressGUID() As Variant
    GetAddressGUID = m_clsNameAndAddressDir.GetAddressGUID()
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOrganisationID
' Description   : Returns the Lender OrganisationID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetOrganisationID() As Variant
    Dim vGUID As Variant
    Dim bSuccess As Boolean
    
    On Error GoTo Failed
    
    bSuccess = False
    
    vGUID = m_clsLenderDetailsTable.GetOrganisationID()

    If Not IsNull(vGUID) Then
        If Len(vGUID) > 0 Then
            bSuccess = True
        End If
    End If
    
    If Not bSuccess Then
        g_clsErrorHandling.RaiseError errOrganisationIDMissing
    End If
    
    GetOrganisationID = vGUID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub SetMIGRateKey(sRateSet As String)
    On Error GoTo Failed
        
    If Len(sRateSet) > 0 Then
        m_sMIGRateKey = sRateSet
    Else
        g_clsErrorHandling.RaiseError errRateSetEmpty, "LenderDetails.SetMIGRateKey"
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
