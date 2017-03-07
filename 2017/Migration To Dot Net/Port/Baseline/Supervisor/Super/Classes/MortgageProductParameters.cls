VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Text Edits

' CheckBoxes

Private m_rs As ADODB.Recordset
Private m_clsTableAccess As TableAccess
Private m_clsProductParameters As MortgageProductParametersTable
Private m_clsProductBands As MortgageProductBandsTable
Sub Class_Initialize()

End Sub
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsLenderDetailsTable = clsTableAccess
End Sub
Public Function Initialise() As Boolean
    Dim bRet As Boolean

    If m_clsLenderDetailsTable Is Nothing Then
        Set m_clsLenderDetailsTable = New MortgageLendersTable
    End If
    
    Set m_clsTableAccess = m_clsLenderDetailsTable
    Set m_rs = m_clsTableAccess.GetTableData()
    
    Initialise = True
    Exit Function
Failed:
    MsgBox "LenderDetails:Initialise, Error is " + Err.Description
    Initialise = False
End Function
Public Function SetScreenFields()
    Dim bRet As Boolean
    On Error GoTo Failed
    bRet = True
    
    'frmLenderDetails.txtLenderDetails(LENDER_CODE).Text = m_clsLenderDetailsTable.GetLenderCode()
frmLenderDetails.txtLenderDetails(LENDER_CODE).Text = m_clsLenderDetailsTable.GetOrganisationID()
    frmLenderDetails.txtLenderDetails(LENDER_NAME).Text = m_clsLenderDetailsTable.GetLenderName()
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(START_DATE), m_clsLenderDetailsTable.GetStartDate(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleDate frmLenderDetails.txtLenderDetails(END_DATE), m_clsLenderDetailsTable.GetEndDate(), SET_CONTROL_VALUE
    frmLenderDetails.txtLenderDetails(MAX_LOAN_COMPONENTS).Text = m_clsLenderDetailsTable.GetMaxNoLoans()
    frmLenderDetails.txtLenderDetails(MIG_START_LTV).Text = m_clsLenderDetailsTable.GetMIGStartLTV()
    frmLenderDetails.txtLenderDetails(IGNORE_MIG_PREMIUM).Text = m_clsLenderDetailsTable.GetIgnoreMIGPremium()
    
    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddIPT(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), m_clsLenderDetailsTable.GetAddIPTToMIGPremium(), SET_CONTROL_VALUE
    
    frmLenderDetails.txtLenderDetails(TT_FEE).Text = m_clsLenderDetailsTable.GetTTFee()
    frmLenderDetails.txtLenderDetails(SEALING_FEE).Text = m_clsLenderDetailsTable.GetSealingFee()
    frmLenderDetails.txtLenderDetails(DEEDS_RELEASE_FEE).Text = m_clsLenderDetailsTable.GetDeedsReleaseFee()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE).Text = m_clsLenderDetailsTable.GetArrangementFeeAmount()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT).Text = m_clsLenderDetailsTable.GetArrangementFeePercent()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX).Text = m_clsLenderDetailsTable.GetArrangementFeePercentMax()

    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddToLoan(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), m_clsLenderDetailsTable.GetAddCostsToLoan(), SET_CONTROL_VALUE
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_PORTING_FEE), m_clsLenderDetailsTable.GetAllowPortingFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ADMIN_FEE), m_clsLenderDetailsTable.GetAllowPortingFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_VALUATION_FEE), m_clsLenderDetailsTable.GetAllowValuationFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_REINSPECTION_FEE), m_clsLenderDetailsTable.GetAllowReinspectionFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_MIG_FEE), m_clsLenderDetailsTable.GetAllowMIGFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_TT_FEE), m_clsLenderDetailsTable.GetAllowTTFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_SEALING_FEE), m_clsLenderDetailsTable.GetAllowSealingFeeAdded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_DEEDS_RELEASE_FEE), m_clsLenderDetailsTable.GetAllowDeedsReleaseFeeAdded(), SET_CONTROL_VALUE

    SetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "SetScreenFields: Error is " + Err.Description
    SetScreenFields = False
End Function
Public Function GetScreenFields()
    Dim bRet As Boolean
    Dim vTmp As Variant
    bRet = True
    
    m_clsLenderDetailsTable.SetLenderName frmLenderDetails.txtLenderDetails(LENDER_NAME)
    'm_clsLenderDetailsTable.SetLenderCode frmLenderDetails.txtLenderDetails(LENDER_CODE)
    m_clsLenderDetailsTable.SetOrganisationID frmLenderDetails.txtLenderDetails(LENDER_CODE)
    g_clsFormProcessing.HandleDate frmLenderDetails(START_DATE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetStartDate vTmp
    
    g_clsFormProcessing.HandleDate frmLenderDetails(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetEndDate vTmp
    
    frmLenderDetails.txtLenderDetails(MAX_LOAN_COMPONENTS) = m_clsLenderDetailsTable.GetMaxNoLoans()
    frmLenderDetails.txtLenderDetails(MIG_START_LTV) = m_clsLenderDetailsTable.GetMIGStartLTV()
    frmLenderDetails.txtLenderDetails(IGNORE_MIG_PREMIUM) = m_clsLenderDetailsTable.GetIgnoreMIGPremium()
    
    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddIPT(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), m_clsLenderDetailsTable.GetAddIPTToMIGPremium(), SET_CONTROL_VALUE
    
    frmLenderDetails.txtLenderDetails(TT_FEE) = m_clsLenderDetailsTable.GetTTFee()
    frmLenderDetails.txtLenderDetails(SEALING_FEE) = m_clsLenderDetailsTable.GetSealingFee()
    frmLenderDetails.txtLenderDetails(DEEDS_RELEASE_FEE) = m_clsLenderDetailsTable.GetDeedsReleaseFee()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE) = m_clsLenderDetailsTable.GetArrangementFeeAmount()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT) = m_clsLenderDetailsTable.GetArrangementFeePercent()
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX) = m_clsLenderDetailsTable.GetArrangementFeePercentMax()

    g_clsFormProcessing.HandleRadioButtons frmLenderDetails.optAddToLoan(OPT_YES), frmLenderDetails.optAddIPT(OPT_NO), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAddCostsToLoan (vTmp)
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_PORTING_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowPortingFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_ADMIN_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowPortingFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_VALUATION_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowValuationFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_REINSPECTION_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowReinspectionFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_MIG_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowMIGFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_TT_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowTTFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_SEALING_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowSealingFeeAdded vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_DEEDS_RELEASE_FEE), vTmp, SET_CONTROL_VALUE
    m_clsLenderDetailsTable.SetAllowDeedsReleaseFeeAdded vTmp
    
    GetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "GetScreenFields: Error is " + Err.Description
    GetScreenFields = False
End Function
Friend Function GetContactDetailsGUID() As Variant
    GetContactDetailsGUID = m_clsLenderDetailsTable.GetContactDetailsGUID()
End Function
Friend Function GetAddressGUID() As Variant
    GetAddressGUID = m_clsLenderDetailsTable.GetAddressGUID()
End Function
Friend Function GetOrganisationID() As Variant
    GetOrganisationID = m_clsLenderDetailsTable.GetOrganisationID()
End Function
Friend Sub SetMortgageParameterKeys(colValues As Collection)
    m_clsTableAccess.SetKeyMatchValues colValues
End Sub
