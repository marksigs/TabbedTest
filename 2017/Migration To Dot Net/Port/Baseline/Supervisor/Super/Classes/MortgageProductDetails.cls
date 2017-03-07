VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageProductDetails.cls
' Description   : Contains all processing for the main mortgage product tab on the Mortgage
'                 Product details screen
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port
' DJP       28/08/01    Don't show Manual Decrease Limit Percent unless the field is there.
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' DJP       04/12/01    SYS2831 Support client variants
' DJP/SA    21/12/01    SYS3545 Populate Country combo on Add/Edit
' STB       10/01/02    SYS3225 Products now use starttime as part of key.
' STB       20/05/02    SYS4609 Update admin system with product details.
' BMIDS
' AW        08/07/02    BMIDS00177/178Removed Admin system processing
' SA        20/09/02    BMIDS00246 Commented all refs to optAvailbleSelfEmployed as it has been removed from main screen.
' SA        04/11/02    BMIDS00290 ValidateArrangementFee changed to check percentage over 100.
' SA        06/11/02    BMIDS00290 ValidateArrangementFee changed.
' MC        25/05/2004  BMIDS775/REG016/018 CAT Indicator option button value handling code added
' MC        02/06/2004  BMIDS763    BMIDS763 - CC075 - Get and Set Selected New FEE SET member functions added
' HMA       20/12/2004  BMIDS959   Remove MortgageProductBands Table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' Prog      Date        Description
' TW        10/05/2002  EP2_7 Added processing for Additional Borrowing Fee and Credit Limit Increase Fee
' GHun      17/11/2006  EP2_19 Changed SetScreenFields and SaveFieldData to support FreeLegalFees
' TW        11/12/2006  EP2_20 Added processing for Transfer Of Equity Fee
' TW        23/02/2007  EP2_1354 - DBM183 Introduce refunded valuation products.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Control Constants
Private Const MORTGAGE_PRODUCT_CODE = 0
Private Const LENDER_CODE = 2
Private Const DISPLAY_ORDER = 3
Private Const PORTING_FEE = 4

Private Const START_DATE = 5
Private Const END_DATE = 6
Private Const WITHDRAWN_DATE = 7
Private Const MIN_APPLICANT_AGE = 8
Private Const MAX_APPLICANT_AGE = 10
Private Const MINIMUM_TERM = 12
Private Const MAXIMUM_TERM = 13
Private Const MINIMUM_LOAN_AMOUNT = 14
Private Const MINIMUM_LTV = 16
Private Const MAXIMUM_LOAN_AMOUNT = 15
Private Const MAXIMUM_LTV = 17
Private Const PREF_RATE_PERDIOD = 18

Private Const ARRANGEMENT_FEE_AMOUNT = 14
Private Const ARRANGEMENT_FEE_PERCENTAGE = 15
Private Const ARRANGEMENT_FEE_PERCENTAGE_MAX = 16
Private Const INTEREST_RATE_DECREASE = 1

' Private data
Private m_bIsEdit As Boolean
Private m_colProductKeys As Collection
Private m_clsLenderDetails As MortgageLendersTable
Private m_clsMortProdLanguage As MortProdLanguageTable
Private m_clsMortgageProductTable As MortgageProductTable

' Private key data
Private m_vOrgID As Variant

' Private Set data
Private m_sBaseRateSet As String
Private m_sAdminFeeSet As String
Private m_sValuationFeeSet As String
' TW 09/10/2006 EP2_7
Private m_sAdditionalBorrowingFeeSet As String
Private m_sCreditLimitIncreaseFeeSet As String
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
Private m_sTransferOfEquityFeeSet As String
' TW 11/12/2006 EP2_20 End

'*=[MC]BMIDS763
Private m_sIAFeeSet As String
Private m_sTTFeeSet As String
Private m_sProductSwitchFeeSet As String



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Class_Initialize()
    Set m_clsLenderDetails = New MortgageLendersTable
    Set m_clsMortProdLanguage = New MortProdLanguageTable
    Set m_clsMortgageProductTable = New MortgageProductTable
End Sub

Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    m_bIsEdit = bIsEdit
    
    ' SYS3545 Populate Country combo on Add/Edit
    ' Populate the Country of Origin combo
    g_clsFormProcessing.PopulateCombo "CountryOfOrigin", frmProductDetails.cboCountry
    
    PopulateLenderCodes
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetAddState()
    On Error GoTo Failed
    
    ' Need to Create required tables.
    g_clsFormProcessing.CreateNewRecord m_clsMortgageProductTable
    g_clsFormProcessing.CreateNewRecord m_clsMortProdLanguage
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetEditState()
    On Error GoTo Failed
    
    TableAccess(m_clsMortgageProductTable).SetKeyMatchValues m_colProductKeys
    TableAccess(m_clsMortgageProductTable).GetTableData
    TableAccess(m_clsMortgageProductTable).ValidateData
    
    If TableAccess(m_clsMortgageProductTable).RecordCount = 1 Then
        m_vOrgID = m_clsMortgageProductTable.GetOrganisationID
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate product - record count is " & TableAccess(m_clsMortgageProductTable).RecordCount
    End If
    
    ' Disable the Product code and Start Date + Time.
    frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).Enabled = False
    frmProductDetails.txtProductDetails(START_DATE).Enabled = False
    frmProductDetails.txtStartTime.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function GetLenderCode() As String
    On Error GoTo Failed
    Dim sLenderCode As String
    Dim clsLender As New MortgageLendersTable
    Dim colMatchValues As Collection
    
    If Len(m_vOrgID) > 0 Then
        Set colMatchValues = New Collection
        colMatchValues.Add m_vOrgID
        
        TableAccess(clsLender).SetKeyMatchValues colMatchValues
        TableAccess(clsLender).GetTableData
    
        If TableAccess(clsLender).RecordCount = 1 Then
            sLenderCode = clsLender.GetLenderCode()
        End If
        TableAccess(clsLender).CloseRecordSet
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "PopulateLenderCode - OrganisaionID is empty"
    End If
    
    GetLenderCode = sLenderCode
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub PopulateLenderCodes()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sLenderField As String
    
    Set rs = m_clsLenderDetails.GetLenderCodes()
    
    If Not rs Is Nothing Then
        Set frmProductDetails.cboLenderCode.RowSource = rs
        sLenderField = m_clsLenderDetails.GetLenderCodeField()
        frmProductDetails.cboLenderCode.ListField = sLenderField
    End If

    If rs.RecordCount = SINGLE_LENDER Then
        ' Single Lender
        frmProductDetails.cboLenderCode.SelText = rs(sLenderField)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub LoadMinTerm(txtTerm As MSGEditBox)
    On Error GoTo Failed:
    Dim vTerm As Variant
    Dim nTerm As Long

    vTerm = m_clsMortgageProductTable.GetMinimumTerm()
    
    If Not IsEmpty(vTerm) And Not IsNull(vTerm) Then
        If Len(vTerm) > 0 Then
            nTerm = CLng(vTerm) / 12
            txtTerm.Text = CStr(nTerm)
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub LoadMaxTerm(txtTerm As MSGEditBox)
    Dim vTerm As Variant
    Dim nTerm As Long
    On Error GoTo Failed
    vTerm = m_clsMortgageProductTable.GetMaximumTerm()
    
    If Not IsEmpty(vTerm) And Not IsNull(vTerm) Then
        If Len(vTerm) > 0 Then
            nTerm = CLng(vTerm) / 12
            txtTerm.Text = CStr(nTerm)
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveMinTerm(txtTerm As MSGEditBox)
    Dim sTerm As String
    Dim nTerm As Long
    On Error GoTo Failed
    
    sTerm = txtTerm.Text
    
    If Len(sTerm) > 0 And IsNumeric(sTerm) Then
        nTerm = CLng(sTerm) * 12
        
        m_clsMortgageProductTable.SetMinimumTerm CVar(nTerm)
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveMaxTerm(txtTerm As MSGEditBox)
    Dim sTerm As String
    Dim nTerm As Long
    On Error GoTo Failed
    
    sTerm = txtTerm.Text
    
    If Len(sTerm) > 0 And IsNumeric(sTerm) Then
        nTerm = CLng(sTerm) * 12
        
        m_clsMortgageProductTable.SetMaximumTerm CVar(nTerm)
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetProductLanguage()
    On Error GoTo Failed
    Dim vVal As Variant
    Dim sProductName As String
    Dim sProductDetails As String
    
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsMortgageProductTable
    ' Product code and start date
    Set colMatchValues = clsTableAccess.GetKeyMatchValues()
    
    Set clsTableAccess = m_clsMortProdLanguage
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    ' Find the MortageProductLanguage table for this product.
    clsTableAccess.GetTableData
    
    If clsTableAccess.RecordCount > 0 Then
        Dim bSetDefault As Boolean
        
        bSetDefault = True
        sProductName = m_clsMortProdLanguage.GetProductName()
        frmProductDetails.txtProductName.Text = sProductName

        sProductDetails = m_clsMortProdLanguage.GetProductDetails()
        frmProductDetails.txtDetails.Text = sProductDetails
        
        vVal = m_clsMortgageProductTable.GetCountryOfOrigin()
    
        If Not IsNull(vVal) And Not IsEmpty(vVal) Then
            If Len(CStr(vVal)) > 0 Then
                bSetDefault = False
                g_clsFormProcessing.HandleComboExtra frmProductDetails.cboCountry, vVal, SET_CONTROL_VALUE
            End If

        End If
    
        If bSetDefault Then
            If frmProductDetails.cboCountry.ListCount > 0 Then
                frmProductDetails.cboCountry.Text = frmProductDetails.cboCountry.List(0)
            End If
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveProductLanguage()
    On Error GoTo Failed
    Dim sLanguage As String
    Dim sLanguageID As String
    Dim sLanguageCombo As String
    Dim colValues As Collection
    Dim colIDS As New Collection
    Dim sProductCode As String
    Dim vStartDate As Variant

    ' Mortgage Product Language
    vStartDate = frmProductDetails.txtProductDetails(START_DATE).Text & " " & frmProductDetails.txtStartTime.Text
    sProductCode = frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    
    ' Set the key values for use by other tabs later.
    Set colValues = New Collection
    colValues.Add sProductCode
    colValues.Add vStartDate
    TableAccess(m_clsMortgageProductTable).SetKeyMatchValues colValues
    
    ' Reset the values collection
    Set colValues = New Collection
    
    sLanguageCombo = "MortgageProductLanguage"
    g_clsCombo.FindComboGroup sLanguageCombo, colValues, colIDS
    
    If Not IsEmpty(vStartDate) And Not IsNull(vStartDate) Then
        If Len(vStartDate) > 0 And Len(sProductCode) > 0 Then
            m_clsMortProdLanguage.SetProductCode sProductCode
            m_clsMortProdLanguage.SetStartDate vStartDate
            m_clsMortProdLanguage.SetProductName frmProductDetails.txtProductName.Text
            m_clsMortProdLanguage.SetProductDetails frmProductDetails.txtDetails.Text
                        
            If colValues.Count > 0 Then
                sLanguage = colValues(1) ' First one
                sLanguageID = colIDS(1)
                
                If Len(sLanguage) > 0 And Len(sLanguageID) > 0 Then
                    m_clsMortProdLanguage.SetLanguage sLanguageID
                End If
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveProductLanguage - Start Date and Product Code must be valid"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SaveProductLanguage - Start Date must be valid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetDateAndTime()
    On Error GoTo Failed
    Dim sTime As String
    Dim vStartDate As Variant
    
    vStartDate = m_clsMortgageProductTable.GetStartDate()
    
    sTime = Format(vStartDate, "hh:mm:ss")
    ' DJP SQL Server port - don't format date to four year format - leave to local locale setting
    g_clsFormProcessing.HandleDate frmProductDetails.txtProductDetails(START_DATE), vStartDate, SET_CONTROL_VALUE
    frmProductDetails.txtStartTime.Text = sTime
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetScreenFields()
    On Error GoTo Failed
    Dim vVal As Variant
    Dim sProductCode As String
    Dim vStartDate As Variant
    
    sProductCode = m_clsMortgageProductTable.GetMortgageProductCode()
    vStartDate = m_clsMortgageProductTable.GetStartDate()
    
    If Not IsEmpty(vStartDate) And Not IsNull(vStartDate) And Len(sProductCode) > 0 Then
        frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).Text = sProductCode
        SetDateAndTime
        
        GetProductLanguage
    
        g_clsFormProcessing.HandleDate frmProductDetails.txtProductDetails(END_DATE), m_clsMortgageProductTable.GetEndDate(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleDate frmProductDetails.txtProductDetails(WITHDRAWN_DATE), m_clsMortgageProductTable.GetWithdrawnDate(), SET_CONTROL_VALUE
    
        frmProductDetails.txtProductDetails(MIN_APPLICANT_AGE).Text = m_clsMortgageProductTable.GetMinApplicantAge()
        frmProductDetails.txtProductDetails(MAX_APPLICANT_AGE).Text = m_clsMortgageProductTable.GetMaxApplicantAge()
        
        LoadMinTerm frmProductDetails.txtProductDetails(MINIMUM_TERM)
        LoadMaxTerm frmProductDetails.txtProductDetails(MAXIMUM_TERM)
        
        frmProductDetails.txtProductDetails(MINIMUM_LOAN_AMOUNT).Text = m_clsMortgageProductTable.GetMinimumLoanAmount()
        frmProductDetails.txtProductDetails(MAXIMUM_LOAN_AMOUNT).Text = m_clsMortgageProductTable.GetMaximumLoanAmount()
        frmProductDetails.txtProductDetails(PORTING_FEE).Text = m_clsMortgageProductTable.GetPortingFee()
        frmProductDetails.txtProductDetails(DISPLAY_ORDER).Text = m_clsMortgageProductTable.GetDisplayOrder()
    
        frmProductDetails.txtProductDetails(MINIMUM_LTV).Text = m_clsMortgageProductTable.GetMinimumLTV()
        frmProductDetails.txtProductDetails(MAXIMUM_LTV).Text = m_clsMortgageProductTable.GetMaximumLTV()
        frmProductDetails.txtProductDetails(PREF_RATE_PERDIOD).Text = m_clsMortgageProductTable.GetPreferentialRatePeriod()
        
        frmProductDetails.txtProductDetails(INTEREST_RATE_DECREASE).Text = m_clsMortgageProductTable.GetInterestRateDecreaseLimit(frmProductDetails.lblInterestRateDecrease, frmProductDetails.txtProductDetails(INTEREST_RATE_DECREASE))
        
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optNonPanelLender(OPT_YES), frmProductDetails.optNonPanelLender(OPT_NO), m_clsMortgageProductTable.GetNonPanelLender(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optMortCalcAvailable(OPT_YES), frmProductDetails.optMortCalcAvailable(OPT_NO), m_clsMortgageProductTable.GetAvailableForMortgageCalculator(), SET_CONTROL_VALUE
        '-- SA BMIDS00246
        'g_clsFormProcessing.HandleRadioButtons frmProductDetails.optAvailableSelfEmployed(OPT_YES), frmProductDetails.optAvailableSelfEmployed(OPT_NO), m_clsMortgageProductTable.GetAvailableForSelfEmployed(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optExistingCustomer(OPT_YES), frmProductDetails.optExistingCustomer(OPT_NO), m_clsMortgageProductTable.GetExistCustomerMortgageProduct(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optImpairedCredit(OPT_YES), frmProductDetails.optImpairedCredit(OPT_NO), m_clsMortgageProductTable.GetAvailableForImpairedCredit(), SET_CONTROL_VALUE
        
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCashback(OPT_YES), frmProductDetails.optCashback(OPT_NO), m_clsMortgageProductTable.GetCashbackProduct(), SET_CONTROL_VALUE
        
        'EP2_19 GHun
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optFreeLegalFees(OPT_YES), frmProductDetails.optFreeLegalFees(OPT_NO), m_clsMortgageProductTable.GetFreeLegalFees(), SET_CONTROL_VALUE
        'EP2_19 End
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optAvailableQuickQuote(OPT_YES), frmProductDetails.optAvailableQuickQuote(OPT_NO), m_clsMortgageProductTable.GetAvailableQuickQuote(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optMemberOfStaffProduct(OPT_YES), frmProductDetails.optMemberOfStaffProduct(OPT_NO), m_clsMortgageProductTable.GetStaffMortgageProduct(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optFlexibleProduct(OPT_YES), frmProductDetails.optFlexibleProduct(OPT_NO), m_clsMortgageProductTable.GetFlexibleMortgageProduct(), SET_CONTROL_VALUE

' TW 23/02/2007 EP2_1354
        g_clsFormProcessing.HandleRadioButtons frmProductDetails.optRefundOfValuation(OPT_YES), frmProductDetails.optRefundOfValuation(OPT_NO), m_clsMortgageProductTable.GetRefundOfValuation(), SET_CONTROL_VALUE
' TW 23/02/2007 EP2_1354 End
    
        frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_AMOUNT).Text = m_clsMortgageProductTable.GetArrangementFeeAmount()
        frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE).Text = m_clsMortgageProductTable.GetArrangementFeePercent()
        frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MAX).Text = m_clsMortgageProductTable.GetArrangementFeePercentMax()
        
        vVal = GetLenderCode
        g_clsFormProcessing.HandleDataComboText frmProductDetails.cboLenderCode, vVal, SET_CONTROL_VALUE
        frmProductDetails.cboLenderCode.Enabled = False
        If m_bIsEdit Then
            frmProductDetails.txtProductDetails(START_DATE).Enabled = False
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Start Date and Product code are invalid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function SetOrgIDFromCode() As Variant
    On Error GoTo Failed
    Dim sLenderCode As String
    
    sLenderCode = frmProductDetails.cboLenderCode.BoundText
    m_clsLenderDetails.FindLender sLenderCode
    
    m_vOrgID = m_clsLenderDetails.GetOrganisationID
    
    SetOrgIDFromCode = m_vOrgID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SaveOrgID()
    On Error GoTo Failed
    Dim sLenderCode As String
    
    SetOrgIDFromCode
    m_clsMortgageProductTable.SetOrganisationID m_vOrgID
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()
    Dim vTmp As Variant
    SaveProductLanguage
    SaveSets
    
    If Not m_bIsEdit Then
        SaveOrgID
    End If
    
    m_clsMortgageProductTable.SetStartDate frmProductDetails.txtProductDetails(START_DATE).Text & " " & frmProductDetails.txtStartTime.Text
    
    g_clsFormProcessing.HandleDate frmProductDetails.txtProductDetails(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetEndDate vTmp

    g_clsFormProcessing.HandleDate frmProductDetails.txtProductDetails(WITHDRAWN_DATE), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetWithdrawnDate vTmp

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optNonPanelLender(OPT_YES), frmProductDetails.optNonPanelLender(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetNonPanelLender vTmp

    m_clsMortgageProductTable.SetMinApplicantAge (frmProductDetails.txtProductDetails(MIN_APPLICANT_AGE).Text)
    m_clsMortgageProductTable.SetMaxApplicantAge (frmProductDetails.txtProductDetails(MAX_APPLICANT_AGE).Text)
    SaveMinTerm frmProductDetails.txtProductDetails(MINIMUM_TERM)
    SaveMaxTerm frmProductDetails.txtProductDetails(MAXIMUM_TERM)
    
    m_clsMortgageProductTable.SetMinimumLoanAmount (frmProductDetails.txtProductDetails(MINIMUM_LOAN_AMOUNT).Text)
    m_clsMortgageProductTable.SetMaximumLoanAmount (frmProductDetails.txtProductDetails(MAXIMUM_LOAN_AMOUNT).Text)
    m_clsMortgageProductTable.SetPortingFee (frmProductDetails.txtProductDetails(PORTING_FEE).Text)
    m_clsMortgageProductTable.SetDisplayOrder (frmProductDetails.txtProductDetails(DISPLAY_ORDER).Text)
    
    m_clsMortgageProductTable.SetMinimumLTV (frmProductDetails.txtProductDetails(MINIMUM_LTV).Text)
    m_clsMortgageProductTable.SetMaximumLTV (frmProductDetails.txtProductDetails(MAXIMUM_LTV).Text)
    m_clsMortgageProductTable.SetPreferentialRatePeriod (frmProductDetails.txtProductDetails(PREF_RATE_PERDIOD).Text)

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optMortCalcAvailable(OPT_YES), frmProductDetails.optMortCalcAvailable(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetAvailableForMortgageCalculator (vTmp)
    '--BMIDS00246 SA
    'g_clsFormProcessing.HandleRadioButtons frmProductDetails.optAvailableSelfEmployed(OPT_YES), frmProductDetails.optAvailableSelfEmployed(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetAvailableForSelfEmployed (vTmp)

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optImpairedCredit(OPT_YES), frmProductDetails.optImpairedCredit(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetAvailableForImpairedCredit (vTmp)

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCashback(OPT_YES), frmProductDetails.optCashback(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetCashbackProduct (vTmp)
    
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optExistingCustomer(OPT_YES), frmProductDetails.optExistingCustomer(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetExistCustomerMortgageProduct (vTmp)
    
     '*=[MC]BMIDS775/REG016/018 Regulation change
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCATStandard(OPT_YES), frmProductDetails.optCATStandard(OPT_NO), vTmp, GET_CONTROL_VALUE
    Call m_clsMortgageProductTable.SetCATIndicator(vTmp)
    '*=[END] REG016/BMIDS0775
    
    'EP2_19 GHun
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optFreeLegalFees(OPT_YES), frmProductDetails.optFreeLegalFees(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetFreeLegalFees (vTmp)
    'EP2_19 End
    
' TW 23/02/2007 EP2_1354
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optRefundOfValuation(OPT_YES), frmProductDetails.optRefundOfValuation(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetRefundOfValuation (vTmp)
' TW 23/02/2007 EP2_1354 End
    
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optAvailableQuickQuote(OPT_YES), frmProductDetails.optAvailableQuickQuote(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetAvailableQuickQuote (vTmp)

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optMemberOfStaffProduct(OPT_YES), frmProductDetails.optMemberOfStaffProduct(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetStaffMortgageProduct (vTmp)

    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optFlexibleProduct(OPT_YES), frmProductDetails.optFlexibleProduct(OPT_NO), vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetFlexibleMortgageProduct (vTmp)
    m_clsMortgageProductTable.SetArrangementFeeAmount frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_AMOUNT).Text
    m_clsMortgageProductTable.SetArrangementFeePercent frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE).Text
    m_clsMortgageProductTable.SetArrangementFeePercentMax frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MAX).Text
    m_clsMortgageProductTable.SetMortgageProductCode frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    m_clsMortgageProductTable.SetInterestRateDecreaseLimit frmProductDetails.txtProductDetails(INTEREST_RATE_DECREASE).Text


    ' Country of Origin
    g_clsFormProcessing.HandleComboExtra frmProductDetails.cboCountry, vTmp, GET_CONTROL_VALUE
    m_clsMortgageProductTable.SetCountryOfOrigin vTmp
    
    'TODO: SYS4609 Re-enable this call when Optimus is up to date.
    'Send the update request to ODI.
    'UpdateODIData
    
    DoUpdates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetMortgageProductCode() As Variant
    GetMortgageProductCode = m_clsMortgageProductTable.GetMortgageProductCode()
End Function

Public Function GetStartDate() As Variant
    GetStartDate = m_clsMortgageProductTable.GetStartDate()
End Function

Public Function GetOrganisationID() As Variant
    On Error GoTo Failed
    
    ' Need to get it from the LenderTable
    If IsEmpty(m_vOrgID) Then
        SetOrgIDFromCode
    End If
    
    GetOrganisationID = m_vOrgID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function ValidateScreenData() As Boolean
    
    On Error GoTo Failed
    
    'If the start time is not equal to a blank mask, then validate it.
    If frmProductDetails.txtStartTime.Text <> "__:__:__" Then
        'Ensure the starttime is a valid time.
        If IsDate(Format$(Now, "dd/mm/yyyy") & " " & frmProductDetails.txtStartTime.Text) = False Then
            frmProductDetails.txtStartTime.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "The Start Time specified is not a valid time"
        End If
    Else
        'Default the time to midnight.
        frmProductDetails.txtStartTime.Text = "00:00:00"
    End If
        
    If Not m_bIsEdit Then
        CheckIfProductExists
    End If
    
    g_clsValidation.ValidateProductDates frmProductDetails.txtProductDetails(START_DATE), frmProductDetails.txtProductDetails(END_DATE), frmProductDetails.txtProductDetails(WITHDRAWN_DATE)
    ValidateArrangementFee
    ValidateMinAndMax
    
    ' BMIDS959  Remove ValidateRedemption
    ' ValidateRedemption
    
    ValidateScreenData = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub ValidateArrangementFee()
    On Error GoTo Failed
    Dim sFee As String
    Dim sFeePercent As String
    
    sFee = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_AMOUNT).Text
    sFeePercent = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE).Text
    
    If Len(sFee) > 0 And Len(sFeePercent) > 0 Then
        frmProductDetails.SSTab1.Tab = mortAreaParameters - 1
        g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_AMOUNT)
        g_clsErrorHandling.RaiseError errGeneralError, "Only one of Arrangement Fee Amount and Arrangement Fee Percent can be entered"
    End If
    '++ SA BMIDS00290 {
    '++ BMIDS00290 6/11 check length and convert first!!
    If Len(sFeePercent) > 0 Then
        If CLng(sFeePercent) > 100 Then
            frmProductDetails.SSTab1.Tab = mortAreaParameters - 1
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE)
            g_clsErrorHandling.RaiseError errGeneralError, "Arrangement Fee Percent cannot be greater than 100."
        End If
    End If
    
    '++BMIDS00290 }
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' BMIDS959  Remove Sub ValidateRedemption() and Sub ValidateRedemptionPeriod
Private Sub ValidateMinAndMax()
    On Error GoTo Failed
    Dim sMinLTV As String
    Dim sMaxLTV As String
    Dim sMinAge As String
    Dim sMaxAge As String
    Dim sMinTerm As String
    Dim sMaxTerm As String
    Dim sMinAmount As String
    Dim sMaxAmount As String
    
    sMinLTV = frmProductDetails.txtProductDetails(MINIMUM_LTV).Text
    sMaxLTV = frmProductDetails.txtProductDetails(MAXIMUM_LTV).Text
    
    If Len(sMinLTV) > 0 And Len(sMaxLTV) > 0 Then
        If CLng(sMaxLTV) < CLng(sMinLTV) Then
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtProductDetails(MINIMUM_LTV)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum LTV cannot be less than Minimum LTV"
        End If
    End If
    
    sMinTerm = frmProductDetails.txtProductDetails(MINIMUM_TERM).Text
    sMaxTerm = frmProductDetails.txtProductDetails(MAXIMUM_TERM).Text
    
    If Len(sMinTerm) > 0 And Len(sMaxTerm) > 0 Then
        If CLng(sMaxTerm) < CLng(sMinTerm) Then
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtProductDetails(MINIMUM_TERM)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Term cannot be less than Minimum Term"
        End If
    End If
    
    sMinAmount = frmProductDetails.txtProductDetails(MINIMUM_LOAN_AMOUNT).Text
    sMaxAmount = frmProductDetails.txtProductDetails(MAXIMUM_LOAN_AMOUNT).Text
    
    If Len(sMinAmount) > 0 And Len(sMaxAmount) > 0 Then
        If CLng(sMaxAmount) < CLng(sMinAmount) Then
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtProductDetails(MINIMUM_LOAN_AMOUNT)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Loan Amount cannot be less than Minimum Loan Amount"
        End If
    End If
    
    sMinAge = frmProductDetails.txtProductDetails(MIN_APPLICANT_AGE).Text
    sMaxAge = frmProductDetails.txtProductDetails(MAX_APPLICANT_AGE).Text
    
    If Len(sMinAge) > 0 And Len(sMaxAge) > 0 Then
        If CLng(sMaxAge) < CLng(sMinAge) Then
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtProductDetails(MIN_APPLICANT_AGE)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum Applicant Age cannot be less than Minimum Applicant Age"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CheckIfProductExists()
    On Error GoTo Failed
    Dim bExists As Boolean
    Dim sProductCode As String
    Dim sProductStartDate As String
    Dim clsMortgageProduct As MortgageProductTable
    
    Set clsMortgageProduct = New MortgageProductTable
    
    sProductCode = frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).Text
    sProductStartDate = frmProductDetails.txtProductDetails(START_DATE).Text & " " & frmProductDetails.txtStartTime.Text
    
    bExists = clsMortgageProduct.DoesProductExist(sProductCode, sProductStartDate)

    If bExists Then
        frmProductDetails.txtProductDetails(MORTGAGE_PRODUCT_CODE).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Product exists - please enter a unique combination of Product Code and Start Date"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Public Sub DoUpdates()
    On Error GoTo Failed
    
    ' Do updates
    TableAccess(m_clsMortgageProductTable).Update
    TableAccess(m_clsMortProdLanguage).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAdminFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sAdminFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'*=[MC]BMIDS763
Public Sub SetInsuranceAdminFeeSet(sFeeSet As String)
    m_sIAFeeSet = sFeeSet
End Sub

Public Sub SetTTFeeSet(sFeeSet As String)
    m_sTTFeeSet = sFeeSet
End Sub

Public Sub SetProductSwitchFeeSet(sFeeSet As String)
    m_sProductSwitchFeeSet = sFeeSet
End Sub

'*=SECTION END


Public Sub SetBaseRateSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sBaseRateSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetValuationFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sValuationFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAdditionalBorrowingFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sAdditionalBorrowingFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetTransferOfEquityFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sTransferOfEquityFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetCreditLimitIncreaseFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sCreditLimitIncreaseFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub SaveSets()
    On Error GoTo Failed
    
    m_clsMortgageProductTable.SetAdminFeeSet m_sAdminFeeSet
    m_clsMortgageProductTable.SetBaseRateFeeSet m_sBaseRateSet
    m_clsMortgageProductTable.SetValuationFeeSet m_sValuationFeeSet
    
    '*=[MC]BMIDS763
    m_clsMortgageProductTable.SetInsuranceAdminFeeSet m_sIAFeeSet
    m_clsMortgageProductTable.SetTTFeeSet m_sTTFeeSet
    m_clsMortgageProductTable.SetProductSwitchFeeSet m_sProductSwitchFeeSet
    '*=SECTION END
    
' TW 09/10/2006 EP2_7
    m_clsMortgageProductTable.SetAdditionalBorrowingFeeSet m_sAdditionalBorrowingFeeSet
    m_clsMortgageProductTable.SetCreditLimitIncreaseFeeSet m_sCreditLimitIncreaseFeeSet
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    m_clsMortgageProductTable.SetTransferOfEquityFeeSet m_sTransferOfEquityFeeSet
' TW 11/12/2006 EP2_20 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdTypeOfBuyer.SetProductKeys"
    End If
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsMortgageProductTable = clsTableAccess
End Sub

Public Function GetProductTable() As MortgageProductTable
    On Error GoTo Failed
    
    Set GetProductTable = m_clsMortgageProductTable
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub Class_Terminate()
    Set m_clsMortProdLanguage = Nothing
    Set m_clsMortgageProductTable = Nothing
End Sub


Private Sub UpdateODIData()
    
End Sub


