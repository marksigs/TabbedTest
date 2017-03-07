VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductDetailsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageProductDetailsCS
' Description   : Client specific version of the MortgageProductDetails class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       05/12/01    SYS2831 Support client variants
'--------------------------------------------------------------------------------------------
' BMIDS Change History
'---------------------
' Prog      Date        Description
' GD        10/05/2002  Client Specific : BMIDS Changes,
'                       Deuce Ref : BMIDS00002(MASTER) AND BMIDS00007(CHILD)
'                       For Requirement BM014("Can Be Ported") and BM019("Exclusive Or Semi Exclusive")
' GD        20/05/2002  Client Specific : BMIDS Changes,Deuce Ref : BMIDS0011
'                       For requirement BM017(part of)
' GD        23/05/2002  BMIDS00014
' MO        11/06/02    BMIDS00040 : Added in validation for arrangement fees.
' MC        25/05/2004  [MC]BMIDS775/REG016/018 CAT indicator option button handling code added
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
' JD        16/06/2004  BMIDS765 CC076 Rental Income Rate changes
' JD        28/07/04    BMIDS809 Add a <Select> option to the rental income rate combo
' HM        23/08/2005  WP16 MAR42 - ERC Free Percentage is added on Other Rates/Groups Tab
'--------------------------------------------------------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' Prog      Date        Description
' TW        10/05/2002  09/10/2006  EP2_7 Added processing for Additional Borrowing Fee and Credit Limit Increase Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Base class
Implements MortgageProductDetails


' Private data
Private m_clsProductDetails As MortgageProductDetails

Private Const ARRANGEMENT_FEE_PERCENTAGE_MIN = 17


' Text Edit boxes
Private Const ARRANGEMENT_FEE = 14
Private Const ARRANGEMENT_FEE_PERCENTAGE = 15
Private Const ARRANGEMENT_FEE_PERCENTAGE_MAX = 16
Private Sub Class_Initialize()
    Set m_clsProductDetails = New MortgageProductDetails
End Sub

Private Function MortgageProductDetails_GetProductTable() As MortgageProductTable
    Set MortgageProductDetails_GetProductTable = m_clsProductDetails.GetProductTable
End Function

Private Sub MortgageProductDetails_InitTab()
    m_clsProductDetails.InitTab
End Sub


Private Sub MortgageProductDetails_SaveScreenData()
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    'GD Save BMIDS Screen data
    'GD PLUS BMIDS0011
    SaveScreenData
    m_clsProductDetails.SaveScreenData

    
End Sub


Private Sub MortgageProductDetails_DoUpdates()
    m_clsProductDetails.DoUpdates
End Sub

Private Function MortgageProductDetails_GetMortgageProductCode() As Variant
    MortgageProductDetails_GetMortgageProductCode = m_clsProductDetails.GetMortgageProductCode
End Function

Private Function MortgageProductDetails_GetOrganisationID() As Variant
    MortgageProductDetails_GetOrganisationID = m_clsProductDetails.GetOrganisationID
End Function

Private Function MortgageProductDetails_GetStartDate() As Variant
    MortgageProductDetails_GetStartDate = m_clsProductDetails.GetStartDate
End Function


Private Sub MortgageProductDetails_Initialise(bIsEdit As Boolean)
'GD BMIDS0011
    m_clsProductDetails.Initialise bIsEdit

    PopulateRedemptionFeeSetCombo
    'GD BMIDS 0014
    PopulateMIGRateSetCombo
    PopulateIncomeMultipleSetCombo
    PopulateRentalIncomeRateSetCombo  'JD BMIDS765
    'Populate Special Groups Combo
    'g_clsFormProcessing.PopulateCombo "SpecialGroup", frmProductDetails.cboSpecialGroups
End Sub


Private Sub MortgageProductDetails_SetAdditionalBorrowingFeeSet(sFeeSet As String)
    m_clsProductDetails.SetAdditionalBorrowingFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetAdminFeeSet(sFeeSet As String)
    m_clsProductDetails.SetAdminFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetBaseRateSet(sFeeSet As String)
    m_clsProductDetails.SetBaseRateSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetCreditLimitIncreaseFeeSet(sFeeSet As String)
    m_clsProductDetails.SetCreditLimitIncreaseFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetInsuranceAdminFeeSet(sFeeSet As String)
    m_clsProductDetails.SetInsuranceAdminFeeSet sFeeSet
End Sub

Private Function MortgageProductDetails_SetOrgIDFromCode() As Variant
    MortgageProductDetails_SetOrgIDFromCode = m_clsProductDetails.SetOrgIDFromCode
End Function

Private Sub MortgageProductDetails_SetProductKeys(colValues As Collection)
    m_clsProductDetails.SetProductKeys colValues
End Sub


Private Sub MortgageProductDetails_SetProductSwitchFeeSet(sFeeSet As String)
    m_clsProductDetails.SetProductSwitchFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetScreenFields()

    m_clsProductDetails.SetScreenFields
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER)
    'GD Set BMIDS Screen Fields
    'GD PLUS BMIDS0011
    SetScreenFields
    
End Sub


Private Sub MortgageProductDetails_SetTableClass(clsTableAccess As TableAccess)
    m_clsProductDetails.SetTableClass clsTableAccess
End Sub

Private Sub MortgageProductDetails_SetTransferOfEquityFeeSet(sFeeSet As String)
    m_clsProductDetails.SetTransferOfEquityFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetTTFeeSet(sFeeSet As String)
    m_clsProductDetails.SetTTFeeSet sFeeSet
End Sub

Private Sub MortgageProductDetails_SetValuationFeeSet(sFeeSet As String)
    m_clsProductDetails.SetValuationFeeSet sFeeSet
End Sub

Private Function MortgageProductDetails_ValidateScreenData() As Boolean
    MortgageProductDetails_ValidateScreenData = m_clsProductDetails.ValidateScreenData
    MortgageProductDetails_ValidateScreenData = ValidateClientScreenData
End Function

Private Sub SetScreenFields()
    Dim clsProductTable As MortgageProductTable
    'GD BMIDS0014
    Dim clsRedemptionFeeSetTable As RedemptionFeeSetTable
    Dim clsMIGRateSetTable As MPMigRateSetTable
    Dim clsIncomeMultipleSetTable As IncomeMultipleSetTable
    Dim clsRentalIncomeRateSetBandTable As RentalIncomeRateSetBandTable  'JD BMIDS765
    Dim vTmp As Variant
    Dim colKeys As New Collection
    Dim vTmp2 As Variant
    Set clsProductTable = m_clsProductDetails.GetProductTable
    
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER)
    'GD Set BMIDS Screen Fields
    
    '---------------------Can Be Ported Indicator Radio Buttons----------------------------
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCanBePorted(OPT_YES), frmProductDetails.optCanBePorted(OPT_NO), clsProductTable.GetCanBePorted(), SET_CONTROL_VALUE

     '*=[MC]BMIDS775/REG016/018
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCATStandard(OPT_YES), frmProductDetails.optCATStandard(OPT_NO), clsProductTable.GetCATIndicator(), SET_CONTROL_VALUE
    '[END]REG016/BMIDS775
   
    '---------------------Exclusive Or Semi-Exclusive Indicator Radio Buttons----------------
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optExclusiveOrSemi(OPT_YES), frmProductDetails.optExclusiveOrSemi(OPT_NO), clsProductTable.GetExclusiveOrSemi(), SET_CONTROL_VALUE
    'GD BMIDS0011
    
    '---------------------Populate the Redemption Fee Set combo based on data is DB-----------
    vTmp = clsProductTable.GetRedemptionFeeSet ' Mortgageproduct.RedemptionFeeSet
    If vTmp <> "" Then
        colKeys.Add vTmp
        Set clsRedemptionFeeSetTable = New RedemptionFeeSetTable
        TableAccess(clsRedemptionFeeSetTable).SetKeyMatchValues colKeys
        TableAccess(clsRedemptionFeeSetTable).GetTableData
        vTmp2 = clsRedemptionFeeSetTable.GetDescription 'RedemptionFeeSet.RedemptionFeeSetDescription
        frmProductDetails.cboRedemptionFeeSet.Text = vTmp & " - " & vTmp2
    Else
        'Set to "None"
        frmProductDetails.cboRedemptionFeeSet.Text = ""
    End If
    'GD BMIDS 0014
    Dim colKeysMIG As New Collection
    
    '---------------------Populate the MIG Rate Set combo based on data is DB----------------
    vTmp = clsProductTable.GetMPMigRateSet ' Mortgageproduct.MPMIGRateSet
    If vTmp <> "" Then
        colKeysMIG.Add vTmp
        Set clsMIGRateSetTable = New MPMigRateSetTable
        TableAccess(clsMIGRateSetTable).SetKeyMatchValues colKeysMIG
        TableAccess(clsMIGRateSetTable).GetTableData
        vTmp2 = clsMIGRateSetTable.GetMPMigDescription 'MPMIGRateSet.MIGRateSetDescription
        frmProductDetails.cboMIGRateSet.Text = vTmp & " - " & vTmp2
    Else
        'Set to "None"
        frmProductDetails.cboMIGRateSet.Text = ""
    End If
     'GD BMIDS 0014
    Dim colKeysIMS As New Collection
    
    '-------------------Populate the Income Multiple Set combo based on data is DB----------
    vTmp = clsProductTable.GetIncomeMultiplierCode ' Mortgageproduct.IncomeMultiplierCode
    If vTmp <> "" Then
        colKeysIMS.Add vTmp
        Set clsIncomeMultipleSetTable = New IncomeMultipleSetTable
        TableAccess(clsIncomeMultipleSetTable).SetKeyMatchValues colKeysIMS
        TableAccess(clsIncomeMultipleSetTable).GetTableData
        vTmp2 = clsIncomeMultipleSetTable.GetIncMultDesc 'IncomeMultipleSet.IncomeMultiplierDescription
        frmProductDetails.cboIncomeMultipleSet.Text = vTmp & " - " & vTmp2
    Else
        'Set to "None"
        frmProductDetails.cboIncomeMultipleSet.Text = ""
    End If
    
    '---------------------Populate the Rental Income Rate Set combo based on data in DB---- JD BMIDS765 -------
    Dim colKeysRIR As New Collection
    vTmp = clsProductTable.GetRentalIncomeRateSet ' Mortgageproduct.RentalIncomeRateSet
    If vTmp <> "" Then
        colKeysRIR.Add vTmp
        Set clsRentalIncomeRateSetBandTable = New RentalIncomeRateSetBandTable
        'TableAccess(clsRentalIncomeRateSetBandTable).SetKeyMatchValues colKeysRIR
        'TableAccess(clsRentalIncomeRateSetBandTable).GetTableData
        vTmp2 = clsRentalIncomeRateSetBandTable.GetMaxStartDateForSet(vTmp) 'RentalIncomeRateSetBand.StartDate
        frmProductDetails.cboRentalIncomeRateSet.Text = vTmp & " - " & vTmp2
    Else
        'Set to "None"
        frmProductDetails.cboRentalIncomeRateSet.Text = COMBO_NONE  'JD BMIDS809
    End If

    '-----------------Populate the MIG Start LTV text box on Misc tab---------------------
    frmProductDetails.txtMPMIGStartLTV.Text = clsProductTable.GetMPMIGStartLTV
    ' HM MAR42
    frmProductDetails.txtERCFreePercentage.Text = clsProductTable.GetFreeChargePercentage
    '----------------Populate Min Percent Arrangement Fee on main tab
    frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN).Text = clsProductTable.GetMinPercentArrangementFee
    
End Sub

Private Sub SaveScreenData()

    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    'GD Save BMIDS Screen Data
    On Error GoTo Failed
    
    Dim vTmp As Variant
    Dim clsProductTable As MortgageProductTable
    
    Set clsProductTable = m_clsProductDetails.GetProductTable
    
    '"Can Be Ported" Radio Buttons
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optCanBePorted(OPT_YES), frmProductDetails.optCanBePorted(OPT_NO), vTmp, GET_CONTROL_VALUE

    clsProductTable.SetCanBePorted vTmp
    
    '"Exclusive Or Semi-Exclusive" Radio Buttons
    g_clsFormProcessing.HandleRadioButtons frmProductDetails.optExclusiveOrSemi(OPT_YES), frmProductDetails.optExclusiveOrSemi(OPT_NO), vTmp, GET_CONTROL_VALUE

    clsProductTable.SetExclusiveOrSemi vTmp
    
    'GD BMIDS0011
    'Redemption Fee Set Combo
    Dim rs As ADODB.Recordset
    Dim clsRedemptionFeeSet As New RedemptionFeeSetTable
    Dim clsTableAccess As TableAccess
    Dim iSelectedIndex As Integer
    Dim sValueToStore As String
    
    clsRedemptionFeeSet.GetRedemptionFeeSet
    Set rs = TableAccess(clsRedemptionFeeSet).GetRecordSet
    'check if null value is in combo.
    sValueToStore = ""
    If frmProductDetails.cboRedemptionFeeSet.SelectedItem <> "" Then
        iSelectedIndex = frmProductDetails.cboRedemptionFeeSet.SelectedItem
        rs.Move (iSelectedIndex - 1)
        sValueToStore = clsRedemptionFeeSet.GetTheRedemptionFeeSet
    End If
    clsProductTable.SetRedemptionFeeSet (sValueToStore)
    
    
    'GD BMIDS0014
    'MIG Rate Set Combo
    Dim clsMIGRateSet As New MPMigRateSetTable
    clsMIGRateSet.GetMPMigRateSet
    Set rs = TableAccess(clsMIGRateSet).GetRecordSet
    'check if null value is in combo.
    sValueToStore = ""
    If frmProductDetails.cboMIGRateSet.SelectedItem <> "" Then
        iSelectedIndex = frmProductDetails.cboMIGRateSet.SelectedItem
        rs.Move (iSelectedIndex - 1)
        sValueToStore = clsMIGRateSet.GetMPMigFeeSet
    End If
    clsProductTable.SetMPMIGRateSet (sValueToStore)
    
    
    
    'GD BMIDS0014
    'Income Multiple Set Combo
    Dim clsIncomeMultipleSet As New IncomeMultipleSetTable
    clsIncomeMultipleSet.GetIncomeMultipleSet
    Set rs = TableAccess(clsIncomeMultipleSet).GetRecordSet
    'check if null value is in combo.
    sValueToStore = ""
    If frmProductDetails.cboIncomeMultipleSet.SelectedItem <> "" Then
        iSelectedIndex = frmProductDetails.cboIncomeMultipleSet.SelectedItem
        rs.Move (iSelectedIndex - 1)
        sValueToStore = clsIncomeMultipleSet.GetIncMultCode
    End If
    clsProductTable.SetIncomeMultiplierCode (sValueToStore)
    
    'JD BMIDS765
    'Rental Income Rate Set Combo
    Dim clsRentalIncomeRateSetBand As New RentalIncomeRateSetBandTable
    clsRentalIncomeRateSetBand.GetRentalIncomeRateSet
    Set rs = TableAccess(clsRentalIncomeRateSetBand).GetRecordSet
    'check if null value is in combo.
    sValueToStore = ""
    If frmProductDetails.cboRentalIncomeRateSet.SelectedItem <> "" Then
        iSelectedIndex = frmProductDetails.cboRentalIncomeRateSet.SelectedItem
        rs.Move (iSelectedIndex - 1)
        sValueToStore = clsRentalIncomeRateSetBand.GetRateSet
    End If
    clsProductTable.SetRentalIncomeRateSet (sValueToStore)
    
    'MP MIG Start LTV
    clsProductTable.SetMPMIGStartLTV frmProductDetails.txtMPMIGStartLTV.Text
    'HM MAR42
    clsProductTable.SetFreeChargePercentage frmProductDetails.txtERCFreePercentage.Text
    
    'Min Percent Arrangement Fee
    clsProductTable.SetMinPercentArrangementFee frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateRedemptionFeeSetCombo()
'GD BMIDS0011
    On Error GoTo Failed
    Dim clsRedemptionFeeSet As New RedemptionFeeSetTable
    Dim clsTableAccess As TableAccess
    
    clsRedemptionFeeSet.GetRedemptionFeeSet
    
    g_clsFormProcessing.PopulateDataCombo frmProductDetails.cboRedemptionFeeSet, clsRedemptionFeeSet
    Set clsTableAccess = clsRedemptionFeeSet
    
    'If clsTableAccess.RecordCount = 0 Then
        'g_clsErrorHandling.RaiseError errGeneralError, "No Redemption Fee Sets exist"
    'End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateMIGRateSetCombo()
'GD BMIDS0014
    On Error GoTo Failed
    Dim clsMIGRateSet As New MPMigRateSetTable

    Dim clsTableAccess As TableAccess
    
    clsMIGRateSet.GetMPMigRateSet
    
    g_clsFormProcessing.PopulateDataCombo frmProductDetails.cboMIGRateSet, clsMIGRateSet
    Set clsTableAccess = clsMIGRateSet
    
    'If clsTableAccess.RecordCount = 0 Then
        'g_clsErrorHandling.RaiseError errGeneralError, "No MIG Rate Sets exist"
    'End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'PopulateIncomeMultipleSetCombo
Private Sub PopulateIncomeMultipleSetCombo()
'GD BMIDS0014
    On Error GoTo Failed
    Dim clsIncomeMultipleSet As New IncomeMultipleSetTable
    Dim clsTableAccess As TableAccess
    
    clsIncomeMultipleSet.GetIncomeMultipleSet
    
    g_clsFormProcessing.PopulateDataCombo frmProductDetails.cboIncomeMultipleSet, clsIncomeMultipleSet
    Set clsTableAccess = clsIncomeMultipleSet
    
    'If clsTableAccess.RecordCount = 0 Then
    '   g_clsErrorHandling.RaiseError errGeneralError, "No Income Multiples Exist exist"
    'End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateRentalIncomeRateSetCombo()
'JD BMIDS765
    On Error GoTo Failed
    Dim clsRentalIncomeRateSetBand As New RentalIncomeRateSetBandTable
    
    clsRentalIncomeRateSetBand.GetRentalIncomeRateSet
    
    g_clsFormProcessing.PopulateDataCombo frmProductDetails.cboRentalIncomeRateSet, clsRentalIncomeRateSetBand
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidateClientScreenData() As Boolean
    On Error GoTo Failed
    Dim sFee As String
    Dim sFeePercentMax As String
    Dim sFeePercentMin As String
    Dim sFeePercent As String
    
    sFee = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE).Text
    sFeePercentMax = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MAX).Text
    sFeePercentMin = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN).Text
    sFeePercent = frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE).Text
    
    If Val(sFee) > 0 And Val(sFeePercent) > 0 Then
        g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE)
        g_clsErrorHandling.RaiseError errGeneralError, "Only an arrangement fee amount OR an arrangement fee percentage can be entered"
    End If
    
    If Len(sFee) > 0 And (Val(sFeePercentMin) > 0 Or Val(sFeePercentMax) > 0) Then
        g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN)
        g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum arrangement fee values can only be applied against the arrangement fee percentage"
    End If
    
    If Val(sFeePercent) = 0 And (Val(sFeePercentMin) > 0 Or Val(sFeePercentMax) > 0) Then
        g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN)
        g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum arrangement fee values cannot be specified when there is no arrangement fee percentage"
    End If
    
    If Len(sFeePercentMax) > 0 Then
        If Val(sFeePercentMin) > Val(sFeePercentMax) Then
            g_clsFormProcessing.SetControlFocus frmProductDetails.txtAdditionalParams(ARRANGEMENT_FEE_PERCENTAGE_MIN)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum arrangement fee must be greater than or equal to the minimum fee"
        End If
    End If
    
    ValidateClientScreenData = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

