VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageProductTable
' Description   : Contains all access routines and sql queries for the MortgageProduct table.
'                 Implements TableAccess
' Change history
' Prog      Date        Description
' DJP       11/06/01    SQL Server port
' DJP       04/08/01    DJP don't log fact that version fields don't exist
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem, removed
'                       GetLenderCodeField and GetLenderCodes.
' STB       30/01/02    SYS3771 ProductGroupingTable has been merged into this one.
' SDS       06/02/02    SYS2608 (13) In order to check whether mortgage products exist
'                       for the selected base rate , added new method DoesProductExistForBaseRate()
' STB       09/05/02    SYS4486 Amended so duplicate products are not displayed (when STARTDATEs are the same).
' BG        17/06/02    SYS4692 Added new method: DoesProductExistForThisItem
' BMIDS Change history
' Prog      Date        Description
' GD        10/05/2002  Client Specific : BMIDS Changes,
'                       Deuce Ref : BMIDS00002(MASTER) AND BMIDS00007(CHILD)
'                       For Requirement BM014("Can Be Ported") and BM019("Exclusive Or Semi Exclusive")
'
' GD        20/05/2002  Client Specific : BMIDS Changes,Deuce Ref : BMIDS0011
'                       For requirement BM017(part of)
' AW        21/05/02    BM087 - Added DoesProductExistForThisItem()
' GD        23/05/02    BMIDS00014
' AW        29/05/02    BMID00016 Changes to DoesProductExistForThisItem()
' GD        05/06/02    BMIDS00016 Changes to DeleteProduct
' MO        10/06/02    BMIDS00040 Change to MinPercentArrangementFee
' DB        31/01/03    SYS5724 Re-ordred the way that mortgage products are displayed
' DJP       06/03/03    BM0423 Validate the mortgage product when OK is pressed.
' BS        26/03/03    BM0311 Change order of mortgage product columns
' MC        26/05/2004  [MC]BMIDS775/REG016/018 Get/SET CATIndicator properties added.
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
'   JD      14/06/04    BMIDS765 CC076 added DoesProductExistForRentalIncomeRate
' MC        24/06/2004  BMIDS783 GetProducts() METHOD extended to give ProductCode for additional
'                       Search criteria.
' HMA       09/12/2004  BMIDS959  Remove MortgageProductBandsTable
' HM        23/08/2005  WP16 MAR42 - ERC Free Percentage is added on Other Rates/Groups Tab

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM  Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 - Added handling for Additional Borrowing Fee and Credit Limit Increase Fee
' GHun      17/11/2006  EP2_19 Added GetFreeLegalFees and SetFreeLegalFees
' TW        11/12/2006  EP2_20 Added handling for Transfer Of Equity Fee
' TW        23/02/2007  EP2_1354 - DBM183 Introduce refunded valuation products.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Constants
Private Const m_sTableName = "MORTGAGEPRODUCT"
Private Const m_sVersionField As String = "PRODUCTVERSIONNUMBER"

' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_bVersionExists As Boolean

' Enums
Public Enum SearchType
    SearchActive = 1
    SearchWithdrawn
    SearchLender
    SearchNone
End Enum



Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "MortgageProductCode"
    colFields.Add "StartDate"
    
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function

Friend Function GetMortgageProductCode() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MortgageProductCode")
    
    GetMortgageProductCode = vVal
End Function

Friend Function GetStartDate() As Variant
    GetStartDate = m_clsTableAccess.GetDate("StartDate", , , , True)
End Function

Friend Function GetEndDate() As Variant
    
    GetEndDate = m_clsTableAccess.GetDate("EndDate")
End Function

Friend Function GetCountryOfOrigin() As Variant
    GetCountryOfOrigin = m_clsTableAccess.GetVal("CountryOfOrigin")
End Function

Friend Function GetOrganisationID() As Variant
    GetOrganisationID = m_clsTableAccess.GetVal("OrganisationID")
End Function

Friend Function GetNonPanelLender() As Variant
    GetNonPanelLender = m_clsTableAccess.GetVal("NONPANELLENDEROPTION")
End Function

Friend Function GetPortingFee() As Variant
    GetPortingFee = m_clsTableAccess.GetVal("PortingFee")
End Function

Friend Function GetDisplayOrder() As Variant
    GetDisplayOrder = m_clsTableAccess.GetVal("DisplayOrder")
End Function

Friend Function GetArrangementFeeAmount() As Variant
    GetArrangementFeeAmount = m_clsTableAccess.GetVal("MPArrangementFeeAmount")
End Function

Friend Function GetArrangementFeePercent() As Variant
    GetArrangementFeePercent = m_clsTableAccess.GetVal("MPArrangementFeePercent")
End Function

Friend Function GetInterestRateDecreaseLimit(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As Variant
    ' DJP MCAP
    GetInterestRateDecreaseLimit = m_clsTableAccess.GetData("ManualDecreaseLimitPercent", ctrl, ctrlLabel)
    'GetInterestRateDecreaseLimit = m_clsTableAccess.GetVal("ManualDecreaseLimitPercent")
End Function

Friend Function GetArrangementFeePercentMax() As Variant
    GetArrangementFeePercentMax = m_clsTableAccess.GetVal("MPArrangementFeePercentMax")
End Function

Friend Function GetCashbackProduct() As Variant
    GetCashbackProduct = m_clsTableAccess.GetVal("CashbackProduct")
End Function

Friend Function GetAvailableForImpairedCredit() As Variant
    GetAvailableForImpairedCredit = m_clsTableAccess.GetVal("ImpairedCredit")
End Function

Friend Function GetAvailableForMortgageCalculator() As Variant
    GetAvailableForMortgageCalculator = m_clsTableAccess.GetVal("AvailableForMortgageCalculator")
End Function

Friend Function GetAvailableForSelfEmployed() As Variant
    GetAvailableForSelfEmployed = m_clsTableAccess.GetVal("AvailableForSelfEmployed")
End Function

'EP2_19 GHun
Friend Function GetFreeLegalFees() As Variant
    GetFreeLegalFees = m_clsTableAccess.GetVal("FreeLegalFees")
End Function
'EP2_19 End

Friend Function GetAvailableQuickQuote() As Variant
    GetAvailableQuickQuote = m_clsTableAccess.GetVal("AvailableQuickQuote")
End Function

Friend Function GetExistCustomerMortgageProduct() As Variant
    GetExistCustomerMortgageProduct = m_clsTableAccess.GetVal("ExistCustomerMortgageProduct")
End Function

Friend Function GetMaxApplicantAge() As Variant
    GetMaxApplicantAge = m_clsTableAccess.GetVal("MaxApplicantAge")
End Function

Friend Function GetMaximumTerm() As Variant
    GetMaximumTerm = m_clsTableAccess.GetVal("MaximumTerm")
End Function

Friend Function GetMinApplicantAge() As Variant
    GetMinApplicantAge = m_clsTableAccess.GetVal("MinApplicantAge")
End Function

Friend Function GetMinimumTerm() As Variant
    GetMinimumTerm = m_clsTableAccess.GetVal("MinimumTerm")
End Function

Friend Function GetValidProduct() As String
    On Error GoTo Failed
    Dim sVal As String
    Dim sValid As String
    sVal = m_clsTableAccess.GetVal("VALIDPRODUCT")

    sValid = "False"
    If Len(sVal) > 0 Then
        If CLng(sVal) = 1 Then
            sValid = "True"
        End If
    End If
    
    GetValidProduct = sValid
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Friend Function GetPreferentialRatePeriod() As Variant
    GetPreferentialRatePeriod = m_clsTableAccess.GetVal("PreferentialRatePeriod")
End Function

Friend Function GetStaffMortgageProduct() As Variant
    GetStaffMortgageProduct = m_clsTableAccess.GetVal("StaffMortgageProduct")
End Function

Friend Function GetWithdrawnDate() As Variant
    GetWithdrawnDate = m_clsTableAccess.GetDate("WithdrawnDate")
End Function

Friend Function GetAdminFeeSet() As Variant
    GetAdminFeeSet = m_clsTableAccess.GetVal("ADMINISTRATIONFEESET")
End Function

' TW 23/02/2007 EP2_1354
Friend Function GetRefundOfValuation() As Variant
    GetRefundOfValuation = m_clsTableAccess.GetVal("REFUNDOFVALUATION")
End Function
' TW 23/02/2007 EP2_1354 End

Friend Function GetTransferOfEquityFeeSet() As Variant
    GetTransferOfEquityFeeSet = m_clsTableAccess.GetVal("TRANSFEROFEQUITYFEESET")
End Function

Friend Function GetAdditionalBorrowingFeeSet() As Variant
    GetAdditionalBorrowingFeeSet = m_clsTableAccess.GetVal("ADDITIONALBORROWINGFEESET")
End Function

Friend Function GetCreditLimitIncreaseFeeSet() As Variant
    GetCreditLimitIncreaseFeeSet = m_clsTableAccess.GetVal("CREDITLIMITINCREASEFEESET")
End Function


'*=[MC]BMIDS763 - CC075 FEE SETS
Friend Function GetTTFeeSet() As Variant
    GetTTFeeSet = m_clsTableAccess.GetVal("TTFeeSet")
End Function

Friend Function GetIAFeeSet() As Variant
    GetIAFeeSet = m_clsTableAccess.GetVal("InsuranceAdminFeeSet")
End Function

Friend Function GetProductSwitchFeeSet() As Variant
    GetProductSwitchFeeSet = m_clsTableAccess.GetVal("ProductSwitchFeeSet")
End Function
'*=SECTION END

Friend Function GetBaseRateFeeSet() As Variant
    GetBaseRateFeeSet = m_clsTableAccess.GetVal("BaseRateSet")
End Function

Friend Function GetValuationFeeSet() As Variant
    GetValuationFeeSet = m_clsTableAccess.GetVal("ValuationFeeSet")
End Function

Friend Function GetFeeSet() As Variant
    GetFeeSet = m_clsTableAccess.GetVal("FeeSet")
End Function

Friend Function GetMinimumLoanAmount() As Variant
    GetMinimumLoanAmount = m_clsTableAccess.GetVal("MinimumLoanAmount")
End Function

Friend Function GetMaximumLoanAmount() As Variant
    GetMaximumLoanAmount = m_clsTableAccess.GetVal("MaximumLoanAmount")
End Function

Friend Function GetMinimumLTV() As Variant
    GetMinimumLTV = m_clsTableAccess.GetVal("MinimumLTV")
End Function

Friend Function GetMaximumLTV() As Variant
    GetMaximumLTV = m_clsTableAccess.GetVal("MaximumLTV")
End Function

Friend Function GetFlexibleMortgageProduct() As Variant
    GetFlexibleMortgageProduct = m_clsTableAccess.GetVal("FlexibleMortgageProduct")
End Function

Friend Function GetProductName() As Variant
    GetProductName = m_clsTableAccess.GetVal("ProductName")
End Function

Friend Function GetLenderCode() As Variant
    GetLenderCode = m_clsTableAccess.GetVal("LenderCode")
End Function

Friend Function GetProductTextDetails() As Variant
    GetProductTextDetails = m_clsTableAccess.GetVal("ProductTextDetails")
End Function

Public Function GetProductBrandName() As Variant
    GetProductBrandName = m_clsTableAccess.GetVal("BRANDID")
End Function

Public Function GetProductGroup() As Variant
    GetProductGroup = m_clsTableAccess.GetVal("PRODUCTGROUP")
End Function

Public Function GetProductLine() As Variant
    GetProductLine = m_clsTableAccess.GetVal("PRODUCTLINE")
End Function

Friend Function SetMortgageProductCode(sVal As Variant) As Boolean
    SetMortgageProductCode = m_clsTableAccess.SetVal("MortgageProductCode", CVar(sVal))
    
    ' Versioning, if it exists
    g_clsVersion.HandleVersioning Me
End Function

Friend Function SetCountryOfOrigin(vCountry As Variant) As Boolean
    SetCountryOfOrigin = m_clsTableAccess.SetVal("CountryOfOrigin", vCountry)
End Function

Friend Function SetNonPanelLender(sPanelLender As Variant) As Boolean
    SetNonPanelLender = m_clsTableAccess.SetVal("NONPANELLENDEROPTION", sPanelLender)
End Function

Friend Function SetCashbackProduct(sCashback As Variant) As Boolean
    SetCashbackProduct = m_clsTableAccess.SetVal("CashbackProduct", sCashback)
End Function

Friend Function SetInterestRateDecreaseLimit(sVal As Variant) As Boolean
    SetInterestRateDecreaseLimit = m_clsTableAccess.SetVal("ManualDecreaseLimitPercent", CVar(sVal))
End Function

Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetDate("StartDate", CVar(sVal))
End Function

Friend Function SetOrganisationID(sVal As Variant) As Boolean
    SetOrganisationID = m_clsTableAccess.SetVal("OrganisationID", CVar(sVal))
End Function

Friend Function SetArrangementFeeAmount(sVal As Variant) As Boolean
    SetArrangementFeeAmount = m_clsTableAccess.SetVal("MPArrangementFeeAmount", CVar(sVal))
End Function

Friend Function SetArrangementFeePercent(sVal As Variant) As Boolean
    SetArrangementFeePercent = m_clsTableAccess.SetVal("MPArrangementFeePercent", CVar(sVal))
End Function

Friend Function SetArrangementFeePercentMax(sVal As Variant) As Boolean
    SetArrangementFeePercentMax = m_clsTableAccess.SetVal("MPArrangementFeePercentMax", CVar(sVal))
End Function

Friend Function SetAvailableForMortgageCalculator(sVal As Variant) As Boolean
    SetAvailableForMortgageCalculator = m_clsTableAccess.SetVal("AvailableForMortgageCalculator", CVar(sVal))
End Function

Friend Function SetAvailableForImpairedCredit(sVal As Variant) As Boolean
    SetAvailableForImpairedCredit = m_clsTableAccess.SetVal("ImpairedCredit", CVar(sVal))
End Function

Friend Function SetAvailableForSelfEmployed(sVal As Variant) As Boolean
    SetAvailableForSelfEmployed = m_clsTableAccess.SetVal("AvailableForSelfEmployed", CVar(sVal))
End Function

'EP2_19 GHun
Friend Function SetFreeLegalFees(sVal As Variant) As Boolean
    SetFreeLegalFees = m_clsTableAccess.SetVal("FreeLegalFees", CVar(sVal))
End Function
'EP2_19 End

Friend Function SetAvailableQuickQuote(sVal As Variant) As Boolean
    SetAvailableQuickQuote = m_clsTableAccess.SetVal("AvailableQuickQuote", CVar(sVal))
End Function

Friend Function SetEndDate(sVal As Variant) As Boolean
    SetEndDate = m_clsTableAccess.SetDate("EndDate", CVar(sVal))
End Function

Friend Function SetExistCustomerMortgageProduct(sVal As Variant) As Boolean
    SetExistCustomerMortgageProduct = m_clsTableAccess.SetVal("ExistCustomerMortgageProduct", CVar(sVal))
End Function

Friend Function SetMaxApplicantAge(sVal As Variant) As Boolean
    SetMaxApplicantAge = m_clsTableAccess.SetVal("MaxApplicantAge", CVar(sVal))
End Function

Friend Function SetMaximumTerm(sVal As Variant) As Boolean
    SetMaximumTerm = m_clsTableAccess.SetVal("MaximumTerm", CVar(sVal))
End Function

Friend Function SetMinApplicantAge(sVal As Variant) As Boolean
    SetMinApplicantAge = m_clsTableAccess.SetVal("MinApplicantAge", CVar(sVal))
End Function

Friend Function SetMinimumTerm(sVal As Variant) As Boolean
    SetMinimumTerm = m_clsTableAccess.SetVal("MinimumTerm", CVar(sVal))
End Function

Friend Function SetPortingFee(sVal As Variant) As Boolean
    SetPortingFee = m_clsTableAccess.SetVal("PortingFee", CVar(sVal))
End Function

Friend Function SetPreferentialRatePeriod(sVal As Variant) As Boolean
    SetPreferentialRatePeriod = m_clsTableAccess.SetVal("PreferentialRatePeriod", CVar(sVal))
End Function

Friend Function SetStaffMortgageProduct(sVal As Variant) As Boolean
    SetStaffMortgageProduct = m_clsTableAccess.SetVal("StaffMortgageProduct", CVar(sVal))
End Function

Friend Function SetWithdrawnDate(sVal As Variant) As Boolean
    SetWithdrawnDate = m_clsTableAccess.SetDate("WithdrawnDate", CVar(sVal))
End Function

Friend Function SetAdminFeeSet(sVal As Variant) As Boolean
    SetAdminFeeSet = m_clsTableAccess.SetVal("ADMINISTRATIONFEESET", CVar(sVal))
End Function

' TW 23/02/2007 EP2_1354
Friend Function SetRefundOfValuation(sVal As Variant) As Boolean
    SetRefundOfValuation = m_clsTableAccess.SetVal("REFUNDOFVALUATION", CVar(sVal))
End Function
' TW 23/02/2007 EP2_1354 End

Friend Function SetTransferOfEquityFeeSet(sVal As Variant) As Boolean
    SetTransferOfEquityFeeSet = m_clsTableAccess.SetVal("TRANSFEROFEQUITYFEESET", CVar(sVal))
End Function

Friend Function SetAdditionalBorrowingFeeSet(sVal As Variant) As Boolean
    SetAdditionalBorrowingFeeSet = m_clsTableAccess.SetVal("ADDITIONALBORROWINGFEESET", CVar(sVal))
End Function

Friend Function SetCreditLimitIncreaseFeeSet(sVal As Variant) As Boolean
    SetCreditLimitIncreaseFeeSet = m_clsTableAccess.SetVal("CREDITLIMITINCREASEFEESET", CVar(sVal))
End Function


'*=[MC]BMIDS763
Friend Function SetInsuranceAdminFeeSet(sVal As Variant) As Boolean
    SetInsuranceAdminFeeSet = m_clsTableAccess.SetVal("InsuranceAdminFeeSet", CVar(sVal))
End Function

Friend Function SetTTFeeSet(sVal As Variant) As Boolean
    SetTTFeeSet = m_clsTableAccess.SetVal("TTFeeSet", CVar(sVal))
End Function

Friend Function SetProductSwitchFeeSet(sVal As Variant) As Boolean
    SetProductSwitchFeeSet = m_clsTableAccess.SetVal("ProductSwitchFeeSet", CVar(sVal))
End Function
'*=SECTION END


Friend Function SetBaseRateFeeSet(sVal As Variant) As Boolean
    SetBaseRateFeeSet = m_clsTableAccess.SetVal("BaseRateSet", CVar(sVal))
End Function

Friend Function SetValuationFeeSet(sVal As Variant) As Boolean
    SetValuationFeeSet = m_clsTableAccess.SetVal("ValuationFeeSet", CVar(sVal))
End Function

Friend Function SetFeeSet(sVal As Variant) As Boolean
    SetFeeSet = m_clsTableAccess.SetVal("FeeSet", CVar(sVal))
End Function

Friend Function SetMinimumLoanAmount(sVal As Variant) As Boolean
    SetMinimumLoanAmount = m_clsTableAccess.SetVal("MinimumLoanAmount", CVar(sVal))
End Function

Friend Function SetMaximumLoanAmount(sVal As Variant) As Boolean
    SetMaximumLoanAmount = m_clsTableAccess.SetVal("MaximumLoanAmount", CVar(sVal))
End Function

Friend Function SetMinimumLTV(sVal As Variant) As Boolean
    SetMinimumLTV = m_clsTableAccess.SetVal("MinimumLTV", CVar(sVal))
End Function

Friend Function SetMaximumLTV(sVal As Variant) As Boolean
    SetMaximumLTV = m_clsTableAccess.SetVal("MaximumLTV", CVar(sVal))
End Function

Friend Function SetFlexibleMortgageProduct(sVal As Variant) As Boolean
    SetFlexibleMortgageProduct = m_clsTableAccess.SetVal("FlexibleMortgageProduct", CVar(sVal))
End Function

Friend Function SetProductName(sVal As Variant) As Boolean
    SetProductName = m_clsTableAccess.SetVal("ProductName", CVar(sVal))
End Function

Friend Function SetProductTextDetails(sVal As Variant) As Boolean
    SetProductTextDetails = m_clsTableAccess.SetVal("ProductTextDetails", CVar(sVal))
End Function

Friend Function SetDisplayOrder(sVal As Variant) As Boolean
    SetDisplayOrder = m_clsTableAccess.SetVal("DisplayOrder", CVar(sVal))
End Function

Friend Function SetProductBrandName(sVal As Variant) As Boolean
    SetProductBrandName = m_clsTableAccess.SetVal("BRANDID", CVar(sVal))
End Function

Friend Function SetProductGroup(sVal As Variant) As Boolean
    SetProductGroup = m_clsTableAccess.SetVal("PRODUCTGROUP", CVar(sVal))
End Function

Friend Function SetProductLine(sVal As Variant) As Boolean
    SetProductLine = m_clsTableAccess.SetVal("PRODUCTLINE", CVar(sVal))
End Function

Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_clsTableAccess.GetIsDeleted()
End Function

Private Sub TableAccess_SetDeleteFields(colFields As Collection)
    m_clsTableAccess.SetDeleteFields colFields
End Sub

Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_clsTableAccess.GetTableDataFromSearch sSearch
End Sub

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function

Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub

Private Sub TableAccess_DeleteAllRows()
    m_clsTableAccess.DeleteAllRows
End Sub

Private Sub TableAccess_DeleteRow(colMatchData As Collection)
    m_clsTableAccess.DeleteRow colMatchData
End Sub

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function

Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function

Private Function TableAccess_GetKeyMatchFields()
    Set TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function

Private Function TableAccess_GetKeyMatchValues() As Variant
    Set TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function

Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function

Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function

Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function

Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function

Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
        m_clsTableAccess.ValidateData sDesc
End Sub

Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function

Private Sub TableAccess_MoveFirst()
    m_clsTableAccess.MoveFirst
End Sub

Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
End Function

Private Function TableAccess_RecordCount() As Integer
    TableAccess_RecordCount = m_clsTableAccess.RecordCount()
End Function

Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vVal)
End Function

Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)
    m_clsTableAccess.SetKeyMatchValues colValues
End Sub

Private Function TableAccess_SetNextRow() As Boolean
    TableAccess_SetNextRow = m_clsTableAccess.SetNextRow()
End Function

Private Sub TableAccess_SetPopulateType(enumType As PopulateType)

End Sub

Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    m_clsTableAccess = clsTable
End Sub

Private Sub TableAccess_SetUpdated()
    m_clsTableAccess.SetUpdated
End Sub

Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vVal)
End Function

Private Sub TableAccess_Update()
     m_clsTableAccess.Update
End Sub

Private Sub TableAccess_SetPopulateNull()
    m_clsTableAccess.SetPopulateNull
End Sub

Private Function TableAccess_GetSearch() As String
    TableAccess_GetSearch = m_sSearch
End Function

Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function

Public Function TableAccess_GetUpdateValues() As Collection
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    Dim sProductValid As String
    'BS BM0311 26/03/03
    colListLine.Add GetMortgageProductCode()
    colListLine.Add GetStartDate()
    colListLine.Add GetEndDate()
    'BS BM0311 26/03/03
    'colListLine.Add GetMortgageProductCode()
    colListLine.Add GetProductName()
    colListLine.Add GetLenderCode()

    If Not g_clsVersion.DoesVersioningExist() Then
        sProductValid = GetValidProduct()
        colListLine.Add sProductValid
        colListLine.Add sProductValid, OBJECT_EXTRA_VALUE
    End If
End Sub

'*=BMIDS783 - GetProducts() extended to include ProductCode
Public Sub GetProducts(Optional enumSearch As SearchType = SearchNone, Optional sLenderCode As String, Optional sEndDate As String, Optional sName As String, Optional sProductCode As String)
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim sExtra As String
    Dim sSysDate As String
    On Error GoTo Failed
    'Need to fix this to get the version number too, or the product with the version we want
    sSysDate = g_clsSQLAssistSP.GetSystemDate
    
    sSearch = "SELECT " & _
                    "MP.STARTDATE, " & _
                    "MP.ENDDATE, " & _
                    "MP.MORTGAGEPRODUCTCODE, " & _
                    "MP.ORGANISATIONID, " & _
                    "ML.LENDERCODE, " & _
                    "MPL.PRODUCTNAME "
        
    If m_bVersionExists Then
        sSearch = sSearch + ", MP." + m_sVersionField & " "
    Else
        sSearch = sSearch + ", MP.VALIDPRODUCT "
    End If
    
    'SYS4486 - Amended so duplicate products are not displayed (when STARTDATEs are the same).
    sSearch = sSearch + "FROM " & _
                            "MORTGAGEPRODUCT MP, " & _
                            "MORTGAGELENDER ML, " & _
                            "MORTGAGEPRODUCTLANGUAGE MPL " & _
                        "WHERE " & _
                            "MP.ORGANISATIONID = ML.ORGANISATIONID AND " & _
                            "MPL.MORTGAGEPRODUCTCODE = MP.MORTGAGEPRODUCTCODE AND " & _
                            "MPL.STARTDATE = MP.STARTDATE"
    
    If m_bVersionExists Then
        ' Add versioning to the search
        sSearch = sSearch + " AND MP." + m_sVersionField + " = " & g_sVersionNumber & _
                            " AND ML.MORTGAGELENDERVERSIONNUMBER = " & g_sVersionNumber & _
                            " AND MPL." + m_sVersionField + " = " & g_sVersionNumber
    End If
    
    Select Case enumSearch
        Case SearchActive
            ' DJP SQL Server port
          sSearch = sSearch + " AND " + g_clsSQLAssistSP.TruncateDateColumn("MP.STARTDATE") + " <= " + g_clsSQLAssistSP.TruncateDateColumn(sSysDate) + _
          " AND " + g_clsSQLAssistSP.TruncateDateColumn("MP.ENDDATE") + " >= " + g_clsSQLAssistSP.TruncateDateColumn(sSysDate) + _
          " AND (" + g_clsSQLAssistSP.TruncateDateColumn("MP.WITHDRAWNDATE") + " >= " + g_clsSQLAssistSP.TruncateDateColumn(sSysDate) + _
          " OR MP.WITHDRAWNDATE IS NULL)"
        
        Case SearchWithdrawn
          sSearch = sSearch + " AND " + g_clsSQLAssistSP.TruncateDateColumn("MP.WITHDRAWNDATE") + " <= " + g_clsSQLAssistSP.TruncateDateColumn(sSysDate)
            
    End Select
    
    If Len(sEndDate) > 0 Then
        ' DJP SQL Server port
        sSearch = sSearch + " AND " + g_clsSQLAssistSP.TruncateDateColumn("MP.ENDDATE") + " >= " & g_clsSQLAssistSP.FormatDate(sEndDate)
    End If
    
    If Len(sLenderCode) > 0 Then
        sSearch = sSearch + " AND ML.LENDERCODE = " + g_clsSQLAssistSP.FormatString(sLenderCode)
    End If
    
    If Len(sName) > 0 Then
        Dim bTmp As Boolean
        sSearch = sSearch + " AND MPL.PRODUCTNAME LIKE " + g_clsFormProcessing.FormatWildcardString(sName)
    Else
        '*=[MC]ProductCode added to search Criteria.
        If Len(Trim(sProductCode)) > 0 Then
            sSearch = sSearch + " AND MPL.MORTGAGEPRODUCTCODE LIKE " + g_clsFormProcessing.FormatWildcardString(sProductCode)
        End If
    End If
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, "Products"
    m_clsTableAccess.SetRecordSet rs

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub GetOrgIDFromLenderCode(sLenderCode As String, vOrgID As Variant)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String
    Dim clsLender As New MortgageLendersTable
    Dim clsTableAccess As TableAccess
    Dim sLenderVersionField As String
    
    Set clsTableAccess = clsLender
    sTable = clsTableAccess.GetTable()
    sLenderVersionField = clsLender.GetVersionField()
    
    sSearch = "select OrganisationID from " + sTable + " where LenderCode = '" + FormatString(sLenderCode) + "'"
    
    If m_bVersionExists Then
        sSearch = sSearch + " AND " + sLenderVersionField + " = " & g_sVersionNumber
    End If
    
    Set rs = g_clsDataAccess.GetTableData(sTable, sSearch)
    
    ValidateRecordset rs, "MortgageLender"
    
    If rs.RecordCount = 1 Then
        ' DJP SQL Server port
        vOrgID = rs("OrganisationID")
    
        If Len(vOrgID) = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "GetOrgIDFromLenderCode - OrganisationID is empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetOrgIDFromLenderCode - More than one Record for OrgID"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Function DoesProductExist(sProductCode As String, sDate As String) As Boolean
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String

    On Error GoTo Failed

    bRet = False

    ' DJP SQL Server port
    sSearch = "select MortgageProductCode from " + m_sTableName + " where MortgageProductCode = " + g_clsSQLAssistSP.FormatString(sProductCode) + " AND " + _
    g_clsSQLAssistSP.TruncateDateColumn("STARTDATE") + " = " + g_clsSQLAssistSP.FormatDate(sDate)

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
        End If
    End If
    
    DoesProductExist = bRet

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Friend Function DoesProductExistForBaseRate(sBaseRateSet As String) As Collection
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String
    Dim colMortgageProducts As New Collection

    On Error GoTo Failed

    bRet = False

    ' DJP SQL Server port
    sSearch = "select MortgageProductCode from " + m_sTableName + " where BaseRateSet = " + g_clsSQLAssistSP.FormatString(sBaseRateSet)
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
            Do While Not (rs.EOF)
                colMortgageProducts.Add rs("MortgageProductCode").Value
                rs.MoveNext
            Loop
        Else
            Set colMortgageProducts = Nothing
        End If
        
    End If
    
    Set DoesProductExistForBaseRate = colMortgageProducts 'bRet

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function DoesProductExistForRentalIncomeRate(sRateSet As String) As Collection
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String
    Dim colMortgageProducts As New Collection

    On Error GoTo Failed

    bRet = False

    ' JD Function added BMIDS765
    sSearch = "select MortgageProductCode from " + m_sTableName + " where RentalIncomeRateSet = " + g_clsSQLAssistSP.FormatString(sRateSet)
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
            Do While Not (rs.EOF)
                colMortgageProducts.Add rs("MortgageProductCode").Value
                rs.MoveNext
            Loop
        Else
            Set colMortgageProducts = Nothing
        End If
        
    End If
    
    Set DoesProductExistForRentalIncomeRate = colMortgageProducts 'bRet

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub DeleteProduct()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim sTargetDatabase As String
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues()
    sTargetDatabase = m_clsTableAccess.GetDatabase()
    
    ' First, Inclusive and Exclusive incentives
    Dim clsIncentive As New MortProdIncentiveTable
    Set clsTableAccess = clsIncentive
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsIncentive.DeleteIncentives sTargetDatabase
    
    ' BMIDS959  Remove Mortgage Product Bands Table
    
    ' Mortgage Product Language table
    Set clsTableAccess = New MortProdLanguageTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Mortgage Product Purpose of Loan table
    Set clsTableAccess = New MortProdPurpOfLoanTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Property Location
    Set clsTableAccess = New MortProdPropLocTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Mortgage Product Type of Buyer table
    Set clsTableAccess = New MortProdTypeOfBuyerTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Mortgage Product Channel Eligibility table
    Set clsTableAccess = New MortProdChannelEligTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Mortgage Product additional parameters table
    Set clsTableAccess = New MortProdParamsTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    
    ' Interest rates table
    Dim clsIntRates As New MortProdIntRatesTable
    Set clsTableAccess = clsIntRates
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsIntRates.DeleteInterestRates sTargetDatabase

    ' Employment eligibility
    Set clsTableAccess = New MortProdEmpEligTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

    ' Typical APR's
    Set clsTableAccess = New MortProdTypicalAPRTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

    ' Type of application eligibility
    Set clsTableAccess = New MortProdTypeofAppEligTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

    ' Special group
    Set clsTableAccess = New MortProdSpecialGroupTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

    ' Other Fees
    Set clsTableAccess = New OtherFeeTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

'GD BMIDS00016 - Promotions
    'Mortgage Product Condition
    Set clsTableAccess = New MortProdProdCondTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase



    ' Last, the Mortgage Product table
    m_clsTableAccess.DeleteRecords sTargetDatabase
    
    
    'g_clsDataAccess.CommitTrans
    
    Exit Sub
Failed:
    'g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub

Public Function TableAccess_GetDatabase() As String
    TableAccess_GetDatabase = m_clsTableAccess.GetDatabase()
End Function

Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_clsTableAccess.SetOrderCriteria sCriteria
End Sub

Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub

Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function

Private Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub

Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

Public Sub GetProductByID(sProductID As String)
    On Error GoTo Failed

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim sSysDate As String
    
    sSysDate = g_clsSQLAssistSP.GetSystemDate
    
    sSQL = "SELECT " & _
                    "MP.STARTDATE, " & _
                    "MP.ENDDATE, " & _
                    "MP.MORTGAGEPRODUCTCODE, " & _
                    "MP.ORGANISATIONID, " & _
                    "ML.LENDERCODE, " & _
                    "MPL.PRODUCTNAME "
        
    If m_bVersionExists Then
        sSQL = sSQL + ", MP." + m_sVersionField & " "
    Else
        sSQL = sSQL + ", MP.VALIDPRODUCT "
    End If
    
    sSQL = sSQL + "FROM " & _
                            "MORTGAGEPRODUCT MP, " & _
                            "MORTGAGELENDER ML, " & _
                            "MORTGAGEPRODUCTLANGUAGE MPL " & _
                        "WHERE " & _
                            "MP.ORGANISATIONID = ML.ORGANISATIONID AND " & _
                            "MPL.MORTGAGEPRODUCTCODE = MP.MORTGAGEPRODUCTCODE AND " & _
                            "MP.MORTGAGEPRODUCTCODE = " & g_clsSQLAssistSP.FormatString(sProductID)
        
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetCanBePorted() As Variant
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    GetCanBePorted = m_clsTableAccess.GetVal("MORTGAGEPRODUCTPORTABLEIND")
End Function

Public Function SetCanBePorted(sVal As Variant) As Boolean
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    SetCanBePorted = m_clsTableAccess.SetVal("MORTGAGEPRODUCTPORTABLEIND", CVar(sVal))
End Function


'*=[MC]BMIDS775/REG016/018
Public Function GetCATIndicator() As Variant
    '*=[MC]Reg016 BMIDS0775
    GetCATIndicator = m_clsTableAccess.GetVal("CATIND")
End Function
'*=[MC]BMIDS775/REG016/018
Public Function SetCATIndicator(sVal As Variant) As Boolean
    '*=[MC]Reg016 BMIDS0775
    SetCATIndicator = m_clsTableAccess.SetVal("CATIND", CVar(sVal))
End Function



Public Function SetExclusiveOrSemi(sVal As Variant) As Boolean
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    SetExclusiveOrSemi = m_clsTableAccess.SetVal("ExclusiveInd", CVar(sVal))
End Function
Public Function GetExclusiveOrSemi() As Variant
    'GD 10/05/2002  Client Specific : BMIDS Changes : BMIDS00002(MASTER) BMIDS00007(CHILD)
    GetExclusiveOrSemi = m_clsTableAccess.GetVal("ExclusiveInd")
End Function
Public Function GetRedemptionFeeSet() As Variant
'GD BMIDS0011
    GetRedemptionFeeSet = m_clsTableAccess.GetVal("RedemptionFeeSet")
End Function
Public Function SetRedemptionFeeSet(sVal As Variant) As Boolean
'GD BMIDS0011
    SetRedemptionFeeSet = m_clsTableAccess.SetVal("RedemptionFeeSet", CVar(sVal))
End Function
Public Function GetMPMigRateSet() As Variant
'GD BMIDS0014
    GetMPMigRateSet = m_clsTableAccess.GetVal("MPMIGRateSet")
End Function
Public Function SetMPMIGRateSet(sVal As Variant) As Boolean
'GD BMIDS0014
    SetMPMIGRateSet = m_clsTableAccess.SetVal("MPMIGRateSet", CVar(sVal))
End Function
Public Function GetRentalIncomeRateSet() As Variant
'JD BMIDS765
    GetRentalIncomeRateSet = m_clsTableAccess.GetVal("RentalIncomeRateSet")
End Function
Public Function SetRentalIncomeRateSet(sVal As Variant) As Boolean
'JD BMIDS765
    SetRentalIncomeRateSet = m_clsTableAccess.SetVal("RentalIncomeRateSet", CVar(sVal))
End Function
Public Function GetIncomeMultiplierCode() As Variant
'GD BMIDS0014
    GetIncomeMultiplierCode = m_clsTableAccess.GetVal("IncomeMultiplierCode")
End Function
Public Function SetIncomeMultiplierCode(sVal As Variant) As Boolean
'GD BMIDS0014
    SetIncomeMultiplierCode = m_clsTableAccess.SetVal("IncomeMultiplierCode", CVar(sVal))
End Function
Public Function GetMPMIGStartLTV() As Variant
'GD BMIDS0014
    GetMPMIGStartLTV = m_clsTableAccess.GetVal("MPMIGStartLTV")
End Function
Public Function SetMPMIGStartLTV(sVal As Variant) As Boolean
'GD BMIDS0014
    SetMPMIGStartLTV = m_clsTableAccess.SetVal("MPMIGStartLTV", CVar(sVal))
End Function
Public Function GetFreeChargePercentage() As Variant
'HM MAR42
    GetFreeChargePercentage = m_clsTableAccess.GetVal("FreeChargePercentage")
End Function
Public Function SetFreeChargePercentage(sVal As Variant) As Boolean
'HM MAR42
    SetFreeChargePercentage = m_clsTableAccess.SetVal("FreeChargePercentage", CVar(sVal))
End Function
Public Function SetMinPercentArrangementFee(sVal As Variant) As Boolean
'GD BMIDS0014
    SetMinPercentArrangementFee = m_clsTableAccess.SetVal("MPARRANGEMENTFEEPERCENTMIN", CVar(sVal))
End Function
Public Function GetMinPercentArrangementFee() As Variant
'GD BMIDS0014
    GetMinPercentArrangementFee = m_clsTableAccess.GetVal("MPARRANGEMENTFEEPERCENTMIN")
End Function
Friend Function DoesProductExistForThisItem(sItem As String, sValue As String, _
                                            Optional sConnection As String) As Collection
                                            
    Dim bConnectionChanged As Boolean
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String, sCurrConnection As String
    Dim colMortgageProducts As New Collection

    On Error GoTo Failed

    bConnectionChanged = False

    If Len(sConnection) > 0 Then
        sCurrConnection = g_clsDataAccess.GetConnectionKey
        g_clsDataAccess.SetActiveConnection (sConnection)
        bConnectionChanged = True
    End If
    
    sSearch = "select MortgageProductCode from " + m_sTableName + " where " + sItem + " = " + g_clsSQLAssistSP.FormatString(sValue)
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then

            Do While Not (rs.EOF)
                colMortgageProducts.Add rs("MortgageProductCode").Value
                rs.MoveNext
            Loop
        Else
            Set colMortgageProducts = Nothing
        End If
        
    End If
    
    ' Re-establish  old connection
    If bConnectionChanged = True Then
        g_clsDataAccess.SetActiveConnection (sCurrConnection)
    End If
    
    Set DoesProductExistForThisItem = colMortgageProducts

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' DJP BM0423
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateProduct
' Description   : Calls the ValidateProduct Stored Procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ValidateProduct(Optional sTargetDatabase As String)
    Dim sDate As String
    Dim sProductCode As String
    Dim colMatchValues As Collection
    Dim Param As ADODB.Parameter
    Dim colParam As Collection
    Dim cmd As ADODB.Command
    Dim sCurrentDatabase As String
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues
    
    ' Product code is always first
    sProductCode = colMatchValues(1)
    sDate = colMatchValues(2)
    
    Set colParam = New Collection

    Set cmd = g_clsDataAccess.GetActiveCommand()

    Set Param = New ADODB.Parameter
    With Param
        .Name = "ProductCode"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = sProductCode
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With
    
    colParam.Add Param
    
    Set Param = New ADODB.Parameter
    With Param
        .Name = "StartDate"
        .Type = adDate
        .Direction = adParamInput
        .Value = sDate
        ' SQL Server requires the length
    End With
    
    colParam.Add Param
    
    g_clsDataAccess.ExecuteStoredProcedure spValidateProduct, colParam, sTargetDatabase
    
    Exit Sub
Failed:
    
End Sub
