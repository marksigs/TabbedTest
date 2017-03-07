VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageLendersTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortgageLendersTable
' Description   : Contains all data access routines for the MortgageLender table
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port
' AA        27/06/01    Added ability to order Lenders in GetLenders
' DJP       04/08/01    DJP don't log fact that version fields don't exist
' DJP       23/10/01    SYS2831 - change Lender to support client variants
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem, added
'                       FindLender and GetLenderCodes.
' STB       07/05/02    MSMS0069 Added the payment processing fields.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
' MC        04/06/2004  BMIDS763 New Column AllowProductSwitchFee Get & SET member functions added
' TW        09/10/2006  EP2_7 Added  code for Additional Borrowing Fee and Credit Limit Increase Fee

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_bVersionExists As Boolean

' Constants
Private Const m_sTableName = "MORTGAGELENDER"
Private Const m_sVersionField As String = "MORTGAGELENDERVERSIONNUMBER"

' BMids
'
'   AW      13/05/02    Added Get/Set for 'AllowReValuationFeeAdded', 'MinPercentArrangementFee'
'   MO      10/06/02    BMIDS00040 Change to MinPercentArrangementFee
'
'

Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "OrganisationID"
    
    ' If versioning exists, we'll need to add that too...
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub


Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function


Friend Function GetOrganisationID() As Variant
    Dim vGUID As Variant
    vGUID = m_clsTableAccess.GetVal("OrganisationID")
    ' DJP SQL Server port
    GetOrganisationID = vGUID
End Function


Friend Function GetLenderCode() As Variant
    GetLenderCode = m_clsTableAccess.GetVal("LenderCode")
End Function


Friend Function GetContactDetailsGUID() As Variant
    GetContactDetailsGUID = m_clsTableAccess.GetVal("ContactDetailsGUID")
End Function


Friend Function GetAddressGUID() As Variant
    GetAddressGUID = m_clsTableAccess.GetVal("AddressGUID")
End Function


Friend Function GetMIGThreshold() As Variant
    GetMIGThreshold = CStr(m_clsTableAccess.GetVal("MIGTHRESHOLD"))
End Function


Friend Function GetLenderName() As Variant
    GetLenderName = CStr(m_clsTableAccess.GetVal("LenderName"))
End Function


Friend Function GetStartDate() As Variant
    
    GetStartDate = m_clsTableAccess.GetDate("LenderStartDate")

End Function


Friend Function GetEndDate() As Variant
    GetEndDate = m_clsTableAccess.GetDate("LenderEndDate")
    
End Function


Friend Function GetLegalFeeSetKey() As Variant
    ' If this is an Omiga 4 database, we need to return the OrganisationID because that's
    ' how the legal fee set is identified. If it's a Xelector database we need to return
    ' the FeeSet, because that's how legal fees are identified in Xelector.
    
    ' So, first try and read the FeeSet from the Lender table (Xelector)...
    On Error GoTo XelectorFailed
    
    GetLegalFeeSetKey = m_clsTableAccess.GetVal("FEESET")
    
    Exit Function
XelectorFailed:
    On Error GoTo Failed
    g_clsErrorHandling.SaveError 0, ""
    ' So, lets try the OrganisationID
    GetLegalFeeSetKey = GetOrganisationID()
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Friend Function GetRateSet() As Variant
    GetRateSet = m_clsTableAccess.GetVal("RateSet")
End Function


Friend Function GetMaxNoLoans() As Variant
    GetMaxNoLoans = m_clsTableAccess.GetVal("MaxNoLoans")
End Function


Friend Function GetMIGStartLTV() As Variant
    GetMIGStartLTV = m_clsTableAccess.GetVal("MIGStartLTV")
End Function


Friend Function GetMinMIGPremium() As Variant
    GetMinMIGPremium = m_clsTableAccess.GetVal("MinimumMIGPremium")
End Function


Friend Function GetIgnoreMIGPremium() As Variant
    GetIgnoreMIGPremium = m_clsTableAccess.GetVal("IgnoreMIGPremiumAmount")
End Function


Friend Function GetAddIPTToMIGPremium() As Variant
    GetAddIPTToMIGPremium = m_clsTableAccess.GetVal("AddIPTToMIGPremium")
End Function


Friend Function GetTTFee() As Variant
    GetTTFee = m_clsTableAccess.GetVal("TTFee")
End Function


Friend Function GetSealingFee() As Variant
    GetSealingFee = m_clsTableAccess.GetVal("SealingFee")
End Function


Friend Function GetDeedsReleaseFee() As Variant
    GetDeedsReleaseFee = m_clsTableAccess.GetVal("DeedsReleaseFee")
End Function


Friend Function GetArrangementFeeAmount() As Variant
    GetArrangementFeeAmount = m_clsTableAccess.GetVal("MLArrangementFeeAmount")
End Function


Friend Function GetArrangementFeePercent() As Variant
    GetArrangementFeePercent = m_clsTableAccess.GetVal("MLArrangementFeePercent")
End Function


Friend Function GetArrangementFeePercentMax() As Variant
    GetArrangementFeePercentMax = m_clsTableAccess.GetVal("MLArrangementFeePercentMax")
End Function

Friend Function GetArrangementFeePercentMin() As Variant
    GetArrangementFeePercentMin = m_clsTableAccess.GetVal("MLARRANGEMENTFEEPERCENTMIN")
End Function

Friend Function GetAddCostsToLoan() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AddCostsToLoan")
    GetAddCostsToLoan = GetAddToLoanVal(vVal)
End Function


Friend Function GetAllowPortingFeeAdded() As Variant
    GetAllowPortingFeeAdded = m_clsTableAccess.GetVal("AllowPortingFeeAdded")
End Function


Friend Function GetAllowAdminFeeAdded() As Variant
    GetAllowAdminFeeAdded = m_clsTableAccess.GetVal("AllowAdminFeeAdded")
End Function


Friend Function GetAllowValuationFeeAdded() As Variant
    GetAllowValuationFeeAdded = m_clsTableAccess.GetVal("AllowValnFeeAdded")
End Function

Friend Function GetAllowReValuationFeeAdded() As Variant
    GetAllowReValuationFeeAdded = m_clsTableAccess.GetVal("AllowReValuationFeeAdded")
End Function

Friend Function GetAllowReinspectionFeeAdded() As Variant
    GetAllowReinspectionFeeAdded = m_clsTableAccess.GetVal("AllowReinsptFeeAdded")
End Function


Friend Function GetAllowArrangementFeeAdded() As Variant
    GetAllowArrangementFeeAdded = m_clsTableAccess.GetVal("ALLOWARRGEMTFEEADDED")
End Function


Friend Function GetAllowMIGFeeAdded() As Variant
    GetAllowMIGFeeAdded = m_clsTableAccess.GetVal("AllowMIGFeeAdded")
End Function


Friend Function GetAllowTTFeeAdded() As Variant
    GetAllowTTFeeAdded = m_clsTableAccess.GetVal("AllowTTFeeAdded")
End Function


Friend Function GetAllowSealingFeeAdded() As Variant
    GetAllowSealingFeeAdded = m_clsTableAccess.GetVal("AllowSealingFeeAdded")
End Function
'*=[MC]BMIDS763 - FEE SETS
Friend Function GetProductSwitchFeeAdded() As Variant
    GetProductSwitchFeeAdded = m_clsTableAccess.GetVal("AllowProductSwitchFeeAdded")
End Function
'*=SECTION END

' TW 09/10/2006 EP2_7
Friend Function GetCreditLimitIncreaseFeeAdded() As Variant
    GetCreditLimitIncreaseFeeAdded = m_clsTableAccess.GetVal("AllowCreditLimitIncFeeAdded")
End Function
' TW 09/10/2006 EP2_7 End
' TW 09/10/2006 EP2_7
Friend Function GetAllowFurtherLendingFeeAdded() As Variant
    GetAllowFurtherLendingFeeAdded = m_clsTableAccess.GetVal("AllowFurtherLendingFeeAdded")
End Function
' TW 09/10/2006 EP2_7 End

Friend Function GetAllowDeedsReleaseFeeAdded() As Variant
    GetAllowDeedsReleaseFeeAdded = m_clsTableAccess.GetVal("AllowDeedsRelFeeAdd")
End Function


Friend Function GetNonGuaranteedIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetNonGuaranteedIncomeMultiple = m_clsTableAccess.GetData("GUARANTEEDINCOMEMULT", ctrl, ctrlLabel)
End Function


Friend Function GetGuaranteedIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetGuaranteedIncomeMultiple = m_clsTableAccess.GetData("NONGUARANTEEDINCOMEMULT", ctrl, ctrlLabel)
End Function


Friend Function GetSingleIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetSingleIncomeMultiple = m_clsTableAccess.GetData("SINGLEINCOMEMULTIPLE", ctrl, ctrlLabel)
End Function


Friend Function GetHighestEarnerIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetHighestEarnerIncomeMultiple = m_clsTableAccess.GetData("HIGHESTEARNINCMULTIPLE", ctrl, ctrlLabel)
End Function


Friend Function GetSecondEarnerIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetSecondEarnerIncomeMultiple = m_clsTableAccess.GetData("SECONDEARNINCMULTIPLE", ctrl, ctrlLabel)
End Function


Friend Function GetJointIncomeMultiple(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As String
    GetJointIncomeMultiple = m_clsTableAccess.GetData("JOINTINCOMEMULTIPLE", ctrl, ctrlLabel)
End Function


Friend Function SetOrganisationID() As Variant
    Dim vVal As Variant
    Dim bRet As Boolean
    
    Dim clsGUID As New GuidAssist
    
    vVal = clsGUID.CreateGUID()
    
    bRet = m_clsTableAccess.SetVal("OrganisationID", vVal)
    ' DJP SQL Server port - read it back so we can get it in platform GUID format
    vVal = m_clsTableAccess.GetVal("ORGANISATIONID")
    If bRet = True Then
        SetOrganisationID = vVal
    End If

    g_clsVersion.HandleVersioning Me
    
End Function


Friend Function SetMIGThreshold(vVal As Variant) As Boolean
    SetMIGThreshold = m_clsTableAccess.SetVal("MIGTHRESHOLD", CVar(vVal))
End Function


Friend Function SetAddressGUID(vVal As Variant) As Boolean
    SetAddressGUID = m_clsTableAccess.SetVal("AddressGUID", CVar(vVal))
End Function


Friend Function SetContactDetailsGUID(vVal As Variant) As Boolean
    SetContactDetailsGUID = m_clsTableAccess.SetVal("ContactDetailsGUID", CVar(vVal))
End Function


Friend Function SetRateSet(sVal As Variant) As Boolean
    SetRateSet = m_clsTableAccess.SetVal("RateSet", CVar(sVal))
End Function


Friend Function SetLenderName(vVal As Variant) As Boolean
    SetLenderName = m_clsTableAccess.SetVal("LenderName", CVar(vVal))
End Function


Friend Function SetLenderCode(vVal As Variant) As Boolean
    SetLenderCode = m_clsTableAccess.SetVal("LenderCode", CVar(vVal))
End Function


Friend Function SetStartDate(vVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetDate("LenderStartDate", vVal)
End Function


Friend Function SetFeeSet(vVal As Variant) As Boolean
    Const sField As String = "FEESET"
    SetFeeSet = True
    
    If g_clsDataAccess.DoesFieldExist(m_sTableName, sField) Then
        SetFeeSet = m_clsTableAccess.SetVal(sField, CVar(vVal))
    End If

End Function


Friend Function SetEndDate(vVal As Variant) As Boolean
    SetEndDate = m_clsTableAccess.SetDate("LenderEndDate", CVar(vVal))
End Function


Friend Function SetMaxNoLoans(vVal As Variant) As Boolean
    SetMaxNoLoans = m_clsTableAccess.SetVal("MaxNoLoans", CVar(vVal))
End Function


Friend Function SetMIGStartLTV(vVal As Variant) As Boolean
    SetMIGStartLTV = m_clsTableAccess.SetVal("MIGStartLTV", CVar(vVal))
End Function


Friend Function SetMinMIGPremium(vVal As Variant) As Boolean
    SetMinMIGPremium = m_clsTableAccess.SetVal("MinimumMIGPremium", CVar(vVal))
End Function


Friend Function SetIgnoreMIGPremium(vVal As Variant) As Boolean
    SetIgnoreMIGPremium = m_clsTableAccess.SetVal("IgnoreMIGPremiumAmount", CVar(vVal))
End Function


Friend Function SetAddIPTToMIGPremium(vVal As Variant) As Boolean
    SetAddIPTToMIGPremium = m_clsTableAccess.SetVal("AddIPTToMIGPremium", CVar(vVal))
End Function


Friend Function SetTTFee(vVal As Variant) As Boolean
    SetTTFee = m_clsTableAccess.SetVal("TTFee", CVar(vVal))
End Function


Friend Function SetSealingFee(vVal As Variant) As Boolean
    SetSealingFee = m_clsTableAccess.SetVal("SealingFee", CVar(vVal))
End Function


Friend Function SetDeedsReleaseFee(vVal As Variant) As Boolean
    SetDeedsReleaseFee = m_clsTableAccess.SetVal("DeedsReleaseFee", CVar(vVal))
End Function


Friend Function SetArrangementFeeAmount(vVal As Variant) As Boolean
    SetArrangementFeeAmount = m_clsTableAccess.SetVal("MLArrangementFeeAmount", CVar(vVal))
End Function


Friend Function SetArrangementFeePercent(vVal As Variant) As Boolean
    SetArrangementFeePercent = m_clsTableAccess.SetVal("MLArrangementFeePercent", CVar(vVal))
End Function


Friend Function SetArrangementFeePercentMax(vVal As Variant) As Boolean
    SetArrangementFeePercentMax = m_clsTableAccess.SetVal("MLArrangementFeePercentMax", CVar(vVal))
End Function

Friend Function SetArrangementFeePercentMin(vVal As Variant) As Boolean
    SetArrangementFeePercentMin = m_clsTableAccess.SetVal("MLARRANGEMENTFEEPERCENTMIN", CVar(vVal))
End Function

Friend Function SetAddCostsToLoan(vVal As Variant) As Boolean
    vVal = SetAddToLoanVal(vVal)
    SetAddCostsToLoan = m_clsTableAccess.SetVal("AddCostsToLoan", vVal)
End Function


Friend Function SetAllowPortingFeeAdded(vVal As Variant) As Boolean
    SetAllowPortingFeeAdded = m_clsTableAccess.SetVal("AllowPortingFeeAdded", CVar(vVal))
End Function

Friend Function SetAllowAdminFeeAdded(vVal As Variant) As Boolean
    SetAllowAdminFeeAdded = m_clsTableAccess.SetVal("AllowAdminFeeAdded", CVar(vVal))
End Function
' TW 09/10/2006 EP2_7
Friend Function SetAllowFurtherLendingFeeAdded(vVal As Variant) As Boolean
    SetAllowFurtherLendingFeeAdded = m_clsTableAccess.SetVal("AllowFurtherLendingFeeAdded", CVar(vVal))
End Function
' TW 09/10/2006 EP2_7 End
' TW 09/10/2006 EP2_7
Friend Function SetCreditLimitIncreaseFeeAdded(vVal As Variant) As Boolean
    SetCreditLimitIncreaseFeeAdded = m_clsTableAccess.SetVal("AllowCreditLimitIncFeeAdded", CVar(vVal))
End Function
' TW 09/10/2006 EP2_7 End

Friend Function SetAllowValuationFeeAdded(vVal As Variant) As Boolean
    SetAllowValuationFeeAdded = m_clsTableAccess.SetVal("AllowValnFeeAdded", CVar(vVal))
End Function

Friend Function SetAllowReValuationFeeAdded(vVal As Variant) As Boolean
    SetAllowReValuationFeeAdded = m_clsTableAccess.SetVal("AllowReValuationFeeAdded", CVar(vVal))
End Function

Friend Function SetAllowReinspectionFeeAdded(vVal As Variant) As Boolean
    SetAllowReinspectionFeeAdded = m_clsTableAccess.SetVal("AllowReinsptFeeAdded", CVar(vVal))
End Function

Friend Function SetProductSwitchFeeAdded(vVal As Variant) As Boolean
    SetProductSwitchFeeAdded = m_clsTableAccess.SetVal("AllowProductSwitchFeeAdded", CVar(vVal))
End Function

Friend Function SetAllowArrangementFeeAdded(vVal As Variant) As Boolean
    SetAllowArrangementFeeAdded = m_clsTableAccess.SetVal("ALLOWARRGEMTFEEADDED", CVar(vVal))
End Function


Friend Function SetAllowMIGFeeAdded(vVal As Variant) As Boolean
    SetAllowMIGFeeAdded = m_clsTableAccess.SetVal("AllowMIGFeeAdded", CVar(vVal))
End Function


Friend Function SetAllowTTFeeAdded(vVal As Variant) As Boolean
    SetAllowTTFeeAdded = m_clsTableAccess.SetVal("AllowTTFeeAdded", CVar(vVal))
End Function


Friend Function SetAllowSealingFeeAdded(vVal As Variant) As Boolean
    SetAllowSealingFeeAdded = m_clsTableAccess.SetVal("AllowSealingFeeAdded", CVar(vVal))
End Function


Friend Function SetAllowDeedsReleaseFeeAdded(vVal As Variant) As Boolean
    SetAllowDeedsReleaseFeeAdded = m_clsTableAccess.SetVal("AllowDeedsRelFeeAdd", CVar(vVal))
End Function


Friend Function SetNonGuaranteedIncomeMultiple(vVal As Variant) As Boolean
    SetNonGuaranteedIncomeMultiple = m_clsTableAccess.SetVal("GUARANTEEDINCOMEMULT", vVal)
End Function


Friend Function SetGuaranteedIncomeMultiple(vVal As Variant) As Boolean
    SetGuaranteedIncomeMultiple = m_clsTableAccess.SetVal("NONGUARANTEEDINCOMEMULT", vVal)
End Function


Friend Function SetSingleIncomeMultiple(vVal As Variant) As Boolean
    SetSingleIncomeMultiple = m_clsTableAccess.SetVal("SINGLEINCOMEMULTIPLE", vVal)
End Function


Friend Function SetHighestEarnerIncomeMultiple(vVal As Variant) As Boolean
    SetHighestEarnerIncomeMultiple = m_clsTableAccess.SetVal("HIGHESTEARNINCMULTIPLE", vVal)
End Function


Friend Function SetSecondEarnerIncomeMultiple(vVal As Variant) As Boolean
    SetSecondEarnerIncomeMultiple = m_clsTableAccess.SetVal("SECONDEARNINCMULTIPLE", vVal)
End Function


Friend Function SetJointIncomeMultiple(vVal As Variant) As Boolean
    SetJointIncomeMultiple = m_clsTableAccess.SetVal("JOINTINCOMEMULTIPLE", vVal)
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


Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function


Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub


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
    Set m_clsTableAccess = clsTable
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
    'colListLine.Add GetStartDate()
    Dim sLenderCode As String
    sLenderCode = GetLenderCode()
    
    colListLine.Add sLenderCode
    'colListLine.Add GetOrganisationID()
    colListLine.Add GetLenderName()
    colListLine.Add GetStartDate()
    colListLine.Add GetEndDate()
    colListLine.Add sLenderCode, OBJECT_DESCRIPTION_KEY
End Sub


Private Function GetAddToLoanVal(ByVal vVal As Variant) As Variant
    Dim bAdd As Boolean
    
    bAdd = True
    
    ' Default to "First". Only set to last if this radio ("First") is off.
    If Not IsNull(vVal) Then
        If CStr(vVal) = "L" Then
            bAdd = False
        End If
    End If
    
    GetAddToLoanVal = CVar(bAdd)
    
End Function


Private Function SetAddToLoanVal(ByVal vVal As Variant) As Variant
    Dim bAdd As Boolean
    Dim sAdd As String
    
    sAdd = "F"
    
    ' Default to "First". Only set to last if this radio ("First") is off.
    If Not IsNull(vVal) Then
        If CBool(vVal) = False Then
            sAdd = "L"
        End If
    End If
    
    SetAddToLoanVal = CVar(sAdd)
    
End Function


Friend Function DoesLenderExist(sLenderCode As String) As Boolean
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTable As String
    Dim sSearch As String
    
    bRet = False
    
    sSearch = "select LenderCode from " + m_sTableName + " where LenderCode = '" + FormatString(sLenderCode) + "'"

    If m_bVersionExists Then
        sSearch = sSearch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
        End If
    End If
    
    DoesLenderExist = bRet
End Function


Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub DeleteLender()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub GetLenders(Optional sOrderBy As String = "")
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim sExtra As String
    'Need to fix this to get the version number too, or the product with the version we want
    
    sSearch = "select * FROM " + m_sTableName
        
    If m_bVersionExists Then
        sSearch = sSearch + " WHERE " + m_sVersionField + " = " & g_sVersionNumber
    End If
    
    If Len(sOrderBy) > 0 Then
        sSearch = sSearch & " Order By " & sOrderBy
    End If
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, "Products"
    m_clsTableAccess.SetRecordSet rs

    Exit Sub
Failed:
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


Public Sub FindLender(sLenderCode As String)
    On Error GoTo Failed
    
    If Len(sLenderCode) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Lender Code is empty"
    End If
    
    If m_clsTableAccess.RecordCount > 0 Then
        m_clsTableAccess.ApplyFilter "LENDERCODE = " & g_clsSQLAssistSP.FormatString(sLenderCode)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "MortgageLendersTable.FindLender" & Err.DESCRIPTION
End Sub


Friend Function GetLenderCodes() As ADODB.Recordset
    On Error GoTo Failed
    Dim sTable As String
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    sTable = "MORTGAGELENDER"
    
    sSearch = "SELECT LENDERCODE, ORGANISATIONID FROM " + sTable + " ORDER BY LENDERCODE"

    Set rs = g_clsDataAccess.GetTableData(sTable, sSearch)
    m_clsTableAccess.SetRecordSet rs
    
    Set GetLenderCodes = rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Public Function GetLenderCodeField() As String
    GetLenderCodeField = "LenderCode"
End Function


Friend Function GetBarclaysCHAPSSortCode() As Variant
    GetBarclaysCHAPSSortCode = m_clsTableAccess.GetVal("BarclaysCHAPSSortCode")
End Function


Friend Function GetBarclaysCHAPSAccountNumber() As Variant
    GetBarclaysCHAPSAccountNumber = m_clsTableAccess.GetVal("BarclaysCHAPSAccountNumber")
End Function


Friend Function GetBarclaysCHAPSOrderingInst() As Variant
    GetBarclaysCHAPSOrderingInst = m_clsTableAccess.GetVal("BarclaysCHAPSOrderingInst")
End Function


Friend Function GetHexagonCHAPSAccountNumber() As Variant
    GetHexagonCHAPSAccountNumber = m_clsTableAccess.GetVal("HexagonCHAPSAccountNumber")
End Function


Friend Function GetBACSSortCode() As Variant
    GetBACSSortCode = m_clsTableAccess.GetVal("BACSSortCode")
End Function


Friend Function GetBACSAccountNumber() As Variant
    GetBACSAccountNumber = m_clsTableAccess.GetVal("BACSAccountNumber")
End Function


Friend Function SetBarclaysCHAPSSortCode(ByVal vVal As Variant) As Boolean
    SetBarclaysCHAPSSortCode = m_clsTableAccess.SetVal("BarclaysCHAPSSortCode", vVal)
End Function


Friend Function SetBarclaysCHAPSAccountNumber(ByVal vVal As Variant) As Boolean
    SetBarclaysCHAPSAccountNumber = m_clsTableAccess.SetVal("BarclaysCHAPSAccountNumber", vVal)
End Function


Friend Function SetBarclaysCHAPSOrderingInst(ByVal vVal As Variant) As Boolean
    SetBarclaysCHAPSOrderingInst = m_clsTableAccess.SetVal("BarclaysCHAPSOrderingInst", vVal)
End Function


Friend Function SetHexagonCHAPSAccountNumber(ByVal vVal As Variant) As Boolean
    SetHexagonCHAPSAccountNumber = m_clsTableAccess.SetVal("HexagonCHAPSAccountNumber", vVal)
End Function


Friend Function SetBACSSortCode(ByVal vVal As Variant) As Boolean
    SetBACSSortCode = m_clsTableAccess.SetVal("BACSSortCode", vVal)
End Function


Friend Function SetBACSAccountNumber(ByVal vVal As Variant) As Boolean
    SetBACSAccountNumber = m_clsTableAccess.SetVal("BACSAccountNumber", vVal)
End Function
