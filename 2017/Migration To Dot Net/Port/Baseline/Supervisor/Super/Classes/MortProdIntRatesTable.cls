VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIntRatesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdPurpOfLoanTable
' Description   : Contains all data access routines for the InterestRateType table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       03/12/01    SYS2912 SQL Server locking problem, make this a Banded table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Implements BandedTable
' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable

' Constants
Private Const m_sTableName = "INTERESTRATETYPE"
Private Const m_sVersionField As String = "PRODUCTVERSIONNUMBER"

'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009. Amended :TableAccess_GetRowOfData
'                                   Added   :GetBaseRateSet,SetBaseRateSet

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    Dim colUpdateValues As Collection
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "InterestRateTypeSequenceNumber"
    
    
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsTableAccess.SetDuplicateKeys colFields
    
    ' And the Banded table
    Set colUpdateValues = New Collection
    colUpdateValues.Add "MORTGAGEPRODUCTCODE"
    colUpdateValues.Add "STARTDATE"
    
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    m_clsBandedTable.SetUpdateFields colUpdateValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Function GetMortgageProductCode() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MortgageProductCode")
    GetMortgageProductCode = CStr(vVal)
End Function

Friend Function GetStartDate() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("StartDate")

    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If

End Function

Friend Function GetIntRateSeqNumber() As Variant
    GetIntRateSeqNumber = m_clsTableAccess.GetVal("InterestRateTypeSequenceNumber")
End Function

Friend Function GetInterestRateEndDate() As Variant
    GetInterestRateEndDate = m_clsTableAccess.GetVal("InterestRateEndDate")
End Function

Friend Function GetInterestRatePeriod() As Variant
    GetInterestRatePeriod = m_clsTableAccess.GetVal("InterestRatePeriod")
End Function

Friend Function GetRateType() As Variant
    GetRateType = m_clsTableAccess.GetVal("RateType")
End Function

Friend Function GetRate() As Variant
    GetRate = m_clsTableAccess.GetVal("Rate")
End Function

Friend Function GetCeilingRate() As Variant
    GetCeilingRate = m_clsTableAccess.GetVal("CeilingRate")
End Function

Friend Function GetFlooredRate() As Variant
    GetFlooredRate = m_clsTableAccess.GetVal("FlooredRate")
End Function

Friend Function GetRedemptionPeriod() As Variant
    GetRedemptionPeriod = m_clsTableAccess.GetVal("EarlyRedemptionPeriod")
End Function

Friend Function GetPaymentPeriod() As Variant
    GetPaymentPeriod = m_clsTableAccess.GetVal("PartRepaymentPeriod")
End Function

Friend Function SetMortgageProductCode(sVal As Variant) As Boolean
    SetMortgageProductCode = m_clsTableAccess.SetVal("MortgageProductCode", CVar(sVal))

    ' Versioning, if it exists
    g_clsVersion.HandleVersioning Me

End Function

Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("StartDate", CVar(sVal))
End Function

Friend Function SetIntRateSeqNumber(sVal As Variant) As Boolean
    SetIntRateSeqNumber = m_clsTableAccess.SetVal("InterestRateTypeSequenceNumber", CVar(sVal))
End Function

Friend Function SetInterestRateEndDate(sVal As Variant) As Boolean
    SetInterestRateEndDate = m_clsTableAccess.SetVal("InterestRateEndDate", CVar(sVal))
End Function

Friend Function SetInterestRatePeriod(sVal As Variant) As Boolean
    SetInterestRatePeriod = m_clsTableAccess.SetVal("InterestRatePeriod", CVar(sVal))
End Function

Friend Function SetRateType(sVal As Variant) As Boolean
    SetRateType = m_clsTableAccess.SetVal("RateType", CVar(sVal))
End Function

Friend Function SetRate(sVal As Variant) As Boolean
    SetRate = m_clsTableAccess.SetVal("Rate", CVar(sVal))
End Function

Friend Function SetCeilingRate(sVal As Variant) As Boolean
    SetCeilingRate = m_clsTableAccess.SetVal("CeilingRate", CVar(sVal))
End Function

Friend Function SetFlooredRate(sVal As Variant) As Boolean
    SetFlooredRate = m_clsTableAccess.SetVal("FlooredRate", CVar(sVal))
End Function

Friend Function SetRedemptionPeriod(sVal As Variant) As Boolean
    SetRedemptionPeriod = m_clsTableAccess.SetVal("EarlyRedemptionPeriod", CVar(sVal))
End Function

Friend Function SetRepaymentPeriod(sVal As Variant) As Boolean
    SetRepaymentPeriod = m_clsTableAccess.SetVal("PartRepaymentPeriod", CVar(sVal))
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
    On Error GoTo Failed
    
    If g_clsVersion.DoesVersioningExist() Then
        colMatchData.Add g_sVersionNumber
    End If

    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
'GD        16/05/2002   BMIDS00009
    
    colListLine.Add GetIntRateSeqNumber()
    colListLine.Add GetRateType()
    colListLine.Add GetInterestRatePeriod()
    colListLine.Add GetInterestRateEndDate()
    colListLine.Add GetRate()
    colListLine.Add GetCeilingRate()
    colListLine.Add GetFlooredRate()
    'GD BMIDS
    colListLine.Add GetBaseRateSet()
    
End Sub


Friend Sub GetInterestRates(colProductKeys As Collection)  'sCode As String, sDate As String)
On Error GoTo Failed
    
    Dim sSearch As String
    Dim sType As String
    Dim sIncentiveTable As String
    Dim clsMortgageProduct As New MortgageProductTable
    Dim clsTableAccess As TableAccess
    Dim colFields As Collection
    
    Set clsTableAccess = clsMortgageProduct
    Set colFields = clsTableAccess.GetKeyMatchFields()
    sSearch = "SELECT " & _
                    "* " & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE "
    ' DJP SQL Server port - moved to DataAccess
    sSearch = sSearch + g_clsDataAccess.GenerateSearchCriteria(colFields, colProductKeys)

    m_clsTableAccess.GetTableDataFromSearch sSearch
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteInterestRates(Optional sTargetDatabase As String)
    On Error GoTo Failed
    Dim sCommand As String
    Dim sStartDate As String
    Dim sProductCode As String
    Dim colMatchValues As Collection
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues
    
    sProductCode = colMatchValues(1)
    sStartDate = colMatchValues(2)
    
    ' DJP SQL Server port
    sCommand = "DELETE " & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE " & _
                    "MORTGAGEPRODUCTCODE " & " = " & g_clsSQLAssistSP.FormatString(sProductCode) & " AND " & _
                    g_clsSQLAssistSP.TruncateDateColumn("STARTDATE") & " = " & g_clsSQLAssistSP.FormatDate(sStartDate)
    
    g_clsDataAccess.ExecuteCommand sCommand, sTargetDatabase
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

Friend Function SetCalculationType(sName As String) As Boolean
    SetCalculationType = m_clsTableAccess.SetVal("INTERESTCALCULATIONTYPE", CVar(sName))
End Function

Friend Function SetRedemptionTerms(sName As String) As Boolean
    SetRedemptionTerms = m_clsTableAccess.SetVal("EARLYREDEMPTIONTERMS", CVar(sName))
End Function

Friend Function SetRepaymentTerms(sName As String) As Boolean
    SetRepaymentTerms = m_clsTableAccess.SetVal("PARTREPAYMENTTERMS", CVar(sName))
End Function

Public Function GetCalculationType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("INTERESTCALCULATIONTYPE")
    GetCalculationType = CStr(vVal)
End Function

Public Function GetRedemptionTerms() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EARLYREDEMPTIONTERMS")
    GetRedemptionTerms = CStr(vVal)
End Function

Public Function GetRepaymentTerms() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PARTREPAYMENTTERMS")
    GetRepaymentTerms = CStr(vVal)
End Function

Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
End Sub

Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub

Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub

Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub

Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub
Friend Function GetBaseRateSet() As Variant
'GD        16/05/2002   BMIDS00009
    GetBaseRateSet = m_clsTableAccess.GetVal("BaseRateSet")
End Function
Friend Function SetBaseRateSet(sVal As Variant) As Boolean
'GD        16/05/2002   BMIDS00009
    SetBaseRateSet = m_clsTableAccess.SetVal("BaseRateSet", CVar(sVal))
End Function

