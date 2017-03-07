VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductBandsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "MORTGAGEPRODUCTBANDS"
Private m_sSearch As String
Private Const m_sVersionField As String = "PRODUCTVERSIONNUMBER"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "MortgageProductCode"
    colFields.Add "StartDate"
    
    ' If versioning exists, we'll need to add that too...
    If g_clsVersion.DoesVersioningExist() Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetMortgageProductCode() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MortgageProductCode")
    GetMortgageProductCode = CStr(vVal)
End Function
Friend Function GetStartDate() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("StartDate")

    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If

End Function
Friend Function GetInterestRateAdjustment1() As Variant
    GetInterestRateAdjustment1 = m_clsTableAccess.GetVal("InterestRateAdjustment1")
End Function
Friend Function GetInterestRateAdjustment2() As Variant
    GetInterestRateAdjustment2 = m_clsTableAccess.GetVal("InterestRateAdjustment2")
End Function
Friend Function GetInterestRateAdjustment3() As Variant
    GetInterestRateAdjustment3 = m_clsTableAccess.GetVal("InterestRateAdjustment3")
End Function
Friend Function GetInterestRateAdjustment4() As Variant
    GetInterestRateAdjustment4 = m_clsTableAccess.GetVal("InterestRateAdjustment4")
End Function
Friend Function GetInterestRateAdjustment5() As Variant
    GetInterestRateAdjustment5 = m_clsTableAccess.GetVal("InterestRateAdjustment5")
End Function
Friend Function GetUpperLimitBand1() As Variant
    GetUpperLimitBand1 = m_clsTableAccess.GetVal("UpperLimitBand1")
End Function
Friend Function GetUpperLimitBand2() As Variant
    GetUpperLimitBand2 = m_clsTableAccess.GetVal("UpperLimitBand2")
End Function
Friend Function GetUpperLimitBand3() As Variant
    GetUpperLimitBand3 = m_clsTableAccess.GetVal("UpperLimitBand3")
End Function
Friend Function GetUpperLimitBand4() As Variant
    GetUpperLimitBand4 = m_clsTableAccess.GetVal("UpperLimitBand4")
End Function
Friend Function GetUpperLimitBand5() As Variant
    GetUpperLimitBand5 = m_clsTableAccess.GetVal("UpperLimitBand5")
End Function
' Early Redemption
Friend Function GetEarlyRedemptionPeriod1() As Variant
    GetEarlyRedemptionPeriod1 = m_clsTableAccess.GetVal("EarlyRedemptionPeriod1")
End Function
Friend Function GetEarlyRedemptionAmount1() As Variant
    GetEarlyRedemptionAmount1 = m_clsTableAccess.GetVal("EarlyRedemptionAmount1")
End Function
Friend Function GetEarlyRedemptionPercent1() As Variant
    GetEarlyRedemptionPercent1 = m_clsTableAccess.GetVal("EarlyRedemptionPercent1")
End Function
Friend Function GetEarlyRedemptionPeriod2() As Variant
    GetEarlyRedemptionPeriod2 = m_clsTableAccess.GetVal("EarlyRedemptionPeriod2")
End Function
Friend Function GetEarlyRedemptionAmount2() As Variant
    GetEarlyRedemptionAmount2 = m_clsTableAccess.GetVal("EarlyRedemptionAmount2")
End Function
Friend Function GetEarlyRedemptionPercent2() As Variant
    GetEarlyRedemptionPercent2 = m_clsTableAccess.GetVal("EarlyRedemptionPercent2")
End Function
Friend Function GetEarlyRedemptionPeriod3() As Variant
    GetEarlyRedemptionPeriod3 = m_clsTableAccess.GetVal("EarlyRedemptionPeriod3")
End Function
Friend Function GetEarlyRedemptionAmount3() As Variant
    GetEarlyRedemptionAmount3 = m_clsTableAccess.GetVal("EarlyRedemptionAmount3")
End Function
Friend Function GetEarlyRedemptionPercent3() As Variant
    GetEarlyRedemptionPercent3 = m_clsTableAccess.GetVal("EarlyRedemptionPercent3")
End Function
Friend Function GetEarlyRedemptionPeriod4() As Variant
    GetEarlyRedemptionPeriod4 = m_clsTableAccess.GetVal("EarlyRedemptionPeriod4")
End Function
Friend Function GetEarlyRedemptionAmount4() As Variant
    GetEarlyRedemptionAmount4 = m_clsTableAccess.GetVal("EarlyRedemptionAmount4")
End Function
Friend Function GetEarlyRedemptionPercent4() As Variant
    GetEarlyRedemptionPercent4 = m_clsTableAccess.GetVal("EarlyRedemptionPercent4")
End Function
Friend Function GetEarlyRedemptionPeriod5() As Variant
    GetEarlyRedemptionPeriod5 = m_clsTableAccess.GetVal("EarlyRedemptionPeriod5")
End Function
Friend Function GetEarlyRedemptionAmount5() As Variant
    GetEarlyRedemptionAmount5 = m_clsTableAccess.GetVal("EarlyRedemptionAmount5")
End Function
Friend Function GetEarlyRedemptionPercent5() As Variant
    GetEarlyRedemptionPercent5 = m_clsTableAccess.GetVal("EarlyRedemptionPercent5")
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetMortgageProductCode(sVal As Variant) As Boolean
    SetMortgageProductCode = m_clsTableAccess.SetVal("MortgageProductCode", CVar(sVal))

    ' Versioning, if it exists
    g_clsVersion.HandleVersioning Me

End Function
Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetDate("StartDate", CVar(sVal))
End Function
Friend Function SetInterestRateAdjustment1(sVal As Variant) As Boolean
    SetInterestRateAdjustment1 = m_clsTableAccess.SetVal("InterestRateAdjustment1", CVar(sVal))
End Function
Friend Function SetInterestRateAdjustment2(sVal As Variant) As Boolean
    SetInterestRateAdjustment2 = m_clsTableAccess.SetVal("InterestRateAdjustment2", CVar(sVal))
End Function
Friend Function SetInterestRateAdjustment3(sVal As Variant) As Boolean
    SetInterestRateAdjustment3 = m_clsTableAccess.SetVal("InterestRateAdjustment3", CVar(sVal))
End Function
Friend Function SetInterestRateAdjustment4(sVal As Variant) As Boolean
    SetInterestRateAdjustment4 = m_clsTableAccess.SetVal("InterestRateAdjustment4", CVar(sVal))
End Function
Friend Function SetInterestRateAdjustment5(sVal As Variant) As Boolean
    SetInterestRateAdjustment5 = m_clsTableAccess.SetVal("InterestRateAdjustment5", CVar(sVal))
End Function
Friend Function setUpperLimitBand1(sVal As Variant) As Variant
    setUpperLimitBand1 = m_clsTableAccess.SetVal("UpperLimitBand1", sVal)
End Function
Friend Function setUpperLimitBand2(sVal As Variant) As Variant
    setUpperLimitBand2 = m_clsTableAccess.SetVal("UpperLimitBand2", sVal)
End Function
Friend Function setUpperLimitBand3(sVal As Variant) As Variant
    setUpperLimitBand3 = m_clsTableAccess.SetVal("UpperLimitBand3", sVal)
End Function
Friend Function setUpperLimitBand4(sVal As Variant) As Variant
    setUpperLimitBand4 = m_clsTableAccess.SetVal("UpperLimitBand4", sVal)
End Function
Friend Function SetUpperLimitBand5(sVal As Variant) As Variant
    SetUpperLimitBand5 = m_clsTableAccess.SetVal("UpperLimitBand5", sVal)
End Function
' Early Redemption
Friend Function SetEarlyRedemptionPeriod1(sVal As Variant) As Variant
    SetEarlyRedemptionPeriod1 = m_clsTableAccess.SetVal("EarlyRedemptionPeriod1", sVal)
End Function
Friend Function SetEarlyRedemptionAmount1(sVal As Variant) As Variant
    SetEarlyRedemptionAmount1 = m_clsTableAccess.SetVal("EarlyRedemptionAmount1", sVal)
End Function
Friend Function SetEarlyRedemptionPercent1(sVal As Variant) As Variant
    SetEarlyRedemptionPercent1 = m_clsTableAccess.SetVal("EarlyRedemptionPercent1", sVal)
End Function
Friend Function SetEarlyRedemptionPeriod2(sVal As Variant) As Variant
    SetEarlyRedemptionPeriod2 = m_clsTableAccess.SetVal("EarlyRedemptionPeriod2", sVal)
End Function
Friend Function SetEarlyRedemptionAmount2(sVal As Variant) As Variant
    SetEarlyRedemptionAmount2 = m_clsTableAccess.SetVal("EarlyRedemptionAmount2", sVal)
End Function
Friend Function SetEarlyRedemptionPercent2(sVal As Variant) As Variant
    SetEarlyRedemptionPercent2 = m_clsTableAccess.SetVal("EarlyRedemptionPercent2", sVal)
End Function
Friend Function SetEarlyRedemptionPeriod3(sVal As Variant) As Variant
    SetEarlyRedemptionPeriod3 = m_clsTableAccess.SetVal("EarlyRedemptionPeriod3", sVal)
End Function
Friend Function SetEarlyRedemptionAmount3(sVal As Variant) As Variant
    SetEarlyRedemptionAmount3 = m_clsTableAccess.SetVal("EarlyRedemptionAmount3", sVal)
End Function
Friend Function SetEarlyRedemptionPercent3(sVal As Variant) As Variant
    SetEarlyRedemptionPercent3 = m_clsTableAccess.SetVal("EarlyRedemptionPercent3", sVal)
End Function
Friend Function SetEarlyRedemptionPeriod4(sVal As Variant) As Variant
    SetEarlyRedemptionPeriod4 = m_clsTableAccess.SetVal("EarlyRedemptionPeriod4", sVal)
End Function
Friend Function SetEarlyRedemptionAmount4(sVal As Variant) As Variant
    SetEarlyRedemptionAmount4 = m_clsTableAccess.SetVal("EarlyRedemptionAmount4", sVal)
End Function
Friend Function SetEarlyRedemptionPercent4(sVal As Variant) As Variant
    SetEarlyRedemptionPercent4 = m_clsTableAccess.SetVal("EarlyRedemptionPercent4", sVal)
End Function
Friend Function SetEarlyRedemptionPeriod5(sVal As Variant) As Variant
    SetEarlyRedemptionPeriod5 = m_clsTableAccess.SetVal("EarlyRedemptionPeriod5", sVal)
End Function
Friend Function SetEarlyRedemptionAmount5(sVal As Variant) As Variant
    SetEarlyRedemptionAmount5 = m_clsTableAccess.SetVal("EarlyRedemptionAmount5", sVal)
End Function
Friend Function SetEarlyRedemptionPercent5(sVal As Variant) As Variant
    SetEarlyRedemptionPercent5 = m_clsTableAccess.SetVal("EarlyRedemptionPercent5", sVal)
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

'Private Function TableAccess_DoUpdateSets() As Boolean
'    TableAccess_DoUpdateSets = m_clsTableAccess.DoUpdateSets()
'End Function
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS) As ADODB.Recordset
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
' DJP This can be added later if required.
Public Function TableAccess_GetUpdateValues() As Collection
End Function
' DJP This can be added later if required.
'Public Function TableAccess_GetUpdateFields() As Collection
'End Function
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetStartDate()
    colListLine.Add GetMortgageProductCode()
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
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
