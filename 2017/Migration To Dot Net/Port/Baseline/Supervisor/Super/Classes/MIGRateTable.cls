VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIGRateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MIGRateTable
' Description   : Contains all data access routines for the MIGRate table
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Update
'                       GetMIGRateSets to accept a rate number instead of a collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
Implements BandedTable

' Private data
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private m_sSearch As String
Private m_sFirstMatch As String

' Constants
Private Const m_sTableName = "MIGRATE"
Private Const m_sVersionField As String = "MIGRATESETVERSIONNUMBER"
Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
        
    colFields.Add "RateSet"
    colFields.Add "MIGRateStartDate"
    
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "RateSet"
    colUpdateFields.Add "MIGRateStartDate"
    
    ' If versioning exists, we'll need to add that too...
    If g_clsVersion.DoesVersioningExist() Then
        colFields.Add m_sVersionField
    End If
    
    m_clsBandedTable.SetUpdateFields colUpdateFields
    
    m_clsTableAccess.Create m_sTableName
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetRateSet() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("RateSet")
    GetRateSet = CStr(vVal)
End Function
Friend Function GetStartDate() As String
    Dim vVal As Variant

    vVal = m_clsTableAccess.GetVal("MIGRateStartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetMaximumLTV() As Variant
    GetMaximumLTV = CStr(m_clsTableAccess.GetVal("MaximumLTV"))
End Function
Friend Function GetMaxLoanAmount() As Variant
    GetMaxLoanAmount = CStr(m_clsTableAccess.GetVal("MaximumLoanAmount"))
End Function
Friend Function GetMIGRate() As Variant
    GetMIGRate = CStr(m_clsTableAccess.GetVal("IndemnityRate"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetRateSet(vVal As Variant) As Boolean
    SetRateSet = m_clsTableAccess.SetVal("RateSet", vVal)

    ' Versioning
    g_clsVersion.HandleVersioning Me
End Function
Friend Function SetStartDate(vVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("MIGRateStartDate", vVal)
End Function
Friend Function SetMaximumLTV(vVal As Variant) As Boolean
    SetMaximumLTV = m_clsTableAccess.SetVal("MaximumLTV", vVal)
End Function
Friend Function SetMaxLoanAmount(vVal As Variant) As Boolean
    SetMaxLoanAmount = m_clsTableAccess.SetVal("MaximumLoanAmount", vVal)
End Function
Friend Function SetMIGRate(vVal As Variant) As Boolean
    SetMIGRate = m_clsTableAccess.SetVal("RateSet", vVal)
End Function
Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub
Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub
Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub
Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetRateSet()
    colHeader.Add GetStartDate()
End Sub
Public Sub SetPopulateNull()
    m_sSearch = "Select * from " + m_sTableName + " where Name is null"
End Sub
Public Sub SetPopulateType(enumType As PopulateType)
    m_clsTableAccess.SetPopulateType enumType
End Sub
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
Private Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
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
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
End Sub
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
Private Function TableAccess_GetUpdateValues() As Collection
    Set TableAccess_GetUpdateValues = m_clsTableAccess.GetUpdateValues
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
    m_clsTableAccess.SetPopulateType enumType
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
Public Function TableAccess_GetUpdateFields() As Collection
    Dim colKeys As New Collection
    
    colKeys.Add "RateSet"
    colKeys.Add "MIGRateStartDate"
    
    Set TableAccess_GetUpdateFields = colKeys
End Function
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetMaximumLTV()
    colListLine.Add GetMaxLoanAmount()
    colListLine.Add Me.GetMIGRate
End Sub
Friend Function GetMIGRateSets(vRateSet As Variant) As ADODB.Recordset
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
    If Not IsEmpty(vRateSet) Then
        sSearch = "SELECT " & _
                        "* " & _
                  "FROM " & _
                        m_sTableName & " A, " & _
                        "MIGRATESET " & _
                   "WHERE " & _
                        "A.MAXIMUMLTV = " & _
                        "(SELECT " & _
                            "MIN(MAXIMUMLTV) " & _
                        "FROM " & _
                            m_sTableName & " B " & _
                        "WHERE " & _
                            "A.RateSet = B.RateSet AND A.MIGRATESTARTDATE = B.MIGRATESTARTDATE ) AND " & _
                            "A.MAXIMUMLOANAMOUNT = "
        sSearch = sSearch & " (SELECT " & _
                                    "MIN(MAXIMUMLOANAMOUNT) " & _
                                "FROM " & _
                                    m_sTableName & " C " & _
                                "WHERE " & _
                                    "A.RATESET = C.RATESET AND " & _
                                    "A.MIGRATESTARTDATE = C.MIGRATESTARTDATE AND A.MAXIMUMLTV = C.MAXIMUMLTV) AND " & _
                                    "A.RATESET = MIGRATESET.RATESET AND A.RATESET = " & vRateSet
        
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "MIGRATE - no key values supplied"
    End If

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rs
    
    Set GetMIGRateSets = rs

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
Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub
Public Sub DeleteMIGRates(sRateSet As String, Optional sDatabase As String)
    On Error GoTo Failed
    Dim sSQL As String
    
    sSQL = "DELETE FROM " + m_sTableName + " WHERE RATESET = " + sRateSet
    
    g_clsDataAccess.ExecuteCommand sSQL, sDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetAllRatesFromRateSet(sRateSet As String)
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim colMatchKeys As Collection

    Set colMatchKeys = New Collection
    Set colMatchValues = New Collection

    colMatchKeys.Add "RATESET"
    colMatchValues.Add sRateSet
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    m_clsTableAccess.SetKeyMatchFields colMatchKeys
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
