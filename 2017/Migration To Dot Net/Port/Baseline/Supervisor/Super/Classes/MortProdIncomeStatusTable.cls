VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncomeStatusTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncomeStatusTable
' Description   : Performs all tableaccess functions on the MortgageProductIncomeStatus table
'
' Change history
' Prog      Date        Description
' TW        30/11/2006  EP2_253 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "MORTGAGEPRODUCTINCOMESTATUS"
Public Sub GetAvailableIncomeStatusSettings(strMortgageProductCode As String, strStartDate As String)
'Dim sSearch As String
'Dim rsConditions As ADODB.Recordset
'
'    On Error GoTo Failed
'    sSearch = "SELECT VALUENAME AS MORTGAGEINCOMESTATUS, M.* From MORTGAGEPRODUCTINCOMESTATUS M " & _
'              "INNER JOIN COMBOVALUE C ON M.INCOMESTATUS = C.VALUEID " & _
'              "WHERE C.GROUPNAME = 'ApplicationIncomeStatus' " & _
'              "AND MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' " & _
'              "AND STARTDATE = '" & strStartDate & "'" & _
'              "AND M.INCOMESTATUS NOT IN " & _
'              "(SELECT INCOMESTATUS From MORTGAGEPRODUCTINCOMESTATUS " & _
'              "WHERE MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' " & _
'              "AND STARTDATE = '" & strStartDate & "')"
'
'    Set rsConditions = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTINCOMESTATUS", sSearch)
'
'    m_clsTableAccess.SetRecordSet rsConditions
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'-----------
Dim sSearch As String
Dim rsConditions As ADODB.Recordset

    On Error GoTo Failed

    sSearch = "SELECT VALUENAME, VALUEID From COMBOVALUE " & _
              "WHERE GROUPNAME = 'ApplicationIncomeStatus' " & _
              "AND VALUEID NOT IN " & _
              "(SELECT INCOMESTATUS From MORTGAGEPRODUCTINCOMESTATUS " & _
              "WHERE MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' " & _
              "AND STARTDATE = '" & strStartDate & "')"

    Set rsConditions = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTINCOMESTATUS", sSearch)

    m_clsTableAccess.SetRecordSet rsConditions

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Sub
Public Sub GetLinkedIncomeStatusSettings(strMortgageProductCode As String, strStartDate As String)
'Dim sSearch As String
'Dim rsConditions As ADODB.Recordset
'
'    On Error GoTo Failed
'
'    sSearch = "SELECT VALUENAME AS MORTGAGEINCOMESTATUS, M.* From MORTGAGEPRODUCTINCOMESTATUS M " & _
'              "INNER JOIN COMBOVALUE C ON M.INCOMESTATUS = C.VALUEID " & _
'              "WHERE C.GROUPNAME = 'ApplicationIncomeStatus' " & _
'              "AND MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' " & _
'              "AND STARTDATE = '" & strStartDate & "'"
'
'    Set rsConditions = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTINCOMESTATUS", sSearch)
'
'    m_clsTableAccess.SetRecordSet rsConditions
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'-------------
Dim sSearch As String
Dim rsConditions As ADODB.Recordset

    On Error GoTo Failed
    
    sSearch = "SELECT VALUENAME, VALUEID From COMBOVALUE " & _
              "WHERE GROUPNAME = 'ApplicationIncomeStatus' " & _
              "AND VALUEID IN " & _
              "(SELECT INCOMESTATUS From MORTGAGEPRODUCTINCOMESTATUS " & _
              "WHERE MORTGAGEPRODUCTCODE = '" & strMortgageProductCode & "' " & _
              "AND STARTDATE = '" & strStartDate & "')"

    Set rsConditions = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTINCOMESTATUS", sSearch)
    
    m_clsTableAccess.SetRecordSet rsConditions
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Sub

Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "MORTGAGEPRODUCTCODE"
    colFields.Add "STARTDATE"
'    colFields.Add "INCOMESTATUS"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Friend Function GetMortgageProductCode() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MORTGAGEPRODUCTCODE")
    GetMortgageProductCode = vVal
End Function

Friend Function GetStartDate() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("STARTDATE")
    GetStartDate = vVal
End Function

Friend Function GetIncomeStatus() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("INCOMESTATUS")
    GetIncomeStatus = vVal
End Function

Friend Function SetMortgageProductCode(sVal As Variant) As Boolean
    SetMortgageProductCode = m_clsTableAccess.SetVal("MORTGAGEPRODUCTCODE", CVar(sVal))
End Function

Friend Function SetStartDate(sVal As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("STARTDATE", CVar(sVal))
End Function

Friend Function SetIncomeStatus(sVal As Variant) As Boolean
    SetIncomeStatus = m_clsTableAccess.SetVal("INCOMESTATUS", CVar(sVal))
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
    ' There aren't any...
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




