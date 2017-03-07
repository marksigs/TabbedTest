VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskPriority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : RateTable
' Description   : Contains all data access functions for the BaseRate Table
' Change history
' Prog      Date        Description
' AA        26/02/01    Added class
' PSC       28/02/06    Added AdjustmentMinutes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private m_sSearch As String
Private Const m_sTableName = "TaskPriority"

Implements BandedTable
Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As Collection
    Dim colDupKeys As Collection
    Dim colUpdateFields As Collection
    
    Set colDupKeys = New Collection
    Set colFields = New Collection
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    
    m_clsTableAccess.Create m_sTableName
    
    'Match Fields
    colFields.Add "TASKID"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    ' Banded table initialisation
    Set colUpdateFields = New Collection
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "TASKID"

    m_clsBandedTable.SetUpdateFields colUpdateFields
    
    'Duplicate Keys
    colDupKeys.Add "TASKID"
    colDupKeys.Add "CASEPRIORITY"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
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

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

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

    colListLine.Add "Hello"
    
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

Friend Function GetTaskID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TASKID")
    GetTaskID = CStr(vVal)
End Function
Friend Function GetTaskPriority() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TASKPRIORITY")
    GetTaskPriority = CStr(vVal)
End Function
Friend Function GetTaskPriorityText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TASKPRIORITYTEXT")
    GetTaskPriorityText = CStr(vVal)
End Function
Friend Function GetAdjustmentDays() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ADJUSTMENTDAYS")
    GetAdjustmentDays = CStr(vVal)
End Function
Friend Function GetAdjustmentHours() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ADJUSTMENTHOURS")
    GetAdjustmentHours = CStr(vVal)
End Function

' PSC 28/02/2006 MAR1341 - Start
Friend Function GetAdjustmentMinutes() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ADJUSTMENTMINUTES")
    GetAdjustmentMinutes = CStr(vVal)
End Function
' PSC 28/02/2006 MAR1341 - End

Friend Function SetTaskID(sTaskID As String) As Boolean
    SetTaskID = m_clsTableAccess.SetVal("TASKID", CVar(sTaskID))
End Function
Friend Function SetTaskPriority(nPriority As Long) As Boolean
    SetTaskPriority = m_clsTableAccess.SetVal("TASKPRIORITY", CVar(nPriority))
End Function
Friend Function SetPriorityText(sPriority As String) As Boolean
    SetPriorityText = m_clsTableAccess.SetVal("CASEPRIORITYTEXT", CVar(sPriority))
End Function
Friend Function SetAdjustmentDays(nDays As Long) As Boolean
    SetAdjustmentDays = m_clsTableAccess.SetVal("ADJUSTMENTDAYS", CVar(nDays))
End Function
Friend Function SetAdjustmentHours(nHours As Long) As Boolean
    SetAdjustmentHours = m_clsTableAccess.SetVal("ADJUSTMENTHOURS", CVar(nHours))
End Function
' PSC 28/02/2006 MAR1341 - Start
Friend Function SetAdjustmentMinutes(nMinutes As Long) As Boolean
    SetAdjustmentMinutes = m_clsTableAccess.SetVal("ADJUSTMENTMINUTES", CVar(nMinutes))
End Function
' PSC 28/02/2006 MAR1341 - End
Public Sub ClearChasingPeriod(sField As String)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim vVal As Variant
    Dim nCnt As Long
    
    vVal = Null
    
    Set rs = m_clsTableAccess.GetRecordSet
    
    If Not rs Is Nothing Then
        rs.MoveFirst
        For nCnt = 1 To rs.RecordCount
            rs.fields(sField) = vVal
            rs.MoveNext
        Next nCnt
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


