VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskLinkTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TaskLinkTable
' Description   : Contains all data access functions for the TaskLink Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DB        03/01/03    BM0141 Ensured that tasklist gets promoted along with the businessgroup
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom changes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PB        09/05/2006  EP516  Merged in changes from MAR1545

Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName = "TASKLINK"
Private Const m_sTaskTableName = "TASK"

Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "GROUPID"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
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
Public Function TableAccess_GetUpdateValues() As Collection
End Function
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)

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
Friend Function GetGroupID() As Variant
    GetGroupID = m_clsTableAccess.GetVal("GROUPID")
End Function
Friend Function SetGroupID(vVal As Variant) As Boolean
    SetGroupID = m_clsTableAccess.SetVal("GROUPID", vVal)
End Function
Friend Function GetTaskID() As Variant
    GetTaskID = m_clsTableAccess.GetVal("TASKID")
End Function
Friend Function SetTaskID(vVal As Variant) As Boolean
    SetTaskID = m_clsTableAccess.SetVal("TASKID", vVal)
End Function
Friend Function GetTaskName() As Variant
    GetTaskName = m_clsTableAccess.GetVal("TASKNAME")
End Function
Friend Function SetTaskName(vVal As Variant) As Boolean
    SetTaskName = m_clsTableAccess.SetVal("TASKNAME", vVal)
End Function
Public Sub GetTasksForGroup(sGroupID As String)
    On Error GoTo Failed

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    'DB BM0141 03/01/03
    sSQL = "Select * from " & m_sTableName & _
           " Where GroupID = " & sGroupID
    
    'sSQL = "Select * from " & m_sTableName & _
    '       " Where GroupID = " & g_clsSQLAssistSP.FormatString(sGroupID)
    'DB End
           
    Set rs = g_clsDataAccess.GetTableData(, sSQL, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'EP516  Add code to get all the entries for a Task ID.
Public Sub GetTasksForTaskID(sTaskID As String)
    On Error GoTo Failed

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select * from " & m_sTableName & _
           " Where TaskID = '" & sTaskID & "'"
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetUnassignedTasks(Optional sGroupID As String = "") As ADODB.Recordset
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset

    If Len(sGroupID) > 0 Then

        sSQL = "Select * from " & m_sTaskTableName & _
               " Where Task.TaskID Not In ( Select TaskID from TaskLink where GroupID = " & g_clsSQLAssistSP.FormatString(sGroupID) & ")" & _
               " And DeleteFlag = 0"
    Else
        sSQL = "Select * from " & m_sTaskTableName & " Where DeleteFlag = 0 "
    
    End If
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    Set GetUnassignedTasks = rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function
