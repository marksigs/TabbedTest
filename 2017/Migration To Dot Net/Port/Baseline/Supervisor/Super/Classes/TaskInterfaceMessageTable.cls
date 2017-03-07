VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskInterfaceMessageTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   TaskInterfaceMessageTable
' Description   :   Contains all access functions for the TaskInterfaceMessage table.
' Change history
' Prog      Date        Description
' JD        04/09/2005   MAR40 Created
' PSC       16/12/2005   MAR606 Fix saving of CreateAsComplete field
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "TASKINTERFACEMESSAGE"
Private Const DELETE_FLAG_SET As Integer = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "TASKID"
    m_clsTableAccess.SetKeyMatchFields colFields

End Sub

Friend Function GetTaskID() As Variant
    GetTaskID = CStr(m_clsTableAccess.GetVal("TASKID"))
End Function

Friend Function GetInterfaceType() As Variant
    GetInterfaceType = CStr(m_clsTableAccess.GetVal("INTERFACETYPE"))
End Function

Friend Function GetInterfaceTypeText() As Variant
   GetInterfaceTypeText = CStr(m_clsTableAccess.GetVal("INTERFACETYPETEXT"))
End Function

Friend Function GetMessageType() As Variant
   GetMessageType = CStr(m_clsTableAccess.GetVal("MESSAGETYPE"))
End Function

Friend Function GetMessageTypeText() As Variant
   GetMessageTypeText = CStr(m_clsTableAccess.GetVal("MESSAGETYPETEXT"))
End Function

'PSC 16/12/2005 MAR606 - Start
Friend Function GetCreateAsComplete() As Variant
    GetCreateAsComplete = m_clsTableAccess.GetVal("CreateAsComplete")
End Function
'PSC 16/12/2005 MAR606 - End



Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function

Friend Function SetTaskID(sName As Variant) As Boolean
    SetTaskID = m_clsTableAccess.SetVal("TASKID", CVar(sName))
End Function

Friend Function SetInterfaceType(sTaskName As Variant) As Boolean
    SetInterfaceType = m_clsTableAccess.SetVal("INTERFACETYPE", CVar(sTaskName))
End Function

Friend Function SetMessageType(sValue As Variant) As Boolean
    SetMessageType = m_clsTableAccess.SetVal("MESSAGETYPE", CVar(sValue))
End Function

'PSC 16/12/2005 MAR606 - Start
Friend Function SetCreateAsComplete(sValue As Variant) As Boolean
    SetCreateAsComplete = m_clsTableAccess.SetVal("CreateAsComplete", CVar(sValue))
End Function
'PSC 16/12/2005 MAR606 - Start

Public Function GetTaskField() As String
    GetTaskField = "TASKNAME"
End Function

Public Function GetComboFields() As Collection
    On Error GoTo Failed
    Dim colFields As Collection
    
    Set colFields = New Collection
    
    colFields.Add "TASKNAME"
    colFields.Add "TASKID"
    
    Set GetComboFields = colFields
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function




Public Sub DeleteTask()
    
    Dim sSQL As String
    Dim rsCaseTasks As ADODB.Recordset
    
    On Error GoTo Failed
    
    'STB: SYS3149 - Stop Tasks from being delete if there are CastTasks in existance.
    sSQL = "SELECT CASEACTIVITYGUID, STAGEID, TASKID FROM CASETASK WHERE TASKID = " & g_clsSQLAssistSP.FormatString(m_clsTableAccess.GetKeyMatchValues(1))
    
    'Obtain any existing case tasks with this TaskID.
    Set rsCaseTasks = g_clsDataAccess.GetTableData(, sSQL)
    
    If Not (rsCaseTasks.BOF And rsCaseTasks.EOF) Then
        g_clsErrorHandling.RaiseError errGeneralError, "You cannot delete this task as there are related CaseTasks in existance."
    End If
    'STB: SYS3149 - End.
    
    'GetTasks , , m_clsTableAccess.GetKeyMatchValues(), False
    m_clsTableAccess.GetTableData POPULATE_KEYS
    If m_clsTableAccess.RecordCount() > 0 Then
        'SetDeleteFlag True
        m_clsTableAccess.Update
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetTaskInterfaceMessage(sTaskID As String, Optional nInterfaceType As Long = 0, Optional nMessageType As Long = 0)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsTaskInterfaceMessages As ADODB.Recordset
    
    sSearch = "SELECT * FROM " + m_sTableName + " WHERE TASKID ='" + sTaskID + "'"
    If nInterfaceType > 0 Then
        sSearch = sSearch + " AND INTERFACETYPE= " + CStr(nInterfaceType)
    End If
    If nMessageType > 0 Then
        sSearch = sSearch + " AND MESSAGETYPE= " + CStr(nMessageType)
    End If

    Set rsTaskInterfaceMessages = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rsTaskInterfaceMessages
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
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
    colListLine.Add GetTaskID()
    colListLine.Add GetInterfaceTypeText()
    colListLine.Add GetMessageTypeText()
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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

