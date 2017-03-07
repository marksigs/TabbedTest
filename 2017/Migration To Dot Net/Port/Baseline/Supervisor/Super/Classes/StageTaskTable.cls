VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StageTaskTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   StageTaskTable
' Description   :   Contains all access functions for the Stage table.
' Change history
' Prog      Date        Description
' DJP       10/11/00    Created
' STB       10/04/02    SYS3878 Added CropDeletedTasks.
' GHun      07/03/2006  MAR1300 Added TaskProcessAuthorityLevel
' PB        24/04/2006  EP367   Changes merged into Epsom Supervisor
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Implements BandedTable

'Private data
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "STAGETASK"
Private Const DELETE_FLAG_SET As Integer = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As Collection
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "STAGEID"
    colFields.Add "TASKID"
    
    m_clsTableAccess.SetKeyMatchFields colFields

    Set colUpdateFields = New Collection
    colUpdateFields.Add "STAGEID"
    m_clsBandedTable.SetUpdateFields colUpdateFields
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    ' Order criteria used when executing queries
    'm_clsTableAccess.SetOrderCriteria ""
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetStageID() As Variant
    GetStageID = CStr(m_clsTableAccess.GetVal("STAGEID"))
End Function
Friend Function GetTaskID() As Variant
    GetTaskID = CStr(m_clsTableAccess.GetVal("TASKID"))
End Function
Public Function GetDefaultTask() As Variant
    GetDefaultTask = CStr(m_clsTableAccess.GetVal("DEFAULTFLAG"))
End Function
Public Function GetAdditionalTask() As Variant
    GetAdditionalTask = CStr(m_clsTableAccess.GetVal("ADDITIONALFLAG"))
End Function
Public Function GetMandatory() As Variant
    GetMandatory = CStr(m_clsTableAccess.GetVal("MANDATORYFLAG"))
End Function
Public Function GetCarriedForward() As Variant
    GetCarriedForward = CStr(m_clsTableAccess.GetVal("CARRYFORWARDINDICATOR"))
End Function
Public Function GetAuthorityLevel() As Variant
    GetAuthorityLevel = CStr(m_clsTableAccess.GetVal("USERAUTHORITYLEVEL"))
End Function
Public Function GetAuthorityLevelText() As Variant
    GetAuthorityLevelText = CStr(m_clsTableAccess.GetVal("USERAUTHORITYLEVELTEXT"))
End Function
Public Function GetProcessAuthorityLevel() As Variant
    GetProcessAuthorityLevel = CStr(m_clsTableAccess.GetVal("TASKPROCESSAUTHORITYLEVEL"))
End Function
Public Function GetProcessAuthorityLevelText() As Variant
    GetProcessAuthorityLevelText = CStr(m_clsTableAccess.GetVal("TASKPROCESSAUTHORITYLEVELTEXT"))
End Function
Public Function GetUserID() As Variant
    GetUserID = CStr(m_clsTableAccess.GetVal("TASKUSERID"))
End Function
Public Function GetUnitID() As Variant
    GetUnitID = CStr(m_clsTableAccess.GetVal("TASKUNITID"))
End Function
Public Function GetTriggerTask() As Variant
    GetTriggerTask = CStr(m_clsTableAccess.GetVal("TRIGGERTASKIND"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetStageID(sStageID As Variant) As Boolean
    SetStageID = m_clsTableAccess.SetVal("STAGEID", CVar(sStageID))
End Function
Friend Function SetTaskID(sTaskID As Variant) As Boolean
    SetTaskID = m_clsTableAccess.SetVal("TASKID", CVar(sTaskID))
End Function
Friend Function SetDefaultTask(sDefaultTask As Variant) As Boolean
    SetDefaultTask = m_clsTableAccess.SetVal("DEFAULTFLAG", CVar(sDefaultTask))
End Function
Friend Function SetAdditionalTask(sAdditionalTask As Variant) As Boolean
    SetAdditionalTask = m_clsTableAccess.SetVal("ADDITIONALFLAG", CVar(sAdditionalTask))
End Function
Friend Function SetMandatory(sMandatory As Variant) As Boolean
    SetMandatory = m_clsTableAccess.SetVal("MANDATORYFLAG", CVar(sMandatory))
End Function
Friend Function SetCarriedForward(sCarriedForward As Variant) As Boolean
    SetCarriedForward = m_clsTableAccess.SetVal("CARRYFORWARDINDICATOR", CVar(sCarriedForward))
End Function
Friend Function SetAuthorityLevel(sAuthorityLevel As Variant) As Boolean
    SetAuthorityLevel = m_clsTableAccess.SetVal("USERAUTHORITYLEVEL", CVar(sAuthorityLevel))
End Function
Friend Function SetProcessAuthorityLevel(sAuthorityLevel As Variant) As Boolean
    SetProcessAuthorityLevel = m_clsTableAccess.SetVal("TASKPROCESSAUTHORITYLEVEL", CVar(sAuthorityLevel))
End Function
Friend Function SetUserID(sUserID As Variant) As Boolean
    SetUserID = m_clsTableAccess.SetVal("TASKUSERID", CVar(sUserID))
End Function
Friend Function SetUnitID(sUnitID As Variant) As Boolean
    SetUnitID = m_clsTableAccess.SetVal("TASKUNITID", CVar(sUnitID))
End Function
Friend Function SetTriggerTask(sTriggerTask As Variant) As Boolean
    SetTriggerTask = m_clsTableAccess.SetVal("TRIGGERTASKIND", sTriggerTask)
End Function
Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTaskField() As String
    GetTaskField = "TASKNAME"
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetTasksForStage
' Description   :   Returns a list of all tasks for the stage passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MAR1120  Add parameter
Public Sub GetTasksForStage(ByVal sStageID As String, Optional bIgnoreDeleteFlag As Boolean = False)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsTasks As ADODB.Recordset
    Dim sTarget As String
    
    sStageID = g_clsSQLAssistSP.FormatString(sStageID)

    sSearch = "SELECT * FROM " & m_sTableName & " WHERE "
    
    If bIgnoreDeleteFlag Then
        sSearch = sSearch & "STAGEID = " & sStageID
    Else
        sSearch = sSearch & "STAGEID = " & sStageID & " AND DELETEFLAG <> " & DELETE_FLAG_SET
    End If
    
    sTarget = m_clsTableAccess.GetDatabase()
    Set rsTasks = g_clsDataAccess.GetTableData(m_sTableName, sSearch, sTarget)

    m_clsTableAccess.SetRecordSet rsTasks

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
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
    colListLine.Add GetTaskID()
    colListLine.Add GetBooleanFromNumber(GetDefaultTask())
    colListLine.Add GetBooleanFromNumber(GetAdditionalTask())
    colListLine.Add GetBooleanFromNumber(GetMandatory())
    colListLine.Add GetBooleanFromNumber(GetCarriedForward())
    colListLine.Add GetBooleanFromNumber(GetTriggerTask())
    colListLine.Add GetAuthorityLevel()
    colListLine.Add GetProcessAuthorityLevel()   'MAR1300 GHun
    colListLine.Add GetUserID()
    colListLine.Add GetUnitID()
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
Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub
Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
End Sub
Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub
Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub





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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function         :   StageTaskTable
' Description      :   Sets the deleteflag to true for all stagetask records with a s
'                      stageid = to the one passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteStageTaskForStage(sStageID As String)
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim nThisRecord As Long
    
    ' DJP SQL Server port - SQL Server needs the primary key to be present in an update
    sSQL = "SELECT " & _
                "DELETEFLAG, " & _
                "STAGEID, " & _
                "TASKID " & _
            "FROM " & _
                m_sTableName & " " & _
            "WHERE " & _
                "STAGEID = " & g_clsSQLAssistSP.FormatString(sStageID)
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL, m_clsTableAccess.GetDatabase)
    
    If rs.RecordCount > 0 Then
        m_clsTableAccess.SetRecordSet rs
        m_clsTableAccess.MoveFirst
        For nThisRecord = 1 To rs.RecordCount
            SetDeleteFlag True
            m_clsTableAccess.MoveNext
        Next
        
        ' DJP SQL Server port, only update if there are records to delete
        m_clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : CropDeletedTasks
' Description : Delete any records with a DELETEFLAG = 1 which match the TaskID specified.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub CropDeletedTasks(ByVal sStageID As String, ByVal sTaskID As String)
       
    Dim sSQL As String
        
    'Build a query to obtain all the StageTasks.
    sSQL = "DELETE FROM STAGETASK WHERE STAGEID = " & g_clsSQLAssistSP.FormatString(sStageID) & " AND TASKID = " & g_clsSQLAssistSP.FormatString(sTaskID) & " AND DELETEFLAG = 1"
    
    'Get a recordset with all the StageTasks for this StageID
    g_clsDataAccess.ExecuteCommand (sSQL)
        

End Sub
