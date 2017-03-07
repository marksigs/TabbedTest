VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   TaskTable
' Description   :   Contains all access functions for the Task table.
' Change history
' Prog      Date        Description
' DJP       10/11/00    Created
' STB       05/04/02    SYS3149 - Orphaned CaseTasks are blocked.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS
' Prog      Date        Description
' AW        13/05/02    BM092   Added Get/Set for ConfirmPrint
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' JD        03/09/2005  MAR32   TAS changes.
' RF        13/12/2005  MAR202 Handle packs
' PSC       28/02/2006  MAR1341 Add CHASINGPERIODMINUTES
' AW        31/08/2006  EP1110 Added EDITABLETASKIND, PROGRESSTASKIND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM2 Change History:
'
' Prog      Date        Ref     Description
' TW        12/12/2006  EP2_453 - E2CR63 Remote Underwriting add Remote Owner Task (Checkbox; default = checked)
' GHun      02/10/2007  OMIGA00003234 Added RuleFrequency
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "TASK"
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

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "TASKNAME"
End Sub

Friend Function GetTaskID() As Variant
    GetTaskID = CStr(m_clsTableAccess.GetVal("TASKID"))
End Function

Friend Function GetTaskName() As Variant
    GetTaskName = CStr(m_clsTableAccess.GetVal("TASKNAME"))
End Function

Friend Function GetTaskTypeText() As Variant
   GetTaskTypeText = CStr(m_clsTableAccess.GetVal("TASKTYPETEXT"))
End Function

Friend Function GetTaskType() As Variant
   GetTaskType = CStr(m_clsTableAccess.GetVal("TASKTYPE"))
End Function

Friend Function GetChasingPeriodDays() As Variant
   GetChasingPeriodDays = CStr(m_clsTableAccess.GetVal("CHASINGPERIODDAYS"))
End Function

Friend Function GetChasingPeriodHours() As Variant
   GetChasingPeriodHours = CStr(m_clsTableAccess.GetVal("CHASINGPERIODHOURS"))
End Function

' PSC 28/02/2006 MAR1341 - Start
Friend Function GetChasingPeriodMinutes() As Variant
   GetChasingPeriodMinutes = CStr(m_clsTableAccess.GetVal("CHASINGPERIODMINUTES"))
End Function
' PSC 28/02/2006 MAR1341 - End

Friend Function GetChasingTask() As Variant
   GetChasingTask = CStr(m_clsTableAccess.GetVal("CHASINGTASK"))
End Function

Friend Function GetContactType() As Variant
   GetContactType = CStr(m_clsTableAccess.GetVal("CONTACTTYPE"))
End Function

Friend Function GetOwnerTypeText() As Variant
   GetOwnerTypeText = CStr(m_clsTableAccess.GetVal("OWNERTYPETEXT"))
End Function

Friend Function GetOwnerType() As Variant
   GetOwnerType = CStr(m_clsTableAccess.GetVal("TASKOWNERTYPE"))
End Function

Friend Function GetInputProcess() As Variant
   GetInputProcess = CStr(m_clsTableAccess.GetVal("INPUTPROCESS"))
End Function

Friend Function GetInterface() As Variant
   GetInterface = CStr(m_clsTableAccess.GetVal("INTERFACE"))
End Function

Friend Function GetOutputDocument() As Variant
   GetOutputDocument = CStr(m_clsTableAccess.GetVal("OUTPUTDOCUMENT"))
End Function

Friend Function GetRuleRef() As Variant
   GetRuleRef = CStr(m_clsTableAccess.GetVal("TASKRULEREFERENCE"))
End Function

Friend Function GetApplicantTask() As Variant
    GetApplicantTask = CStr(m_clsTableAccess.GetVal("APPLICANT"))
End Function

Friend Function GetNotApplicable() As Variant
    GetNotApplicable = CStr(m_clsTableAccess.GetVal("NOTAPPLICABLEFLAG"))
End Function
' AW 31/08/2006 EP1110 - Start
Friend Function GetEditableTaskInd() As Variant
   GetEditableTaskInd = CStr(m_clsTableAccess.GetVal("EDITABLETASKIND"))
End Function

Friend Function GetProgressTaskInd() As Variant
   GetProgressTaskInd = CStr(m_clsTableAccess.GetVal("PROGRESSTASKIND"))
End Function
' AW 31/08/2006 EP1110 - End
Friend Function GetConfirmPrint() As Variant
    GetConfirmPrint = CStr(m_clsTableAccess.GetVal("CONFIRMPRINTIND"))
End Function

Friend Function GetAutomaticTask() As Variant
    GetAutomaticTask = CStr(m_clsTableAccess.GetVal("AUTOMATICTASKIND"))
End Function
Friend Function GetAlwaysAutoTask() As Variant
    GetAlwaysAutoTask = CStr(m_clsTableAccess.GetVal("ALWAYSAUTOMATICONCREATION"))
End Function
Friend Function GetTASTask() As Variant
    GetTASTask = CStr(m_clsTableAccess.GetVal("TASENABLED"))
End Function
Friend Function GetTASQueueName() As Variant
    GetTASQueueName = CStr(m_clsTableAccess.GetVal("TASPROCESSQUEUENAME"))
End Function
Friend Function GetTASTriggerNxtStage() As Variant
    GetTASTriggerNxtStage = CStr(m_clsTableAccess.GetVal("TASLASTTASKALLOWSTAGEMOVE"))
End Function

Friend Function GetCustomerInvolved() As Variant
    GetCustomerInvolved = CStr(m_clsTableAccess.GetVal("CONTACTCUSTOMERIND"))
End Function

Friend Function GetRemoteOwnerTask() As Variant
    GetRemoteOwnerTask = CStr(m_clsTableAccess.GetVal("REMOTEOWNERTASKIND"))
End Function
Friend Function GetCustomerTask() As Variant
    GetCustomerTask = CStr(m_clsTableAccess.GetVal("CUSTOMERTASK"))
End Function

Friend Function GetPackControlNumber() As Variant
    GetPackControlNumber = CStr(m_clsTableAccess.GetVal("PACKCONTROLNUMBER"))
End Function

'OMIGA00003234 GHun
Friend Function GetRuleFrequency() As Variant
    GetRuleFrequency = CStr(m_clsTableAccess.GetVal("RULEFREQUENCY"))
End Function
'OMIGA00003234 End

Friend Function SetTaskID(sName As Variant) As Boolean
    SetTaskID = m_clsTableAccess.SetVal("TASKID", CVar(sName))
End Function

Friend Function SetTaskName(sTaskName As Variant) As Boolean
    SetTaskName = m_clsTableAccess.SetVal("TASKNAME", CVar(sTaskName))
End Function

Friend Function SetTaskType(sTaskType As Variant) As Boolean
    SetTaskType = m_clsTableAccess.SetVal("TASKTYPE", CVar(sTaskType))
End Function

Friend Function SetChasingPeriodDays(sChasingPeriod As Variant) As Boolean
    SetChasingPeriodDays = m_clsTableAccess.SetVal("CHASINGPERIODDAYS", CVar(sChasingPeriod))
End Function

Friend Function SetChasingPeriodHours(sChasingPeriod As Variant) As Boolean
    SetChasingPeriodHours = m_clsTableAccess.SetVal("CHASINGPERIODHOURS", CVar(sChasingPeriod))
End Function

' PSC 28/02/2006 MAR1341 - Start
Friend Function SetChasingPeriodMinutes(sChasingPeriod As Variant) As Boolean
    SetChasingPeriodMinutes = m_clsTableAccess.SetVal("CHASINGPERIODMINUTES", CVar(sChasingPeriod))
End Function
' PSC 28/02/2006 MAR1341 - End

Friend Function SetChasingTask(sChasingTask As Variant) As Boolean
    SetChasingTask = m_clsTableAccess.SetVal("CHASINGTASK", CVar(sChasingTask))
End Function

Friend Function SetContactType(sContactType As Variant) As Boolean
    SetContactType = m_clsTableAccess.SetVal("CONTACTTYPE", CVar(sContactType))
End Function

Friend Function SetInputProcess(sInputProcess As Variant) As Boolean
    SetInputProcess = m_clsTableAccess.SetVal("INPUTPROCESS", CVar(sInputProcess))
End Function

Friend Function SetOutputDocument(sOutputDocument As Variant) As Boolean
    SetOutputDocument = m_clsTableAccess.SetVal("OUTPUTDOCUMENT", CVar(sOutputDocument))
End Function

Friend Function SetOwnerType(sTaskType As Variant) As Boolean
    SetOwnerType = m_clsTableAccess.SetVal("TASKOWNERTYPE", CVar(sTaskType))
End Function

Friend Function SetRuleRef(sRuleRef As Variant) As Boolean
    SetRuleRef = m_clsTableAccess.SetVal("TASKRULEREFERENCE", CVar(sRuleRef))
End Function

Friend Function SetApplicantTask(sApplicantTask As Variant) As Boolean
    SetApplicantTask = m_clsTableAccess.SetVal("APPLICANT", CVar(sApplicantTask))
End Function

Friend Function SetNotApplicable(sNotApplicable As Variant) As Boolean
    SetNotApplicable = m_clsTableAccess.SetVal("NOTAPPLICABLEFLAG", CVar(sNotApplicable))
End Function
' AW 31/08/2006 EP1110 - Start
Friend Function SetEditableTaskInd(sEditableTaskInd As Variant) As Boolean
    SetEditableTaskInd = m_clsTableAccess.SetVal("EDITABLETASKIND", CVar(sEditableTaskInd))
End Function

Friend Function SetProgressTaskInd(sProgressTaskInd As Variant) As Boolean
    SetProgressTaskInd = m_clsTableAccess.SetVal("PROGRESSTASKIND", CVar(sProgressTaskInd))
End Function
' AW 31/08/2006 EP1110 - End
Friend Function SetConfirmPrint(sConfirmPrint As Variant) As Boolean
    SetConfirmPrint = m_clsTableAccess.SetVal("CONFIRMPRINTIND", CVar(sConfirmPrint))
End Function

Friend Function SetAutomaticTask(sAutomaticTask As Variant) As Boolean
    SetAutomaticTask = m_clsTableAccess.SetVal("AUTOMATICTASKIND", CVar(sAutomaticTask))
End Function
Friend Function SetAlwaysAutoTask(sAlwaysAutoTask As Variant) As Boolean
    SetAlwaysAutoTask = m_clsTableAccess.SetVal("ALWAYSAUTOMATICONCREATION", CVar(sAlwaysAutoTask))
End Function
Friend Function SetTASTask(sTAS As Variant) As Boolean
    SetTASTask = m_clsTableAccess.SetVal("TASENABLED", CVar(sTAS))
End Function
Friend Function SetTASQueueName(sQueueName As Variant) As Boolean
    SetTASQueueName = m_clsTableAccess.SetVal("TASPROCESSQUEUENAME", CVar(sQueueName))
End Function
Friend Function SetTASTriggerNxtStage(sTrigger As Variant) As Boolean
    SetTASTriggerNxtStage = m_clsTableAccess.SetVal("TASLASTTASKALLOWSTAGEMOVE", CVar(sTrigger))
End Function

Friend Function SetCustomerInvolved(sCustomerInvolved As Variant) As Boolean
    SetCustomerInvolved = m_clsTableAccess.SetVal("CONTACTCUSTOMERIND", CVar(sCustomerInvolved))
End Function

Friend Function SetRemoteOwnerTask(sRemoteOwnerTask As Variant) As Boolean
    SetRemoteOwnerTask = m_clsTableAccess.SetVal("REMOTEOWNERTASKIND", CVar(sRemoteOwnerTask))
End Function
Friend Function SetCustomerTask(sCustomerTask As Variant) As Boolean
    SetCustomerTask = m_clsTableAccess.SetVal("CUSTOMERTASK", CVar(sCustomerTask))
End Function


Friend Function SetInterface(sInterface As Variant) As Boolean
    SetInterface = m_clsTableAccess.SetVal("INTERFACE", CVar(sInterface))
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

Friend Function SetPackControlNumber(sPackControlNumber As Variant) As Boolean
    SetPackControlNumber = m_clsTableAccess.SetVal("PACKCONTROLNUMBER", CVar(sPackControlNumber))
End Function

'OMIGA00003234 GHun
Friend Function SetRuleFrequency(ByVal strRuleFrequency As String) As Boolean
    SetRuleFrequency = m_clsTableAccess.SetVal("RULEFREQUENCY", CVar(strRuleFrequency))
End Function
'OMIGA00003234 End

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

Public Sub GetTasksWithCombos()
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsTasks As ADODB.Recordset
    
    sSearch = "SELECT " & _
        "TASK.*, C1.VALUENAME ""TASKTYPETEXT"", C2.VALUENAME ""OWNERTYPETEXT"" " & _
    "FROM " & _
        "TASK , COMBOVALUE C1, COMBOVALUE C2 " & _
    "WHERE " & _
        "DELETEFLAG <> 1 AND TASKTYPE = C1.VALUEID  AND C1.GROUPNAME = " & g_clsSQLAssistSP.FormatString("TaskType") & "AND " & _
        "TASKOWNERTYPE = C2.VALUEID AND C2.GROUPNAME = " & g_clsSQLAssistSP.FormatString("UserRole")
    
    Set rsTasks = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rsTasks
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetTasksNotForStage(Optional sStageID As String = "")
    On Error GoTo Failed
    
    If Len(sStageID) = 0 Then
        GetTasks
        GetTasksNotForStage sStageID
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

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
        SetDeleteFlag True
        m_clsTableAccess.Update
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetTasks(Optional colExcludeTasks As Collection, Optional ByVal sExcludeStageID As String = "", Optional colTaskMatch As Collection = Nothing, Optional bCheckDeleteFlag As Boolean = True)
    On Error GoTo Failed
    Dim nCount As Long
    Dim nThisItem As Long
    Dim sSearch As String
    Dim sTaskID As String
    Dim sExcludeTask As Variant
    Dim rsTasks As ADODB.Recordset
    Dim bExcludeStage As Boolean
    
    bExcludeStage = Len(sExcludeStageID) > 0
    
    If bCheckDeleteFlag Then
        sSearch = "SELECT * FROM " + m_sTableName + " WHERE DELETEFLAG <> " & DELETE_FLAG_SET
    Else
        sSearch = "SELECT * FROM " + m_sTableName
    End If
    
    If Not colTaskMatch Is Nothing Then
        If colTaskMatch.Count > 0 Then
            If bCheckDeleteFlag Then
                sSearch = sSearch + " AND ("
            Else
                sSearch = sSearch + " WHERE ("
            End If
            
            nCount = colTaskMatch.Count
            For nThisItem = 1 To nCount
                sTaskID = colTaskMatch(nThisItem)
                sTaskID = g_clsSQLAssistSP.FormatString(sTaskID)
                sSearch = sSearch + " TASKID = " & sTaskID
                            
                If nThisItem < nCount Then
                    sSearch = sSearch + " OR "
                End If
            
            Next
            sSearch = sSearch + ")"
        End If
    End If
    
    If Not colExcludeTasks Is Nothing Then
        For Each sExcludeTask In colExcludeTasks
            If Len(sExcludeTask) > 0 Then
                sExcludeTask = g_clsSQLAssistSP.FormatString(sExcludeTask)
                sSearch = sSearch & " AND TASKID <> " & sExcludeTask
            End If
        Next
    End If
    
    If bExcludeStage Then
        sSearch = sSearch & " AND " & m_sTableName & ".TASKID NOT IN ( " & _
        "SELECT TASKID FROM STAGETASK WHERE STAGEID = " & g_clsSQLAssistSP.FormatString(sExcludeStageID) & " AND " & _
        "STAGETASK.DELETEFLAG = " & DELETE_FLAG_SET & ")"
    End If
    
    sSearch = sSearch + " ORDER BY TASKNAME"
    
    Set rsTasks = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    m_clsTableAccess.SetRecordSet rsTasks
    
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
    colListLine.Add GetTaskName()
    colListLine.Add GetTaskTypeText()
    colListLine.Add GetOwnerTypeText()
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

