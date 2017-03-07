VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchScheduleTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BatchTable
' Description   : Contains all data access functions for the BatchSchedule Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DJP       26/06/01    SQL Server port
' SA        18/01/02    SYS3327 GetRetryIndicator - if null - send back zero.
' SA        14/02/02    SYS4073 Validation codes for Program type and Schedule Status required
' GHun      18/04/02    SYS4424 In Supervisor you cant view errors in incomplete batch runs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sBatchNumber As String
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName = "BATCHSCHEDULE"
Private Const m_sBatchTableName = "BATCH"
'SYS4424 the BATCHAUDIT table is required in the SQL query
Private Const m_sBatchAuditTableName = "BATCHAUDIT"
Private Const m_sStatusComboGroup = "BatchStatus"
Private Const m_sProgramTypeComboGroup = "BatchProgramType"
Private Const m_sScheduleStatusComboGroup = "BatchScheduleStatus"

Implements TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GHun      24/03/2003  BM0438 Batch problems - BATCHSCHEDULE table design
' GHun      24/03/2003  BM0425 Don't lock records when reading from the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BM0438
Private Const m_sBatchProgressTableName = "BATCHSCHEDULEPROGRESS"
'BM0438 End

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.SetKeyMatchFields colFields
    m_clsTableAccess.Create m_sTableName

    
    colFields.Add "BATCHNUMBER"
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
    
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

Public Function TableAccess_GetUpdateValues() As Collection
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)

    Dim vTmp As Variant
    
    colListLine.Add GetBatchRunNumber
    colListLine.Add GetBatchStatusText
    colListLine.Add GetBatchCreationDate
    vTmp = GetBatchCompletionDate
    If CDate(vTmp) > 0 Then
        colListLine.Add vTmp
    Else
        colListLine.Add ""
    End If
    colListLine.Add GetBatchExecutionDate
    colListLine.Add GetNumberOfRecords
    colListLine.Add GetNumberOfSuccesses
    colListLine.Add GetNumberOfFailures
    
    'SYS4424 Batch Schedule error information should also be shown
    If GetErrorNumber <> 0 Then
        colListLine.Add GetErrorNumber
    Else
        colListLine.Add ""
    End If
    colListLine.Add GetErrorSource
    colListLine.Add GetErrorDescription
    'SYS4424
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

Friend Function SetBatchNumber(sBatchrunNo As String) As Boolean
    SetBatchNumber = m_clsTableAccess.SetVal("BATCHRUNNUMBER", CVar(sBatchrunNo))
End Function

Friend Function SetBatchRunNumber(sBatchNo As String) As Boolean
    SetBatchRunNumber = m_clsTableAccess.SetVal("BATCHNUMBER", CVar(sBatchNo))
End Function

Friend Function SetStatus(nStatus As Long) As Boolean
    SetStatus = m_clsTableAccess.SetVal("STATUS", CVar(nStatus))
End Function

Friend Function SetExecutionDate(dExecutionDate As Date) As Boolean
    SetExecutionDate = m_clsTableAccess.SetVal("EXECUTIONDATETIME", CVar(dExecutionDate))
End Function

Friend Function SetCreationDate(dCreationDate As Date) As Boolean
    SetCreationDate = m_clsTableAccess.SetVal("CREATIONDATETIME", CVar(dCreationDate))
End Function

Friend Function SetCompletionDate(dCompletionDate As Date) As Boolean
    SetCompletionDate = m_clsTableAccess.SetVal("COMPLETIONDATETIME", CVar(dCompletionDate))
End Function

Friend Function SetNumberOfRecords(nRecords As Long) As Boolean
    SetNumberOfRecords = m_clsTableAccess.SetVal("TOTALNUMBEROFRECORDS", CVar(nRecords))
End Function

Friend Function SetNumberOfFailures(nFailures As Long) As Boolean
    SetNumberOfFailures = m_clsTableAccess.SetVal("TOTALNUMBEROFFAILURES", CVar(nFailures))
End Function

Friend Function SetNumberOfSuccesses(nSuccesses As Long) As Boolean
    SetNumberOfSuccesses = m_clsTableAccess.SetVal("TOTALNUMBEROFSUCCESSES", CVar(nSuccesses))
End Function

Friend Function SetRetryBatchIndicator(nIndicator As Integer) As Boolean
    SetRetryBatchIndicator = m_clsTableAccess.SetVal("RETRYBATCHINDICATOR", CVar(nIndicator))
End Function

Friend Function SetFailedBatchRunNumber(nFailedBatchNumber As Integer) As Boolean
    SetFailedBatchRunNumber = m_clsTableAccess.SetVal("FAILEDBATCHRUNNUMBER", CVar(nFailedBatchNumber))
End Function

Public Function GetBatchNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BATCHNUMBER")
    GetBatchNumber = CStr(vVal)
End Function

Public Function GetBatchRunNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BATCHRUNNUMBER")
    GetBatchRunNumber = CStr(vVal)
End Function

Public Function GetBatchStatus() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("STATUS")
    GetBatchStatus = CLng(vVal)
End Function

Public Function GetBatchStatusText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("STATUSTEXT")
    GetBatchStatusText = CStr(vVal)
End Function

Public Function GetBatchCreationDate() As Date
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CREATIONDATETIME")
    GetBatchCreationDate = CDate(vVal)
End Function

Public Function GetBatchExecutionDate() As Date
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EXECUTIONDATETIME")
    GetBatchExecutionDate = CDate(vVal)
End Function

Public Function GetBatchCompletionDate() As Date
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("COMPLETIONDATETIME")
    If IsDate(vVal) Then
        GetBatchCompletionDate = CDate(vVal)
    Else
        GetBatchCompletionDate = 0
    End If
End Function

Public Function GetNumberOfRecords() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TOTALNUMBEROFRECORDS")
    GetNumberOfRecords = CLng(vVal)
End Function

Public Function GetNumberOfFailures() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TOTALNUMBEROFFAILURES")
    GetNumberOfFailures = CLng(vVal)
End Function

Public Function GetNumberOfSuccesses() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TOTALNUMBEROFSUCCESSES")
    GetNumberOfSuccesses = CLng(vVal)
End Function

Public Function GetBatchRetryIndicator() As Integer
    'SYS3327 If the RetryBatchIndicator is Null, we can't convert it!
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("RETRYBATCHINDICATOR")
    If vVal <> "" Then
        GetBatchRetryIndicator = CInt(vVal)
    Else
        ' Send back a zero - the failed batch has not been re-run.
        GetBatchRetryIndicator = 0
    End If
End Function

Public Function GetFailedBatchRunNumber() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FAILEDBATCHRUNNUMBER")
    GetFailedBatchRunNumber = CLng(vVal)
End Function

Public Function GetProgramType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PROGRAMTYPE")
    GetProgramType = CStr(vVal)
End Function

Public Function GetProgValidationType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ProgValidType")
    GetProgValidationType = CStr(vVal)
End Function

Public Function GetStatusValidationType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("StatusValidType")
    GetStatusValidationType = CStr(vVal)
End Function

Public Function GetProgramTypeText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PROGRAMTYPETEXT")
    GetProgramTypeText = CStr(vVal)
End Function

'BM0425 the returned recordset is never used. OOSS does allow a function to be changed to
'       a Sub so the unused return value is set to the smallest data type available (byte)
'Public Function GetBatchScheduleTableData(sBatchNumber As String) As ADODB.Recordset
Public Function GetBatchScheduleTableData(ByVal sBatchNumber As String) As Byte
'BM0425 End

On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    'BM0438 Join on BatchScheduleProgress
    'BM0425 Remove redundant SYS4424 subquery and use NOLOCK locking hint
    'SYS4073 Select Validation type for Program type and Schedule Status.
    'SYS4424 Added AuditErrorsFound subquery, so we know when to enable the View Errors button
    'sSQL = "SELECT " & m_sTableName & ".*, ComboStatus.ValueName StatusText, ComboProgram.ValueName ProgramTypeText, ComboValidProgType.ValidationType ProgValidType, ComboValidStatus.ValidationType StatusValidType, " & _
           "(SELECT Count(*) FROM " & m_sBatchAuditTableName & " WHERE " & m_sBatchAuditTableName & ".BatchNumber = " & m_sTableName & ".BatchNumber AND " & m_sBatchAuditTableName & ".BatchRunNumber = " & m_sTableName & ".BatchRunNumber AND NOT(" & m_sBatchAuditTableName & ".ErrorNumber is null)) AuditErrorsFound " & _
           "FROM " & m_sTableName & "," & m_sBatchTableName & ", ComboValue ComboStatus, ComboValue ComboProgram, Combovalidation ComboValidProgType, " & _
           " ComboValidation ComboValidStatus " & _
           "WHERE " & m_sTableName & ".BatchNumber = " & g_clsSQLAssistSP.FormatString(sBatchNumber) & " " & _
           "AND " & m_sTableName & ".BatchNumber = " & m_sBatchTableName & ".BatchNumber " & _
           "AND ComboStatus.ValueID = " & m_sTableName & ".Status AND ComboStatus.GroupName = " & g_clsSQLAssistSP.FormatString(m_sScheduleStatusComboGroup) & " " & _
           "AND ComboProgram.ValueID = Batch.ProgramType AND ComboProgram.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & " " & _
           "AND ComboValidProgType.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & " " & _
           "AND ComboValidProgType.ValueId = Batch.ProgramType " & _
           "AND ComboValidStatus.GroupName = " & g_clsSQLAssistSP.FormatString(m_sScheduleStatusComboGroup) & " " & _
           "AND ComboValidStatus.ValueId = " & m_sTableName & ".Status " & _
           "ORDER BY BatchRunNumber"
    sSQL = "SELECT BS.BatchNumber, BS.BatchRunNumber, BS.Status, BS.CreationDateTime, BS.ExecutionDateTime, " & _
        "BS.CompletionDateTime, BS.RetryBatchIndicator, BS.FailedBatchRunNumber, BS.ErrorNumber, BS.ErrorSource, BS.ErrorDescription, " & _
        "BSP.TotalNumberOfRecords, BSP.TotalNumberOfSuccesses, BSP.TotalNumberOfFailures," & _
        "ComboStatus.ValueName StatusText, ComboProgram.ValueName ProgramTypeText, ComboValidProgType.ValidationType ProgValidType, ComboValidStatus.ValidationType StatusValidType " & _
        "FROM " & m_sTableName & " BS WITH (NOLOCK)" & _
        "INNER JOIN " & m_sBatchTableName & " B ON B.BatchNumber = BS.BatchNumber " & _
        "INNER JOIN " & m_sBatchProgressTableName & " BSP ON BSP.BatchNumber = BS.BatchNumber AND BSP.BatchRunNumber = BS.BatchRunNumber " & _
        "INNER JOIN ComboValue ComboStatus ON ComboStatus.ValueID = BS.Status AND ComboStatus.GroupName = " & g_clsSQLAssistSP.FormatString(m_sScheduleStatusComboGroup) & _
        "INNER JOIN ComboValidation ComboValidStatus ON ComboValidStatus.ValueId = BS.Status AND ComboValidStatus.GroupName = " & g_clsSQLAssistSP.FormatString(m_sScheduleStatusComboGroup) & _
        "INNER JOIN ComboValue ComboProgram ON ComboProgram.ValueID = B.ProgramType AND ComboProgram.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & _
        "INNER JOIN ComboValidation ComboValidProgType ON ComboValidProgType.ValueId = B.ProgramType AND ComboValidProgType.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & _
        "WHERE BS.BatchNumber = " & g_clsSQLAssistSP.FormatString(sBatchNumber) & " " & _
        "ORDER BY BS.BatchRunNumber"
    'BM0438/BM0425 End
    
    'BM0425 The recordset is only required for reading
    'Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL, , , True)
    'BM0425 End
    
    m_clsTableAccess.SetRecordSet rs
    
    'BM0438 the returned recordset is never used
    'Set GetBatchScheduleTableData = rs
    'BM0438 End
    
Failed:
    Set rs = Nothing
    If Err.Number <> 0 Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If
End Function

'BM0425 Use a separate query for updates to minimizing joins and locking
Public Sub GetBatchScheduleForUpdate(ByVal sBatchNumber As String, ByVal sBatchRunNumber As String)

    Dim rs      As ADODB.Recordset
    Dim sSQL    As String
    
On Error GoTo Failed
    sSQL = "SELECT BatchNumber, BatchRunNumber, Status, CreationDateTime, ExecutionDateTime, " & _
        "CompletionDateTime, RetryBatchIndicator, FailedBatchRunNumber, ErrorNumber, ErrorSource, ErrorDescription " & _
        "FROM " & m_sTableName & " WITH (ROWLOCK)" & _
        " WHERE BatchNumber = " & g_clsSQLAssistSP.FormatString(sBatchNumber) & _
        " AND BatchRunNumber = " & g_clsSQLAssistSP.FormatString(sBatchRunNumber)
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL, , , False)
    m_clsTableAccess.SetRecordSet rs
    
Failed:
    Set rs = Nothing
    If Err.Number <> 0 Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If
End Sub
'BM0425 End

Public Function CancelAllPendingBatchRecords(Optional sBatchNumber As String) As Boolean
    On Error GoTo Failed
    Dim clsComboValidation As ComboValidationTable
    Dim nCurrentStatusID As Long
    Dim bRet As Boolean
    Dim nCnt As Long
    Dim nRunningStatusID As Long
    Dim nCancelledStatusID As Long
    Dim nWaitingStatusID As Long
    Dim rs As ADODB.Recordset
    Dim colValues As Collection
    
    Set clsComboValidation = New ComboValidationTable
    Set colValues = New Collection
    
    If Len(sBatchNumber) = 0 Then
        sBatchNumber = m_sBatchNumber
    End If
    
    'Get the ValueID for Schedule status of "RUNNING" and "WAITING"
    nRunningStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_RUNNING)
    colValues.Add nRunningStatusID
    nWaitingStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_WAITING)
    colValues.Add nWaitingStatusID
    
    Set rs = GetPendingBatchJobs(sBatchNumber, colValues)
    m_clsTableAccess.SetRecordSet rs
    
    'Are there any records returned?
    If m_clsTableAccess.RecordCount > 0 Then
        'Set the Status to Cancelled
        nCancelledStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_CANCELLED)
        For nCnt = 1 To m_clsTableAccess.RecordCount
            bRet = SetStatus(nCancelledStatusID)
            m_clsTableAccess.Update
            m_clsTableAccess.MoveNext
        Next
    Else
        bRet = True
    End If
    
    If Not bRet Then
        g_clsErrorHandling.RaiseError errCanNotUpdateBatchJob
    End If

    CancelAllPendingBatchRecords = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Function

Private Function GetPendingBatchJobs(sBatchNumber As String, colValues As Collection) As ADODB.Recordset

    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim nCnt As Long
    
    Set rs = New ADODB.Recordset
    
    sSQL = "SELECT Status FROM " & m_sTableName & " Where BatchNumber = " & g_clsSQLAssistSP.FormatString(sBatchNumber) & "  " & _
           "AND Status = "
    
    If Not colValues Is Nothing Then
        For nCnt = 1 To colValues.Count
            sSQL = sSQL & colValues(nCnt)
            
            If nCnt < colValues.Count Then
                sSQL = sSQL & " OR BatchNumber = " & g_clsSQLAssistSP.FormatString(sBatchNumber) & " AND Status = "
            End If
    
        Next nCnt
    End If
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    Set GetPendingBatchJobs = rs

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function CancelBatchRecord() As Boolean
    On Error GoTo Failed
    
    Dim clsComboValidation As ComboValidationTable
    Dim nCurrentStatusID As Long
    Dim bRet As Boolean
    Dim nRunningStatusID As Long
    Dim nCancelledStatusID As Long
    Dim nWaitingStatusID As Long
    Dim rs As ADODB.Recordset
    
    Set clsComboValidation = New ComboValidationTable
    
    
    nCurrentStatusID = GetBatchStatus
    'Get the ValueID for Schedule status of "RUNNING" and "WAITING"
    nRunningStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_RUNNING)

    nWaitingStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_WAITING)

    'Get the ValueID of Cancelled
    nCancelledStatusID = clsComboValidation.GetSingleValueFromValidation(m_sScheduleStatusComboGroup, BATCH_SCHEDULE_STATUS_CANCELLED)

    'Is the status Runnning or Waiting?
    If nCurrentStatusID = nRunningStatusID Or nCurrentStatusID = nWaitingStatusID Then
        bRet = True
        bRet = SetStatus(nCancelledStatusID)
        If bRet Then
            m_clsTableAccess.Update
        End If
    Else
        bRet = False
    End If
    
    If Not bRet Then
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot Cancel the Batch Job"
    End If

    CancelBatchRecord = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Function

Public Function GetNumberOfAuditErrors() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AUDITERRORSFOUND")
    GetNumberOfAuditErrors = CLng(vVal)
End Function

Public Function GetErrorNumber() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORNUMBER")
    If vVal <> "" Then
        GetErrorNumber = CLng(vVal)
    Else
        GetErrorNumber = 0
    End If
End Function

Public Function GetErrorSource() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORSOURCE")
    'BM0425 trim unnecessary white space
    GetErrorSource = Trim(CStr(vVal))
End Function

Public Function GetErrorDescription() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORDESCRIPTION")
    'BM0425 trim unnecessary white space
    GetErrorDescription = Trim(CStr(vVal))
End Function
