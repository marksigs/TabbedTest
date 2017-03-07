VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BatchTable
' Description   : Contains all data access functions for the BatchJobs Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DJP       22/06/01    SQL Server port, use global SQLAssist not local one
' DJP       22/06/01    SQL Server port, use Supervisor error object instead
' SA        03/01/02    SYS3327 new function GetNextBatchNumber
' SA        10/06/02    SYS4497 Ability to delete batch jobs at different stages added.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sBatchNumber As String
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName = "BATCH"
Private Const m_sProgramTypeComboGroup = "BatchProgramType"
Private Const m_sStatusComboGroup = "BatchStatus"
Private Const m_sFrequencyComboGroup = "BatchFrequency"
Private Const m_sScheduleStatusComboGroup = "BatchScheduleStatus"

Implements TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GHun      24/03/2003  BM0438 Batch problems - BATCHSCHEDULE table design
' BS          20/05/2003  BM0240 Amend GENERIC_SQL test in GetNextBatchNumber
' JD        30/11/2004  BMIDS604 refresh the list to update the status field on screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "BATCHNUMBER"
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

    colListLine.Add GetBatchNumber
    colListLine.Add GetProgramTypeText
    colListLine.Add GetUserID
    colListLine.Add GetStatusText
    colListLine.Add GetBatchFrequencyText
    colListLine.Add GetBatchDescription
    colListLine.Add GetExecutionDateTime
    colListLine.Add GetStatus, OBJECT_EXTRA_VALUE
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

Public Sub GetBatch()

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
    'SQL Server port, use global SQLAssist not local one
    sSQL = "SELECT " & m_sTableName & ".*, ComboStatus.ValueName StatusText, ComboProgramType.ValueName ProgramTypeText, ComboFrequency.ValueName FrequencyText" & _
           " FROM combovalue ComboStatus, combovalue ComboProgramType, combovalue ComboFrequency, " & m_sTableName & _
           " WHERE" & _
           " ComboStatus.GroupName = " & g_clsSQLAssistSP.FormatString(m_sStatusComboGroup) & _
           " AND comboProgramType.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & _
           " AND ComboFrequency.GroupName = " & g_clsSQLAssistSP.FormatString(m_sFrequencyComboGroup) & _
           " AND ComboStatus.ValueID = Status AND ComboFrequency.ValueID = Frequency AND ComboProgramType.ValueID = ProgramType"
           

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    ValidateRecordset rs, "BatchJobs"
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetBatchNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BATCHNUMBER")
    GetBatchNumber = CStr(vVal)
End Function

Public Function GetProgramTypeText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ProgramTypeText")
    GetProgramTypeText = CStr(vVal)
End Function

Public Function GetStatusText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("StatusText")
    GetStatusText = CStr(vVal)
End Function

Public Function GetStatus() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("Status")
    GetStatus = CLng(vVal)
End Function

Public Function GetUserID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("USERID")
    GetUserID = CStr(vVal)
End Function

Public Function GetBatchFrequency() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FREQUENCY")
    GetBatchFrequency = CLng(vVal)
End Function

Public Function GetBatchFrequencyText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FREQUENCYTEXT")
    GetBatchFrequencyText = CStr(vVal)
End Function

Public Function GetBatchType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BatchTypeText")
    GetBatchType = CStr(vVal)
End Function

Public Function GetBatchDescription() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("DESCRIPTION")
    GetBatchDescription = CStr(vVal)
End Function

Public Function GetBatchRetries() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BATCHRETRIES")
    GetBatchRetries = CStr(vVal)
End Function

Public Function GetBatchFailedNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FAILEDBATCHNUMBER")
    GetBatchFailedNumber = CStr(vVal)
End Function

Public Function GetBatchFailedRunNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FAILEDRUNNUMBER")
    GetBatchFailedRunNumber = CStr(vVal)
End Function

Public Function GetExecutionDateTime() As Date
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EXECUTIONDATETIME")
    GetExecutionDateTime = CDate(vVal)
End Function

Public Function GetProgramType() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PROGRAMTYPE")
    GetProgramType = CLng(vVal)
End Function

Friend Function SetProgramType(nProgType As Long) As Boolean
    SetProgramType = m_clsTableAccess.SetVal("PROGRAMTYPE", CVar(nProgType))
End Function

Friend Function SetBatchNumber(sBatchNo As String) As Boolean
    SetBatchNumber = m_clsTableAccess.SetVal("BATCHNUMBER", CVar(sBatchNo))
End Function

Friend Function SetUserID(sUserID As String) As Boolean
    SetUserID = m_clsTableAccess.SetVal("USERID", CVar(sUserID))
End Function

Friend Function SetStatus(nStatus As Long) As Boolean
    SetStatus = m_clsTableAccess.SetVal("STATUS", CVar(nStatus))
End Function

Friend Function SetFrequency(nFrequency As Long) As Boolean
    SetFrequency = m_clsTableAccess.SetVal("FREQUENCY", CVar(nFrequency))
End Function

Friend Function SetDescription(sDescription As String) As Boolean
    SetDescription = m_clsTableAccess.SetVal("DESCRIPTION", CVar(sDescription))
End Function

Friend Function SetExecutionDate(dExecutionDate As Date) As Boolean
    SetExecutionDate = m_clsTableAccess.SetVal("EXECUTIONDATETIME", CVar(dExecutionDate))
End Function

Public Sub GetProgramTypes(sGroup As String)

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    
    sSQL = "SELECT ValueName from ComboValue WHERE Groupname = " & _
           g_clsSQLAssistSP.FormatString(sGroup)
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    m_clsTableAccess.SetRecordSet rs
    
End Sub

Public Function GetProgramTypeAsCollection(colValues As Collection, colIDS As Collection) As Collection
    
    On Error GoTo Failed

    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim vTmp As Variant

    Set colValues = New Collection
    Set colIDS = New Collection
    
    
    sSQL = "SELECT ComboValue.ValueName, ComboValidation.ValidationType " & _
           "FROM ComboValue, ComboValidation " & _
           "WHERE ComboValue.GroupName = " & g_clsSQLAssistSP.FormatString(m_sProgramTypeComboGroup) & " " & _
           "AND ComboValue.GroupName = ComboValidation.GroupName (+) " & _
           "AND ComboValue.ValueID = ComboValidation.ValueID (+)"
           
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        While Not rs.EOF
            
            vTmp = rs.fields("ValueName")
            colValues.Add CStr(vTmp)
            
            vTmp = rs.fields("ValidationType")
            If Not IsNull(vTmp) Then
                colIDS.Add vTmp
            Else
                colIDS.Add ""
            End If
    
            rs.MoveNext
        Wend
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
        
End Function

Public Function GetComboField() As String
    Dim sField As String
    
    sField = "ValueName"
    GetComboField = sField
    
    
End Function

Public Function DeleteBatch(nStatus As Long) As Boolean
   
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim colMatchValues As Collection
    Dim sTargetDatabase As String
    Dim clsComboValidation As ComboValidationTable
    'Dim nValidationTypeID As Long      '--SYS4497
    Dim nLaunchedStatusID As Long       '++SYS4497
    Dim nInProgressStatusID As Long     '++SYS4497
    Dim sCommand As String              '++SYS4497
    Dim sBatchNo As String              '++SYS4497
    Dim clsTableAccess As TableAccess
    Dim bCreated As Boolean
    Dim vResponse As Variant
    
    Set clsComboValidation = New ComboValidationTable
    
    'SYS4497 Can now delete batches at all stages except "Launched" and "In Progress"
    'Get the validationTypeID for Batch status "CREATED"
    'nValidationTypeID = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_CREATED, m_sStatusComboGroup)
    nLaunchedStatusID = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_LAUNCHED, m_sStatusComboGroup)
    nInProgressStatusID = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_INPROGRESS, m_sStatusComboGroup)
    
    bRet = True
    'If nStatus = nValidationTypeID Then    'SYS4497
    If nStatus <> nLaunchedStatusID And nStatus <> nInProgressStatusID Then
        bCreated = False
    Else
        bCreated = True
    End If
    
    If Not bCreated Then
        
        vResponse = g_clsFormProcessing.DisplayMessage("Are you sure you want to delete this Batch?" & vbCrLf & vbCrLf & "You will no longer be able to view or re-run failed batch jobs if you do.")
        
        If vResponse = vbYes Then
                        
            Set clsTableAccess = New PaymentProcessingTable
            
            sTargetDatabase = m_clsTableAccess.GetDatabase()
            Set colMatchValues = m_clsTableAccess.GetKeyMatchValues()
                    
            'SYS4497 If status is Created, original code applies:
            If nStatus = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_CREATED, m_sStatusComboGroup) Then
                
                'First Parameters
                clsTableAccess.SetKeyMatchValues colMatchValues
                clsTableAccess.DeleteRecords sTargetDatabase
            
                'Now the Batch Record
                m_clsTableAccess.DeleteRecords sTargetDatabase
            Else
                'SYS4497 {
                'We have several tables to delete from - need to go from bottom up
                'First work out type of batch being deleted to minimise calls to database
                sBatchNo = colMatchValues.Item(1)
                Select Case UCase$(Left$(sBatchNo, 1))
                    Case "P"        'Payment Processing
                        'Delete all PayProcBatchAudit records first
                        sCommand = "DELETE FROM PAYPROCBATCHAUDIT WHERE BATCHNUMBER = '" & sBatchNo & "'"
                        g_clsDataAccess.ExecuteCommand (sCommand)
                    Case "R"        'Rate Change
                        'Delete all RateChangeBatchAudit records first
                        sCommand = "DELETE FROM RATECHANGEBATCHAUDIT WHERE BATCHNUMBER = '" & sBatchNo & "'"
                        g_clsDataAccess.ExecuteCommand (sCommand)
                
                    Case "V"        'Valuer Payments
                       'Delete all ValuationBatchAudit records first
                        sCommand = "DELETE FROM VALUATIONBATCHAUDIT WHERE BATCHNUMBER = '" & sBatchNo & "'"
                        g_clsDataAccess.ExecuteCommand (sCommand)

                End Select
                'Now delete all BatchAudit records
                sCommand = "DELETE FROM BATCHAUDIT WHERE BATCHNUMBER = '" & sBatchNo & "'"
                g_clsDataAccess.ExecuteCommand (sCommand)
                
                'BM0438 Now delete the BatchScheduleProgress records
                sCommand = "DELETE FROM BATCHSCHEDULEPROGRESS WHERE BATCHNUMBER = '" & sBatchNo & "'"
                g_clsDataAccess.ExecuteCommand (sCommand)
                'BM0438 End
                
                'Now the BatchSchedule records
                sCommand = "DELETE FROM BATCHSCHEDULE WHERE BATCHNUMBER = '" & sBatchNo & "'"
                g_clsDataAccess.ExecuteCommand (sCommand)
                
                'Now the Parameters if it's a payment processing batch
                If UCase$(Left$(sBatchNo, 1)) = "P" Then
                    clsTableAccess.SetKeyMatchValues colMatchValues
                    clsTableAccess.DeleteRecords sTargetDatabase
                End If
                
                'Finally the Batch record itself
                m_clsTableAccess.DeleteRecords sTargetDatabase
                
            End If
            'SYS4497}
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot delete batch job which has been launched or is in progress "
    End If
    
    DeleteBatch = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function

Public Function UpdateBatchStatus(nStatus As Long) As Boolean

    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nSetStatus As Long
    Dim clsComboValidation As ComboValidationTable
    Dim nValidationCheck As Long
    Dim nCurStatus As Long
    Dim clsBatchSchedule As BatchScheduleTable
    Dim colKeys As Collection
    Dim sFunction As String
    Dim clsBatchProcessing As BatchProcessing 'SYS3327
    Dim dExecutionDateTime As Date 'SYS3327

    Set clsComboValidation = New ComboValidationTable
    Set clsBatchSchedule = New BatchScheduleTable
    Set clsBatchProcessing = New BatchProcessing 'SYS3327
    
    bRet = False
    If Len(m_sBatchNumber) = 0 Then
        Set colKeys = m_clsTableAccess.GetKeyMatchValues
        m_sBatchNumber = colKeys(1)
    End If
    
    'Get current Batch Record
    m_clsTableAccess.SetRecordSet m_clsTableAccess.GetTableData()
        
    Select Case nStatus
        Case LV_BATCH_PROCESS_LAUNCH
            nValidationCheck = clsComboValidation.GetSingleValueFromValidation(m_sStatusComboGroup, BATCH_STATUS_CREATED)
            nSetStatus = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_LAUNCHED, m_sStatusComboGroup)
            sFunction = "Launch"
        Case LV_BATCH_PROCESS_RESTART
            nValidationCheck = clsComboValidation.GetSingleValueFromValidation(m_sStatusComboGroup, BATCH_STATUS_CANCELLED)
            nSetStatus = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_LAUNCHED, m_sStatusComboGroup)
            sFunction = "Restart"
        Case LV_BATCH_PROCESS_CANCEL
            sFunction = "Cancel"
            nValidationCheck = clsComboValidation.GetSingleValueFromValidation(m_sStatusComboGroup, BATCH_STATUS_LAUNCHED)
            nSetStatus = clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_CANCELLED, m_sStatusComboGroup)
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, "Batch - Invalid Status"
    End Select
    
    nCurStatus = GetStatus
    'Do The validation - Is the currentstatus = to the value of the validationType of the status it needs to be????
    If nCurStatus = nValidationCheck Then
        bRet = SetStatus(nSetStatus)
        If nStatus = LV_BATCH_PROCESS_CANCEL Then
            clsBatchSchedule.CancelAllPendingBatchRecords m_sBatchNumber
        End If
        'SYS3327 Launching should actually Launch!!
        If nStatus = LV_BATCH_PROCESS_LAUNCH Then
            dExecutionDateTime = GetExecutionDateTime()     'SA SYS3327 Need Execution Date/Time if launching
            clsBatchProcessing.Launch m_sBatchNumber, dExecutionDateTime
        End If
    Else
        g_clsMainSupport.GetBatch frmMain.lvListView 'JD BMIDS604 refresh the list
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot " & sFunction & " the Batch Job."
    End If
    'SYS3327 Updating of table is handled by middle tier.
    If bRet And nStatus <> LV_BATCH_PROCESS_LAUNCH And nStatus <> LV_BATCH_PROCESS_CANCEL Then
        m_clsTableAccess.Update
    End If
    
    UpdateBatchStatus = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

Public Function GetNextBatchNumber(sPrefix As String) As String
    On Error GoTo Failed
    Dim clsComboValidation As New ComboValidationTable
    Dim sSQL As String
    Dim rs As Recordset
    Dim sNextBatchNumber As String
    
    'Depending on the prefix, get the next number in the sequence
    'The batch table will not be a large table - hundreds rather than thousands of records - if that!
    'Other options for generating a unique number include holding counters as global parameters or using the
    'SeqNextNumber table. Both these required more changes than this way to code
    'Time, as usual, was at a premium!!
    
    'BS BM0240 20/05/03
    '#If GENERIC_SQL = 1 Then
    #If GENERIC_SQL Then
        'SQL Server
        sSQL = "SELECT MAX(BatchNumber) as BatchNumber FROM Batch WHERE SUBSTRING(BatchNumber, 1, 1)= '" & sPrefix & "'"
    #Else
        'Oracle
        sSQL = "SELECT MAX(BatchNumber) as BatchNumber FROM Batch WHERE SUBSTR(BatchNumber, 1, 1)= ' " & sPrefix & "'"
    #End If
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    If rs.RecordCount > 0 Then  ' there will only be one record
        rs.MoveFirst
        If Not IsNull(rs.fields("BatchNumber")) Then
            sNextBatchNumber = rs.fields("BatchNumber")
            sNextBatchNumber = Right(sNextBatchNumber, 4)
            sNextBatchNumber = sNextBatchNumber + 1
            sNextBatchNumber = sPrefix & Format$(sNextBatchNumber, "0000")
        Else
            'They have got a clear batch table so start from 1
            sNextBatchNumber = sPrefix & "0001"
        End If
    End If
    
    GetNextBatchNumber = sNextBatchNumber
        
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


