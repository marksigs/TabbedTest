VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BatchAuditTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BatchAuditTable
' Description   : Contains all data access functions for the BatchAudit Table
' Change history
' Prog      Date        Description
' SA        12/02/02    Added class SYS4073
' SA        01/03/02    SYS4209 GetBatchAuditTableData - Add in Alias's
' GHun      15/03/02    SYS4279 Supervisor error raised on viewing batch schedule.
' GHun      18/04/02    SYS4424 In Supervisor you cant view errors in incomplete batch runs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_sBatchNumber As String
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_sBatchType As String
Private Const m_sTableName = "BATCHAUDIT"

Implements TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GHun      24/03/2003  BM0425 Don't lock records when reading from the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    On Error GoTo Failed

    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess

    m_clsTableAccess.SetKeyMatchFields colFields
    m_clsTableAccess.Create m_sTableName

    colFields.Add "BATCHNUMBER"
    colFields.Add "BATCHRUNNUMBER"
    
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
    
    Select Case m_sBatchType
        
        Case "P"
            'Payment Processing
            colListLine.Add GetApplicationNo
            colListLine.Add GetPaySeqNo
        
        Case "V"
            'Valuation
            colListLine.Add GetApplicationNo
            colListLine.Add GetApplicationFFNo
            colListLine.Add GetInstructionSeqNo
            colListLine.Add GetValuerInvoiceAmount
        
        Case "R"
            'Rate Change
            colListLine.Add GetApplicationNo
            colListLine.Add GetApplicationFFNo
            colListLine.Add GetQuoteNo
            colListLine.Add GetMortSubQuoteNo
            colListLine.Add GetLoanCompSeqNo
        
    End Select
    
    ' All three batch types have error number, error source and error description
    colListLine.Add GetErrorNo
    colListLine.Add GetErrorSource
    colListLine.Add GetErrorDesc
    
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

'BM0425 the returned recordset is never used. OOSS does allow a function to be changed to
'       a Sub so the unused return value is set to the smallest data type available (byte)
'Public Function GetBatchAuditTableData(sBatchNumber As String, sBatchRunNumber As String) As ADODB.Recordset
Public Function GetBatchAuditTableData(ByVal sBatchNumber As String, ByVal sBatchRunNumber As String) As Byte
'BM0425 End

    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    'SYS4209 Add in Alias for Error fields
    Select Case m_sBatchType
        Case "P"
            'Payment Processing
            ' BM0425 Use NOLOCK locking hint
            ' SYS4424 Changed the query to do a left join on PayProcBatchAudit, rather than inner join
            ' otherwise the query does not return batchaudit records which are missing PayProcBatchAudit records
            sSQL = "SELECT  BS.BatchNumber, BS.BatchRunNumber,ApplicationNumber," & _
                "PaymentSequenceNumber,B.ErrorNumber,B.ErrorDescription,B.ErrorSource " & _
                "FROM BatchAudit b WITH (NOLOCK) " & _
                "LEFT JOIN " & _
                "BatchSchedule BS " & _
                "ON B.BatchNumber = BS.BatchNumber " & _
                "AND B.BatchRunNumber = BS.BatchRunNumber " & _
                "LEFT JOIN " & _
                "PayProcBatchAudit P " & _
                "ON  B.BatchNumber = P.BatchNumber " & _
                "AND B.BatchRunNumber = P.BatchRunNumber " & _
                "AND B.BatchAuditGuid = P.BatchAuditGuid " & _
                "Where B.BatchNumber = '" & sBatchNumber & "' " & _
                "AND BS.BatchRunNumber = " & sBatchRunNumber & _
                " AND B.Status = (SELECT ValueId FROM ComboValidation " & _
                " WHERE ComboValidation.GroupName = 'BatchAuditStatus' " & _
                " AND ComboValidation.ValidationType = 'F')"
        
        Case "R"
            ' BM0425 Use NOLOCK locking hint
            ' SYS4279 Added a space between B.ErrorSource and FROM
            ' SYS4424 Changed the query to do a left join on RateChangeBatchAudit, rather than inner join
            ' otherwise the query does not return batchaudit records which are missing RateChangeBatchAudit records
            sSQL = "SELECT  BS.BatchNumber, BS.BatchRunNumber, ApplicationNumber, " & _
                   "ApplicationFactFindNumber, QuotationNumber, MortgageSubQuoteNumber, " & _
                   "LoanComponentSequenceNumber, B.ErrorNumber, B.ErrorDescription, B.ErrorSource " & _
                   "FROM BatchAudit b WITH (NOLOCK) " & _
                   "LEFT JOIN BatchSchedule BS " & _
                   "ON B.BatchNumber = BS.BatchNumber " & _
                   "AND B.BatchRunNumber = BS.BatchRunNumber " & _
                   "LEFT JOIN RateChangeBatchAudit R " & _
                   "ON B.BatchNumber = R.BatchNumber " & _
                   "AND B.BatchRunNumber = R.BatchRunNumber " & _
                   "AND B.BatchAuditGuid = R.BatchAuditGuid " & _
                   "WHERE B.BatchNumber = '" & sBatchNumber & "' " & _
                   "AND BS.BatchRunNumber = " & sBatchRunNumber & _
                   " AND B.Status = (SELECT ValueId FROM ComboValidation " & _
                   " WHERE ComboValidation.GroupName = 'BatchAuditStatus' " & _
                    " AND ComboValidation.ValidationType = 'F')"
        Case "V"
            'Valuation
            ' BM0425 Use NOLOCK locking hint
            'SYS4424 Changed the query to do a left join on ValuationBatchAudit, rather than inner join
            ' otherwise the query does not return batchaudit records which are missing ValuationBatchAudit records
            sSQL = "SELECT  BS.BatchNumber, BS.BatchRunNumber, ApplicationNumber, " & _
                   "ApplicationFactFindNumber, InstructionSequenceNo, ValuerInvoiceAmount, " & _
                   "B.ErrorNumber, B.ErrorDescription, B.ErrorSource " & _
                   "FROM BatchAudit B WITH (NOLOCK) " & _
                   "LEFT JOIN BatchSchedule BS " & _
                   "ON B.BatchNumber = BS.BatchNumber " & _
                   "AND B.BatchRunNumber = BS.BatchRunNumber " & _
                   "LEFT JOIN ValuationBatchAudit V " & _
                   "ON B.BatchNumber = V.BatchNumber " & _
                   "AND B.BatchRunNumber = V.BatchRunNumber " & _
                   "AND B.BatchAuditGuid = V.BatchAuditGuid " & _
                   "WHERE B.BatchNumber = '" & sBatchNumber & "' " & _
                   "AND BS.BatchRunNumber = " & sBatchRunNumber & _
                   " AND B.Status = (SELECT ValueId FROM ComboValidation " & _
                   " WHERE ComboValidation.GroupName = 'BatchAuditStatus' " & _
                   " AND ComboValidation.ValidationType = 'F')"
    End Select
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL, , , True)
    m_clsTableAccess.SetRecordSet rs

    'BM0425 returned recordset is never used
    'Set GetBatchAuditTableData = rs
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetBatchType(sBatchType As String)
    m_sBatchType = sBatchType
End Sub

Public Function GetApplicationNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("APPLICATIONNUMBER")
    GetApplicationNo = CStr(vVal)
End Function

Public Function GetPaySeqNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PAYMENTSEQUENCENUMBER")
    GetPaySeqNo = CStr(vVal)
End Function

Public Function GetApplicationFFNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("APPLICATIONFACTFINDNUMBER")
    GetApplicationFFNo = CStr(vVal)
End Function

Public Function GetQuoteNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("QUOTATIONNUMBER")
    GetQuoteNo = CStr(vVal)
End Function

Public Function GetMortSubQuoteNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MORTGAGESUBQUOTENUMBER")
    GetMortSubQuoteNo = CStr(vVal)
End Function

Public Function GetLoanCompSeqNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("LOANCOMPONENTSEQUENCENUMBER")
    GetLoanCompSeqNo = CStr(vVal)
End Function

Public Function GetInstructionSeqNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("INSTRUCTIONSEQUENCENO")
    GetInstructionSeqNo = CStr(vVal)
End Function

Public Function GetValuerInvoiceAmount() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("VALUERINVOICEAMOUNT")
    GetValuerInvoiceAmount = CStr(vVal)
End Function

Public Function GetErrorNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORNUMBER")
    GetErrorNo = CStr(vVal)
End Function

Public Function GetErrorSource() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORSOURCE")
    'BM0425 trim unnecessary white space
    GetErrorSource = Trim(CStr(vVal))
End Function

Public Function GetErrorDesc() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ERRORDESCRIPTION")
    'BM0425 trim unnecessary white space
    GetErrorDesc = Trim(CStr(vVal))
End Function

