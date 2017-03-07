VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaymentProcessingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PaymentProcessingTable
' Description   : Contains all data access functions for the PaymentProcessingBatchParameters Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_bUpdated As Boolean
Private Const m_sTableName = "PayProcBatchParams"
Private Const m_sComboField = "ValueName"
Private Const m_sBatchTable = "Batch"

Implements TableAccess

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


Public Function GetBatchNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BATCHNUMBER")
    GetBatchNumber = CStr(vVal)
End Function

Public Function GetPaymentJobType() As Long
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PAYMENTJOBTYPE")
    GetPaymentJobType = CLng(vVal)
End Function
Public Function GetPaymentJobTypeText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PAYMENTJOBTYPETEXT")
    GetPaymentJobTypeText = CStr(vVal)
End Function

Public Function GetBacs() As Integer
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BACS")
    GetBacs = CInt(vVal)
End Function

Public Function GetChaps() As Integer
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CHAPSORTT")
    GetChaps = CInt(vVal)
End Function
Public Function GetCheque() As Integer
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CHEQUE")
    GetCheque = CInt(vVal)
End Function

Public Function GetInternalTransfer() As Integer
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("INTERNALTRANSFER")
    GetInternalTransfer = CInt(vVal)
End Function

Public Function GetIssueDate() As Date
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ISSUEDATE")
    GetIssueDate = CDate(vVal)
End Function

Public Function GetApplicationNoTo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("APPNUMRANGEEND")
    GetApplicationNoTo = CStr(vVal)
End Function

Public Function GetApplicationNoFrom() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("APPNUMRANGESTART")
    GetApplicationNoFrom = CStr(vVal)
End Function

Friend Function SetBatchNumber(sBatchNo As String) As Boolean
    SetBatchNumber = m_clsTableAccess.SetVal("BATCHNUMBER", CVar(sBatchNo))
End Function
Friend Function SetPaymentJobType(nPaymentType As Long) As Boolean
    SetPaymentJobType = m_clsTableAccess.SetVal("PAYMENTJOBTYPE", CVar(nPaymentType))
End Function

Friend Function SetBacs(nBacs As Integer) As Boolean
    SetBacs = m_clsTableAccess.SetVal("BACS", CVar(nBacs))
End Function

Friend Function SetChaps(nChaps As Integer) As Boolean
    SetChaps = m_clsTableAccess.SetVal("CHAPSORTT", CVar(nChaps))
End Function

Friend Function SetCheque(nCheque As Integer) As Boolean
    SetCheque = m_clsTableAccess.SetVal("CHEQUE", CVar(nCheque))
End Function

Friend Function SetInternalTransfer(nInternalTransfer As Integer) As Boolean
    SetInternalTransfer = m_clsTableAccess.SetVal("INTERNALTRANSFER", CVar(nInternalTransfer))
End Function

Friend Function SetIssueDate(dIssueDate As Date) As Boolean
    SetIssueDate = m_clsTableAccess.SetVal("ISSUEDATE", CVar(dIssueDate))
End Function
Friend Function SetApplicationNumberTo(sAppNoTo As String) As Boolean
    SetApplicationNumberTo = m_clsTableAccess.SetVal("APPNUMRANGEEND", CVar(sAppNoTo))
End Function
Friend Function SetApplicationNumberFrom(sAppNoFrom As String) As Boolean
    SetApplicationNumberFrom = m_clsTableAccess.SetVal("APPNUMRANGESTART", CVar(sAppNoFrom))
End Function

Public Function GetComboField() As String
    
    GetComboField = m_sComboField
    
End Function

Public Sub SetUpdated(bUpdated As Boolean)
    m_bUpdated = bUpdated
End Sub

Public Function GetIsUpdated() As Boolean
    GetIsUpdated = m_bUpdated
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub


