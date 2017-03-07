VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransferOfEquityFeeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TrnsferOfEquityFeeTable
' Description   : Contains all access routines and SQL for the TransferOfEquityFee table, which is a
'                 banded table. This class implements both TableAccess for generic routines,
'                 and BandedTable for Banded specific functionality.
' Change history
' Prog      Date        Description
' TW        11/12/2006  EP2_20 Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
Implements BandedTable

' Constants
Private Const m_sTableName = "TRANSFEROFEQUITYFEE"
Private Const m_sVersionField As String = "AFSETVERSIONNUMBER"

' Private data
Private m_sSearch As String
Private m_sFirstMatch As String
Private m_bVersionExists As Boolean
Private m_clsBandedTable As BandedTable
Private m_clsTableAccess As TableAccess
Private m_sTransferOfEquityFeeStartDateField As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Sub SetFindNewestTransferOfEquityFeeSet()
    Dim sSysDate As String
    
    sSysDate = g_clsSQLAssistSP.GetSystemDate
    m_sFirstMatch = "SELECT " & _
                        "CF1.* " & _
                    "FROM " & _
                        "TransferOfEquityFee CF1 " & _
                    "WHERE " & _
                        "STARTDATE IN " & _
                        "(SELECT " & _
                            "MAX(STARTDATE) " & _
                        "FROM " & _
                            "TransferOfEquityFee CF2 " & _
                        "WHERE " & _
                            "CF1.TransferOfEquityFeeSET = CF2.TransferOfEquityFeeSET AND " & _
                            g_clsSQLAssistSP.TruncateDateColumn(sSysDate) & " >= " & g_clsSQLAssistSP.TruncateDateColumn("STARTDATE") & ") AND "
    m_sFirstMatch = m_sFirstMatch & "CF1.NATUREOFLOAN = " & _
                            "(SELECT " & _
                                "MIN(NATUREOFLOAN) " & _
                            "FROM " & _
                                "TransferOfEquityFee CF2 " & _
                            "WHERE " & _
                                "CF1.TransferOfEquityFeeSET = CF2.TransferOfEquityFeeSET AND " & _
                                "CF1.STARTDATE = CF2.STARTDATE)"
    
    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub
Friend Function GetFeeSet() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TransferOfEquityFeeSet")
    GetFeeSet = CStr(vVal)
End Function

Friend Function GetStartDate() As String
    Dim vVal As Variant

    vVal = m_clsTableAccess.GetVal(m_sTransferOfEquityFeeStartDateField)
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetAmount() As String
    GetAmount = CStr(m_clsTableAccess.GetVal("Amount"))
End Function

Friend Function GetNatureOfLoan() As String
    GetNatureOfLoan = CStr(m_clsTableAccess.GetVal("NatureOfLoan"))
End Function

Friend Function GetNatureOfLoanDescription() As String
    GetNatureOfLoanDescription = CStr(m_clsTableAccess.GetVal("NatureOfLoanDescription"))
End Function
Friend Function SetFeeSet(sFeeSet As String) As Boolean
    SetFeeSet = m_clsTableAccess.SetVal("TransferOfEquityFeeSet", CVar(sFeeSet))
    
    g_clsVersion.HandleVersioning Me
End Function

Friend Function SetNatureOfLoan(sNatureOfLoan As String) As Boolean
    SetNatureOfLoan = m_clsTableAccess.SetVal("NatureOfLoan", CVar(sNatureOfLoan))
End Function

Friend Function SetNatureOfLoanDescription(sNatureOfLoanDescription As String) As Boolean
    SetNatureOfLoanDescription = m_clsTableAccess.SetVal("NatureOfLoanDescription", CVar(sNatureOfLoanDescription))
End Function
Friend Function SetStartDate(sStartDate As String) As Boolean
    SetStartDate = m_clsTableAccess.SetVal(m_sTransferOfEquityFeeStartDateField, CVar(sStartDate))
End Function
Friend Function SetAmount(sAmount As String) As Boolean
    SetAmount = m_clsTableAccess.SetVal("Amount", CVar(sAmount))
End Function
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
Public Function GetStartDateField() As String
    GetStartDateField = m_sTransferOfEquityFeeStartDateField
End Function
Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    Dim bVersionExists As Boolean
    
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    
    m_sTransferOfEquityFeeStartDateField = "STARTDATE"
    
    m_sFirstMatch = "select * from TransferOfEquityFee A where A.NATUREOFLOAN = " + _
    "(select min(NATUREOFLOAN) from TransferOfEquityFee C " + _
        "where A.TransferOfEquityFeeSet = C.TransferOfEquityFeeSet and A." + m_sTransferOfEquityFeeStartDateField + " = C." + m_sTransferOfEquityFeeStartDateField + _
        ")"

    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " WHERE " + m_sVersionField + " = " & g_sVersionNumber
    End If
        
        
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
        
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch

    colFields.Add "TransferOfEquityFeeSet"
    colFields.Add m_sTransferOfEquityFeeStartDateField '"StartDate"
    
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "TransferOfEquityFeeSet"
    colUpdateFields.Add m_sTransferOfEquityFeeStartDateField '"StartDate"

    m_clsBandedTable.SetUpdateFields colUpdateFields

    Set colDupKeys = New Collection
    colDupKeys.Add "NATUREOFLOAN"

    m_clsTableAccess.SetDuplicateKeys colDupKeys
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetFeeSet()
    colHeader.Add GetStartDate()
End Sub

Public Sub SetPopulateNull()
    m_sSearch = "Select * from " + m_sTableName + " where Name is null"
End Sub

Public Sub SetPopulateType(enumType As PopulateType)
    m_clsTableAccess.SetPopulateType enumType
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

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub

Private Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
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

Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
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

Private Function TableAccess_GetUpdateValues() As Collection
    Set TableAccess_GetUpdateValues = m_clsTableAccess.GetUpdateValues
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
    m_clsTableAccess.SetPopulateType enumType
End Sub

Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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

Public Function TableAccess_GetUpdateFields() As Collection
    Dim colKeys As New Collection
    
    colKeys.Add "TransferOfEquityFeeSet"
    colKeys.Add m_sTransferOfEquityFeeStartDateField '"StartDate"

    Set TableAccess_GetUpdateFields = colKeys
End Function

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetNatureOfLoanDescription()
    colListLine.Add GetAmount()

End Sub

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")

    On Error GoTo Failed
    Dim sMessage As String
    Dim inum As Integer
    Dim colValues  As New Collection
    Set colValues = m_clsTableAccess.GetKeyMatchValues
    
    Dim clsMortgageProductTable As MortgageProductTable
    Set clsMortgageProductTable = New MortgageProductTable
    
    Set colValues = clsMortgageProductTable.DoesProductExistForThisItem("TransferOfEquityFeeSet", colValues.Item(1), sDatabase)
    
    If (Not colValues Is Nothing) And (colValues.Count > 0) Then
        sMessage = ""
        For inum = 1 To colValues.Count
            sMessage = sMessage & "(" & inum & ") " & colValues.Item(inum) & vbCrLf
        Next
        
        MsgBox "Record can not be deleted. The following Mortgage Products exist for the selected Credit Limit Increase Fee Set : " & vbCrLf & sMessage, vbCritical, "Cannot Delete"
        g_clsErrorHandling.RaiseError , "Record can not be deleted. The following Mortgage Products exist for the selected Credit Limit Increase Fee Set : " & vbCrLf & sMessage, ErrorSystem
    Else
        m_clsTableAccess.DeleteRecords sDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteTransferOfEquityFees()
    On Error GoTo Failed
    ' First, delete the Transfer of Equity fee values
    m_clsTableAccess.DeleteRecords
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



