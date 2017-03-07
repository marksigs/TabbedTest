VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LedgerCodesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LedgerCodesTable
' Description   : Class which handles the table details of Ledger records on form Lender Details
' Change history
' Prog      Date        Description
' CL        16/05/01    Added Class
' DJP       24/10/01    Added combos and ledger codes (previously in LedgerCodes.cls)
' STB       07/11/01    SYS2906: SQL altered to return unique records only.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Private Const m_sTableName = "LEDGERCODES"
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Implements BandedTable
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    Dim colUpdateFields As Collection
    Set m_clsBandedTable = New BandedTable
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "ORGANISATIONID"
    m_clsTableAccess.SetKeyMatchFields colFields
        
    ' Banded table initialisation
    Set colUpdateFields = New Collection
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    Set colUpdateFields = New Collection
    
    colUpdateFields.Add "ORGANISATIONID"

    m_clsBandedTable.SetUpdateFields colUpdateFields
  
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
    'TableAccess_GetSearch = m_sSearch
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
Public Function GetOrganisationID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ORGANISATIONID")
    ' DJP SQL Server port - return true type, not string
    GetOrganisationID = vVal
End Function
Public Function GetLedgerCodeType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("LEDGERCODETYPE")
    GetLedgerCodeType = CStr(vVal)
End Function
Public Function GetValueID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("VALUEID")
    GetValueID = CStr(vVal)
End Function
Public Function GetValueNameText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("VALUENAMETEXT")
    GetValueNameText = CStr(vVal)
End Function
Public Function GetLedgerCode1() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("LEDGERCODE1")
    GetLedgerCode1 = CStr(vVal)
End Function
Public Function GetLedgerCode2() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("LEDGERCODE2")
    GetLedgerCode2 = CStr(vVal)
End Function
' DJP SQL Server port
Friend Function SetOrganisationID(vName As Variant) As Boolean
    SetOrganisationID = m_clsTableAccess.SetVal("ORGANISATIONID", vName)
End Function
Friend Function SetLedgerCodeType(sName As String) As Boolean
    SetLedgerCodeType = m_clsTableAccess.SetVal("LEDGERCODETYPE", CVar(sName))
End Function
Friend Function SetValueID(sName As String) As Boolean
    SetValueID = m_clsTableAccess.SetVal("VALUEID", CVar(sName))
End Function
Friend Function SetValueNameText(sName As String) As Boolean
    SetValueNameText = m_clsTableAccess.SetVal("VALUENAMETEXT", CVar(sName))
End Function
Friend Function SetLedgerCode1(sName As String) As Boolean
    SetLedgerCode1 = m_clsTableAccess.SetVal("LEDGERCODE1", CVar(sName))
End Function
Friend Function SetLedgerCode2(sName As String) As Boolean
    SetLedgerCode2 = m_clsTableAccess.SetVal("LEDGERCODE2", CVar(sName))
End Function
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
Public Function GetComboRecordset() As ADODB.Recordset
On Error GoTo Failed

    Dim strSQL As String
    Dim rs As ADODB.Recordset
                   
    'SB: SYS2906 - Return unique combo named items - this was returning duplicates.
    strSQL = "SELECT COMBOVALIDATION.GROUPNAME, COMBOVALUE.VALUEID, COMBOVALUE.VALUENAME From COMBOVALUE, COMBOVALIDATION Where COMBOVALUE.VALUEID = COMBOVALIDATION.VALUEID AND COMBOVALUE.GROUPNAME = COMBOVALIDATION.GROUPNAME AND ((COMBOVALUE.GROUPNAME = 'OneOffCost') AND (COMBOVALIDATION.VALIDATIONTYPE <> 'TID') OR COMBOVALUE.GROUPNAME = 'PPCoverType' OR COMBOVALUE.GROUPNAME = 'BCCoverType') GROUP BY COMBOVALIDATION.GROUPNAME, COMBOVALUE.VALUEID, COMBOVALUE.VALUENAME ORDER BY COMBOVALUE.VALUENAME"
    'END SYS2906.
    
    Set rs = g_clsDataAccess.GetTableData("COMBOVALUE", strSQL)
    Set GetComboRecordset = rs
    
Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetLedgerCodes() As ADODB.Recordset
On Error GoTo Failed
    
    Dim strSQL As String
    Dim rs As ADODB.Recordset
                    
    strSQL = "SELECT " & _
    "COMBOVALIDATION.VALIDATIONTYPE, " & _
    "COMBOVALIDATION.GROUPNAME, " & _
    "COMBOVALUE.VALUEID, " & _
    "COMBOVALUE.VALUENAME " & _
"FROM " & _
    "COMBOVALUE, " & _
    "COMBOVALIDATION " & _
"WHERE " & _
    "COMBOVALUE.VALUEID = COMBOVALIDATION.VALUEID AND " & _
    "COMBOVALIDATION.GROUPNAME = COMBOVALUE.GROUPNAME AND (COMBOVALIDATION.VALIDATIONTYPE <> 'TID') AND " & _
    "(COMBOVALUE.GROUPNAME = 'OneOffCost') AND (NOT EXISTS " & _
        "(SELECT * " & _
         "FROM " & _
            "LEDGERCODES " & _
          "WHERE " & _
            "LEDGERCODETYPE = COMBOVALUE.GROUPNAME AND LEDGERCODES.VALUEID = COMBOVALUE.VALUEID)) OR " & _
            "(COMBOVALUE.GROUPNAME = 'PPCoverType') AND (NOT EXISTS " & _
                          "(SELECT     * " & _
                            "FROM LEDGERCODES " & _
                            "WHERE " & _
                                "LEDGERCODETYPE = COMBOVALUE.GROUPNAME AND LEDGERCODES.VALUEID = COMBOVALUE.VALUEID)) OR " & _
                                "(COMBOVALUE.GROUPNAME = 'BCCoverType') AND (NOT EXISTS "
strSQL = strSQL & "(SELECT     * " & _
                            "FROM " & _
                                "LEDGERCODES " & _
                            "WHERE " & _
                                "LEDGERCODETYPE = COMBOVALUE.GROUPNAME AND LEDGERCODES.VALUEID = COMBOVALUE.VALUEID))"
    
    Set rs = g_clsDataAccess.GetTableData("COMBOVALUE", strSQL)
    Set GetLedgerCodes = rs
    
Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
