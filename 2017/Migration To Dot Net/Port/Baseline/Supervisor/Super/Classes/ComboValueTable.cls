VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboValueTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ComboValueTable
' Description   : Contains all access routines and sql for the combovalue table.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' SA        15/02/02    SYS4071 Method added to populate Payment Job Type Combo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Implements BandedTable

' Constants
Private Const m_sTableName = "COMBOVALUE"

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    
    m_clsTableAccess.Create m_sTableName

    colFields.Add "GroupName"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    colUpdateFields.Add "GroupName"
    
    m_clsBandedTable.SetUpdateFields colUpdateFields
    m_clsBandedTable.SetTableAccess m_clsTableAccess

    m_clsTableAccess.SetOrderCriteria "VALUEID"
    
    Set colDupKeys = New Collection
    colDupKeys.Add "GROUPNAME"
    colDupKeys.Add "VALUEID"
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetTableName() As String
    GetTableName = m_sTableName
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetGroupName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("GroupName")
    GetGroupName = CStr(vVal)
End Function
Friend Function GetValueID() As String
    GetValueID = CStr(m_clsTableAccess.GetVal("ValueID"))
End Function
Friend Function GetValueName() As String
    GetValueName = CStr(m_clsTableAccess.GetVal("ValueName"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetGroupName(sGroupName As String) As Boolean
    SetGroupName = m_clsTableAccess.SetVal("GroupName", CVar(sGroupName))
End Function
Friend Function SetValueID(sValueId As String) As Boolean
    SetValueID = m_clsTableAccess.SetVal("ValueID", CVar(sValueId))
End Function
Friend Function SetValueName(sValueName As String) As Boolean
    SetValueName = m_clsTableAccess.SetVal("ValueName", CVar(sValueName))
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
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetValueName()
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
Friend Function FindComboValue(sComboGroup As String, sComboValueID As String) As String
    On Error GoTo Failed
    Dim colValues As New Collection
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    Dim sValue As String
    
    sValue = ""
    sSearch = "SELECT " & _
                    "VALUENAME " & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE " & _
                    "GROUPNAME = " & g_clsSQLAssistSP.FormatString(sComboGroup) & " AND " & _
                    "VALUEID = " & sComboValueID

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount = 1 Then
            sValue = rs("ValueName")
        End If
    End If

    FindComboValue = sValue
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function GetComboValues(sGroup As String, colNames As Collection, colValues As Collection) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    bRet = False
    sSearch = "SELECT " & _
                    "* " & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE " & _
                    "GROUPNAME = " & g_clsSQLAssistSP.FormatString(sGroup)

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
            
            Do While Not rs.EOF
                colNames.Add CStr(rs("ValueName"))
                colValues.Add CStr(rs("ValueID"))
                rs.MoveNext
            Loop
        End If
    End If

    GetComboValues = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :GetComboGroupValues
' Description   : Returns a recordset containing all fields for a given combogroup
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetComboGroupValues(sGroup As String) As ADODB.Recordset
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT " & _
                "* " & _
            "FROM " & _
                m_sTableName & " " & _
            "WHERE " & _
                "GROUPNAME = " & g_clsSQLAssistSP.FormatString(sGroup)
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    Set GetComboGroupValues = rs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'SA SYS4071 Don't display option "Process Completed Applications"
Friend Function GetPaymentJobComboValues(sGroup As String, colNames As Collection, colValues As Collection, sValidationType As String) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    bRet = False
    sSearch = "SELECT " & _
                    "* " & _
                "FROM " & _
                    m_sTableName & " C " & _
                "LEFT JOIN ComboValidation CV " & _
                "ON C.GroupName = CV.GroupName " & _
                "AND C.ValueId = CV.ValueId " & _
                "WHERE " & _
                    "C.GROUPNAME = " & g_clsSQLAssistSP.FormatString(sGroup) & _
                    " AND CV.ValidationType <> " & g_clsSQLAssistSP.FormatString(sValidationType)
                    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            bRet = True
            
            Do While Not rs.EOF
                colNames.Add CStr(rs("ValueName"))
                colValues.Add CStr(rs("ValueID"))
                rs.MoveNext
            Loop
        End If
    End If

    GetPaymentJobComboValues = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


