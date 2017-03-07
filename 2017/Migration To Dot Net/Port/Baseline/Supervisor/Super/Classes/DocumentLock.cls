VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : DocumentLock
' Description   : Contains all SQL and data access routines for the FVFileLock table
' Change history
' Prog      Date        Description
' IK        25/02/2003  created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "FVLockFile"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "FileGuid"
    colFields.Add "FileVersion"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

' ik_wip_20030224
Public Sub GetDocumentLocks( _
    Optional ByVal vstrApplicationNumber As String, _
    Optional ByVal vstrUserId As String, _
    Optional ByVal vstrUnitId As String)
    
    On Error GoTo Failed
    
    Dim strSearch As String
    Dim strWhere As String
    Dim rs As ADODB.Recordset
    
    strSearch = "SELECT * FROM UVW_DMSLOCKS"
    
    If Len(vstrApplicationNumber) > 0 Then
        strWhere = " WHERE APPLICATIONNUMBER = " & vstrApplicationNumber
    End If
    
' ik_wip_20030310
'    If Len(vstrUserId) > 0 Then
'        If Len(strWhere) > 0 Then
'            strWhere = strWhere & " AND USERID = " & vstrUserId
'        Else
'            strWhere = " WHERE USERID = " & vstrUserId
'        End If
'    End If
'
'    If Len(vstrUnitId) > 0 Then
'        If Len(strWhere) > 0 Then
'            strWhere = strWhere & " AND UNITID = " & vstrUnitId
'        Else
'            strWhere = " WHERE UNITID = " & vstrUnitId
'        End If
'    End If
' ik_wip_20030310_ends
    
    If Len(strWhere) > 0 Then
        strSearch = strSearch & strWhere
    End If
    
    Set rs = g_clsDataAccess.GetTableData("UVW_DMSLOCKS", strSearch, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteDocumentLock(ByRef fcolKeys As Collection)
    On Error GoTo Failed
    
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    strSQL = _
        "DELETE FROM FVLockFile WHERE FileGuid =0x" & _
        g_clsSQLAssistSP.GuidToString(fcolKeys.Item(1)) & _
        " AND FileVersion = '" & _
        fcolKeys.Item(2) & "'"
    
    g_clsDataAccess.ExecuteCommand strSQL
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions for FVLockFile View
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetApplicationNumber() As Variant
    GetApplicationNumber = m_clsTableAccess.GetVal("ApplicationNumber")
End Function
' ik_wip_20030310
'Friend Function GetUnitID() As Variant
'    GetUnitID = m_clsTableAccess.GetVal("UnitID")
'End Function
'Friend Function GetUserID() As Variant
'    GetUserID = m_clsTableAccess.GetVal("UserID")
'End Function
Friend Function GetLockedBy() As Variant
    GetLockedBy = m_clsTableAccess.GetVal("LockedBy")
End Function
' ik_wip_20030310_ends
Friend Function GetDateLocked() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("DateLocked")
    
    If Not IsNull(vVal) Then
        GetDateLocked = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetFileName() As Variant
    GetFileName = m_clsTableAccess.GetVal("FileName")
End Function
Friend Function GetFileVersion() As Variant
    GetFileVersion = m_clsTableAccess.GetVal("FileVersion")
End Function
Friend Function GetFileGuid() As Variant
    GetFileGuid = m_clsTableAccess.GetVal("FileGuid")
End Function

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub
Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub
Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub
Private Sub TableAccess_DeleteAllRows()
    m_clsTableAccess.DeleteAllRows
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub TableAccess_DeleteRow(colMatchData As Collection)
    m_clsTableAccess.DeleteRow colMatchData
End Sub
Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function
Public Function TableAccess_GetDatabase() As String
    TableAccess_GetDatabase = m_clsTableAccess.GetDatabase()
End Function
Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function
Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_clsTableAccess.GetIsDeleted()
End Function
Private Function TableAccess_GetKeyMatchFields()
    Set TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function
Private Function TableAccess_GetKeyMatchValues() As Variant
    Set TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function
Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function
Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetApplicationNumber()
    colListLine.Add GetDateLocked()
' ik_wip_20030310
'    colListLine.Add GetUserID()
'    colListLine.Add GetUnitID()
    colListLine.Add GetLockedBy()
' ik_wip_20030310_ends
    colListLine.Add GetFileName()
    colListLine.Add GetFileVersion()
End Sub
Private Function TableAccess_GetSearch() As String
    TableAccess_GetSearch = m_sSearch
End Function
Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function
Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_clsTableAccess.GetTableDataFromSearch sSearch
End Sub
Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function
Public Function TableAccess_GetUpdateValues() As Collection
End Function
Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function
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
Private Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub
Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub
Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vVal)
End Function
Private Sub TableAccess_SetDeleteFields(colFields As Collection)
    m_clsTableAccess.SetDeleteFields colFields
End Sub
Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)
    m_clsTableAccess.SetKeyMatchValues colValues
End Sub
Private Function TableAccess_SetNextRow() As Boolean
    TableAccess_SetNextRow = m_clsTableAccess.SetNextRow()
End Function
Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_clsTableAccess.SetOrderCriteria sCriteria
End Sub
Private Sub TableAccess_SetPopulateNull()
    m_clsTableAccess.SetPopulateNull
End Sub
Private Sub TableAccess_SetPopulateType(enumType As PopulateType)
    
End Sub
Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub
Private Sub TableAccess_SetUpdated()
    m_clsTableAccess.SetUpdated
End Sub

Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
End Function
Private Sub TableAccess_Update()
End Sub
Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
End Sub
