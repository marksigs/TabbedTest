VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserRoleTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' BMIDS Change history
' Prog      Date        Description
' BS    26/03/03    BM0423 Remove UnitID from colFields collection

Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "USERROLE"
Private m_sSearch As String

Public Sub GetUserRolesForUserID(ByVal sUserID As String)
    
    On Error GoTo Failed
    Dim sSearch As String
    Dim RsUserRoles As ADODB.Recordset
    Dim nExceptionStage As Integer
    
    If Len(sUserID) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, " GetUserRolesForUserID: UserId is empty"
    End If
    
    sSearch = ""
    sSearch = sSearch + "SELECT * FROM USERROLE "
    sSearch = sSearch + " WHERE "
    sSearch = sSearch + " USERID = '" & sUserID + "' "
    sSearch = sSearch & " ORDER BY UNITID"
    
    Set RsUserRoles = g_clsDataAccess.GetTableData(m_sTableName, sSearch, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet RsUserRoles
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colDupKeys As Collection
    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "UserID"
    'BS BM0423 26/03/03
    'Remove UnitID so all UserRole records for the User are retrieved
    'colFields.Add "UnitID"
    m_clsTableAccess.SetKeyMatchFields colFields

    Set colDupKeys = New Collection
    colDupKeys.Add "UNITID"

    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetUserID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("UserID")
    GetUserID = CStr(vVal)
End Function

Friend Function GetUnitID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("UnitID")
    GetUnitID = CStr(vVal)
End Function
Friend Function GetRole() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("Role")
    GetRole = CStr(vVal)
End Function
Friend Function GetRoleText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("UserRoleText")
    GetRoleText = CStr(vVal)
End Function
Friend Function GetActiveFrom() As String
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("UserRoleActiveFrom")
        
    If Not IsNull(vVal) Then
        GetActiveFrom = Format(vVal, "Short Date")
    End If
End Function
Friend Function GetActiveTo() As String
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("UserRoleActiveTo")
        
    If Not IsNull(vVal) Then
        GetActiveTo = Format(vVal, "Short Date")
    End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetUserID(sUserID As Variant) As Boolean
    SetUserID = m_clsTableAccess.SetVal("UserID", CVar(sUserID))
End Function
Friend Function SetUnitID(sVal As Variant) As Boolean
    SetUnitID = m_clsTableAccess.SetVal("UnitID", CVar(sVal))
End Function
Friend Function SetRole(sVal As Variant) As Boolean
    SetRole = m_clsTableAccess.SetVal("Role", CVar(sVal))
End Function
Friend Function SetRoleText(sVal As Variant) As Boolean
    SetRoleText = m_clsTableAccess.SetVal("UserRoleText", CVar(sVal))
End Function
Public Function GetComboFields() As Collection
    
    On Error GoTo Failed
    Dim colFields As Collection
    
    Set colFields = New Collection
    
    colFields.Add "UNITID"
    colFields.Add "UNITNAME"
    
    Set GetComboFields = colFields
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Sub GetUnits(Optional colExcludeUnits As Collection, Optional ByVal sExcludeUserID As String = "", Optional colUnitMatch As Collection = Nothing, Optional bCheckDeleteFlag As Boolean = True)
    
    On Error GoTo Failed
    Dim nCount As Long
    Dim nThisItem As Long
    Dim sSearch As String
    Dim sUnitID As String
    Dim sExcludeUnit As Variant
    Dim rsUnits As ADODB.Recordset
    Dim bExcludeUser As Boolean
    
    sSearch = ""
    sSearch = "SELECT * FROM UNIT "
    
    If Not colUnitMatch Is Nothing Then
        If colUnitMatch.Count > 0 Then
            
            sSearch = sSearch + " WHERE ("
            nCount = colUnitMatch.Count
            
            For nThisItem = 1 To nCount
                
                sUnitID = colUnitMatch(nThisItem)
                sUnitID = g_clsSQLAssistSP.FormatString(sUnitID)
                
                sSearch = sSearch + " UNITID = " & sUnitID
                            
                If nThisItem < nCount Then
                    sSearch = sSearch + " OR "
                End If
            
            Next
            sSearch = sSearch + ")"
        End If
    End If
    
    If Not colExcludeUnits Is Nothing Then
        If colExcludeUnits.Count > 0 Then
            Dim nCtr As Integer
            sSearch = sSearch & " WHERE  "
            sSearch = sSearch & "  UNITID  NOT IN ( "
            
            For Each sExcludeUnit In colExcludeUnits
                
                If Len(sExcludeUnit) > 0 Then
                    sExcludeUnit = g_clsSQLAssistSP.FormatString(sExcludeUnit)
                    sSearch = sSearch & sExcludeUnit
                    nCtr = nCtr + 1
                    
                    If nCtr < colExcludeUnits.Count Then
                        sSearch = sSearch & " , "
                    End If
                    
                    
                End If
            Next
            sSearch = sSearch & " )"
        End If
    End If
    
    Set rsUnits = g_clsDataAccess.GetTableData("UNIT", sSearch)
    
    m_clsTableAccess.SetRecordSet rsUnits
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Function SetActiveFrom(sVal As Variant) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("UserRoleActiveFrom", CVar(sVal))
End Function
Friend Function SetActiveTo(sVal As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("UserRoleActiveTo", CVar(sVal))
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
' DJP This can be added later if required.
'Public Function TableAccess_GetUpdateFields() As Collection
'End Function
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetUnitID()
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
    colListLine.Add GetRoleText()
   
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

Public Sub FindRecord(sUserID As String, sUnitID As String)
    On Error GoTo Failed

    Dim sFilter As String
    
    sFilter = "USERID = '" & sUserID & "' AND UNITID = '" & sUnitID & "' "
    
    m_clsTableAccess.ApplyFilter sFilter
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
