VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmPermissionsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   FirmPermissionsTable
' Description   :   Contains all access functions for the FirmPermissions table.
' Change history
' Change history
' Prog     Date         Description
' TW       17/10/2006   EP2_15 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "FIRMPERMISSIONS"
Private Const DELETE_FLAG_SET As Integer = 1

Private m_sFirstMatch As String

Friend Function GetCategory() As String
    GetCategory = CStr(m_clsTableAccess.GetVal("Category"))
End Function

Friend Sub SetFindBrokerLinkedActivityFSARecords(strBrokerID As String)
    Dim sSysDate As String
    
    m_sFirstMatch = "SELECT P.*, A.ACTIVITYID, A.CATEGORY, A.ACTIVITYDESCRIPTION FROM ACTIVITYFSA A INNER JOIN FIRMPERMISSIONS P ON A.ACTIVITYID = P.ACTIVITYID " & _
                    "WHERE P.ARFIRMID = '" & strBrokerID & "'"
    
    
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub

Friend Sub SetFindFirmLinkedActivityFSARecords(strPackagerID As String)
    
    m_sFirstMatch = "SELECT P.*, A.ACTIVITYID, A.CATEGORY, A.ACTIVITYDESCRIPTION FROM ACTIVITYFSA A INNER JOIN FIRMPERMISSIONS P ON A.ACTIVITYID = P.ACTIVITYID " & _
                    "WHERE P.PRINCIPALFIRMID = '" & strPackagerID & "'"
    
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
'    colFields.Add "PRINCIPALFIRMID"
    colFields.Add "FIRMPERMISSIONSSEQNO"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
'    m_clsTableAccess.SetOrderCriteria "PRINCIPALFIRMID"
    m_clsTableAccess.SetOrderCriteria "FIRMPERMISSIONSSEQNO"


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetPrincipalFirmID() As String
    GetPrincipalFirmID = CStr(m_clsTableAccess.GetVal("PrincipalFirmID"))
End Function
Friend Function GetARFirmID() As String
    GetARFirmID = CStr(m_clsTableAccess.GetVal("ARFirmID"))
End Function
Friend Function GetActivityID() As String
    GetActivityID = CStr(m_clsTableAccess.GetVal("ActivityID"))
End Function
Friend Function GetFRMPermissions() As String
    GetFRMPermissions = CStr(m_clsTableAccess.GetVal("FRMPermissions"))
End Function
Friend Function GetEffectiveDate() As String
    GetEffectiveDate = CStr(m_clsTableAccess.GetVal("EffectiveDate"))
End Function
Friend Function GetLastUpdatedDate() As String
    GetLastUpdatedDate = CStr(m_clsTableAccess.GetVal("LastUpdatedDate"))
End Function
Friend Function GetFirmPermissionsSeqNo() As String
    GetFirmPermissionsSeqNo = CStr(m_clsTableAccess.GetVal("FirmPermissionsSeqNo"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetPrincipalFirmID(sValue As Variant) As Boolean
    SetPrincipalFirmID = m_clsTableAccess.SetVal("PrincipalFirmID", CVar(sValue))
End Function
Friend Function SetARFirmID(sValue As Variant) As Boolean
    SetARFirmID = m_clsTableAccess.SetVal("ARFirmID", CVar(sValue))
End Function
Friend Function SetActivityID(sValue As Variant) As Boolean
    SetActivityID = m_clsTableAccess.SetVal("ActivityID", CVar(sValue))
End Function
Friend Function SetFRMPermissions(sValue As Variant) As Boolean
    SetFRMPermissions = m_clsTableAccess.SetVal("FRMPermissions", CVar(sValue))
End Function
Friend Function SetEffectiveDate(sValue As Variant) As Boolean
    SetEffectiveDate = m_clsTableAccess.SetVal("EffectiveDate", CVar(sValue))
End Function
Friend Function SetLastUpdatedDate(sValue As Variant) As Boolean
    SetLastUpdatedDate = m_clsTableAccess.SetVal("LastUpdatedDate", CVar(sValue))
End Function
Friend Function SetFirmPermissionSeqNo(sValue As Variant) As Boolean
    SetFirmPermissionSeqNo = m_clsTableAccess.SetVal("FirmPermissionsSeqNo", CVar(sValue))
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
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
    colListLine.Add GetFirmPermissionsSeqNo()
    colListLine.Add GetPrincipalFirmID()
    colListLine.Add GetARFirmID()
    colListLine.Add GetActivityID()
    colListLine.Add GetDescription()
    colListLine.Add GetCategory()

End Sub

Friend Function GetDescription() As String
    GetDescription = CStr(m_clsTableAccess.GetVal("ActivityDescription"))
End Function

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
