VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PackMemberTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : PackMember
' Description  : Helper-class which contains Intermediary routines related to
'                the tree control found on frmMain.
'
' Change history
' Prog      Date        Description
' GHun      17/10/2005  MAR202 Created
' RF        05/12/2005  MAR202 Handle PackMembers
' RF        14/12/2005  MAR867  Complete pack handling changes
' HMA       01/02/2006  MAR967 Change for Pack Member promotion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName As String = "PACKMEMBER"

Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "PACKCONTROLNUMBER"
   'colFields.Add "PACKMEMBERNUMBER"                      ' MAR967
    m_clsTableAccess.SetKeyMatchFields colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Friend Function GetPackControlNumber() As Variant
    GetPackControlNumber = m_clsTableAccess.GetVal("PACKCONTROLNUMBER")
End Function
Friend Function SetPackControlNumber(vVal As Variant) As Boolean
    SetPackControlNumber = m_clsTableAccess.SetVal("PACKCONTROLNUMBER", vVal)
End Function
Friend Function GetPackMemberNumber() As Variant
    GetPackMemberNumber = m_clsTableAccess.GetVal("PACKMEMBERNUMBER")
End Function
Friend Function SetPackMemberNumber(vVal As Variant) As Boolean
    SetPackMemberNumber = m_clsTableAccess.SetVal("PACKMEMBERNUMBER", vVal)
End Function
Friend Function GetPackMemberType() As Variant
    GetPackMemberType = m_clsTableAccess.GetVal("PACKMEMBERTYPE")
End Function
Friend Function SetPackMemberType(vVal As Variant) As Boolean
    SetPackMemberType = m_clsTableAccess.SetVal("PACKMEMBERTYPE", vVal)
End Function
Friend Function GetPackMemberControlNumber() As Variant
    GetPackMemberControlNumber = m_clsTableAccess.GetVal("PACKMEMBERCONTROLNUMBER")
End Function
Friend Function SetPackMemberControlNumber(vVal As Variant) As Boolean
    SetPackMemberControlNumber = m_clsTableAccess.SetVal("PACKMEMBERCONTROLNUMBER", vVal)
End Function
Friend Function GetHostTemplateID() As Variant
    GetHostTemplateID = m_clsTableAccess.GetVal("HOSTTEMPLATEID")
End Function
Friend Function SetHostTemplateId(vVal As Variant) As Boolean
    SetHostTemplateId = m_clsTableAccess.SetVal("HOSTTEMPLATEID", vVal)
End Function

'Friend Function GetHostTemplateName(ByVal vstrHOSTTEMPLATEID As String) As Variant
'On Error GoTo Failed
'    Dim sSearch As String
'    Dim rs As ADODB.Recordset
'    Dim strHostTemplateName As String
'
'    strHostTemplateName = ""
'
'    If Len(vstrHOSTTEMPLATEID) = 0 Then
'        g_clsErrorHandling.RaiseError errGeneralError, _
'            " GetHostTemplateName: HostTemplate ID is empty"
'    End If
'
'    sSearch = "SELECT HOSTTEMPLATENAME " _
'        & " FROM HOSTTEMPLATE " _
'        & " WHERE HOSTTEMPLATEID = '" & vstrHOSTTEMPLATEID & "'"
'
'    Set rs = g_clsDataAccess.ExecuteCommand(sSearch, m_clsTableAccess.GetDatabase)
'
'    If rs.RecordCount > 0 Then
'        If Not IsNull(rs.fields(0).Value) Then
'            strHostTemplateName = rs.fields(0).Value
'        End If
'    End If
'
'    GetHostTemplateName = strHostTemplateName
'
'    Exit Function
'
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'
'End Function
'Friend Function GetHostTemplateDescription(ByVal vstrHOSTTEMPLATEID As String) As Variant
'On Error GoTo Failed
'    Dim sSearch As String
'    Dim rs As ADODB.Recordset
'    Dim strHostTemplateDesc As String
'
'    strHostTemplateDesc = ""
'
'    If Len(vstrHOSTTEMPLATEID) = 0 Then
'        g_clsErrorHandling.RaiseError errGeneralError, _
'            " GetHostTemplateDescription: HostTemplate ID is empty"
'    End If
'
'    sSearch = "SELECT HOSTTEMPLATEDESCRIPTION " _
'        & " FROM HOSTTEMPLATE " _
'        & " WHERE HOSTTEMPLATEID = '" & vstrHOSTTEMPLATEID & "'"
'
'    Set rs = g_clsDataAccess.ExecuteCommand(sSearch, m_clsTableAccess.GetDatabase)
'
'    If rs.RecordCount > 0 Then
'        If Not IsNull(rs.fields(0).Value) Then
'            strHostTemplateDesc = rs.fields(0).Value
'        End If
'    End If
'
'    GetHostTemplateDescription = strHostTemplateDesc
'
'    Exit Function
'
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'
'End Function

Friend Function GetTableData(ByVal vintPackControlNumber As Integer) As Variant
On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
    If Len(vintPackControlNumber) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, _
            " GetTableData: Pack Control Number is empty"
    End If
    
    sSearch = "SELECT * " _
        & " FROM PACKMEMBER " _
        & " WHERE PACKCONTROLNUMBER = " & vintPackControlNumber
        
    Set rs = g_clsDataAccess.GetTableData( _
        m_sTableName, sSearch, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rs
        
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function

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
    colListLine.Add GetPackControlNumber
    colListLine.Add GetPackMemberNumber
    colListLine.Add GetPackMemberType
    colListLine.Add GetPackMemberControlNumber
    colListLine.Add GetHostTemplateID
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

'MAR967  Add new function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetMembersForPack
' Description   :   Returns a list of Pack Member entries for the pack passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetMembersForPack(ByVal sPackID As String)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rsTemplates As ADODB.Recordset
    Dim sTarget As String
    
    sPackID = g_clsSQLAssistSP.FormatString(sPackID)

    sSearch = "SELECT * FROM " & m_sTableName & " WHERE PACKCONTROLNUMBER = " & sPackID
    
    sTarget = m_clsTableAccess.GetDatabase()
    Set rsTemplates = g_clsDataAccess.GetTableData(m_sTableName, sSearch, sTarget)

    m_clsTableAccess.SetRecordSet rsTemplates

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
