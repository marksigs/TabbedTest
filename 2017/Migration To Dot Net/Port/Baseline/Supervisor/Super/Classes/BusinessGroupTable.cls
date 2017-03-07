VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessGroupTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : BusinessGroupTable
' Description   : Contains all data access functions for the CaseTrack Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DJP       22/06/01    SQL Server port - don't use local SQLAssist, use global Supervisor one
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Constants
Private Const m_sTableName = "CASETRACK"
Private Const m_sBusinessAreaComboGroup = "CTBusinessArea"

Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "GROUPID"
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
    ' DJP SQL Server port
    colListLine.Add GetGroupID, OBJECT_EXTRA_VALUE
    colListLine.Add GetGroupName, OBJECT_DESCRIPTION_KEY
    colListLine.Add GetGroupName
    colListLine.Add GetBusinessAreaText
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

Friend Function GetGroupID() As Variant
    GetGroupID = m_clsTableAccess.GetVal("GROUPID")
End Function
Friend Function GetGroupName() As Variant
    GetGroupName = m_clsTableAccess.GetVal("GROUPNAME")
End Function
Friend Function GetBusinessArea() As Variant
    GetBusinessArea = m_clsTableAccess.GetVal("BUSINESSAREA")
End Function
Friend Function GetBusinessAreaText() As Variant
    GetBusinessAreaText = m_clsTableAccess.GetVal("BUSINESSAREATEXT")
End Function
Friend Function GetCompleteText() As Variant
    GetCompleteText = m_clsTableAccess.GetVal("COMPLETETEXT")
End Function
Friend Function GetInCompleteText() As Variant
    GetInCompleteText = m_clsTableAccess.GetVal("INCOMPLETETEXT")
End Function
Friend Function GetNotApplicableText() As Variant
    GetNotApplicableText = m_clsTableAccess.GetVal("NOTAPPLICABLETEXT")
End Function
Friend Function GetNotStartedText() As Variant
    GetNotStartedText = m_clsTableAccess.GetVal("NOTSTARTEDTEXT")
End Function
Friend Function SetGroupID(vVal As Variant) As Boolean
    SetGroupID = m_clsTableAccess.SetVal("GROUPID", vVal)
End Function
Friend Function SetGroupName(vVal As Variant) As Boolean
    SetGroupName = m_clsTableAccess.SetVal("GROUPNAME", vVal)
End Function
Friend Function SetBusinessArea(vVal As Variant) As Boolean
    SetBusinessArea = m_clsTableAccess.SetVal("BUSINESSAREA", vVal)
End Function
Friend Function SetCompleteText(vVal As Variant) As Boolean
    SetCompleteText = m_clsTableAccess.SetVal("COMPLETETEXT", vVal)
End Function
Friend Function SetInCompleteText(vVal As Variant) As Boolean
    SetInCompleteText = m_clsTableAccess.SetVal("INCOMPLETETEXT", vVal)
End Function
Friend Function SetNotApplicableText(vVal As Variant) As Boolean
    SetNotApplicableText = m_clsTableAccess.SetVal("NOTAPPLICABLETEXT", vVal)
End Function
Friend Function SetNotStartedText(vVal As Variant) As Boolean
    SetNotStartedText = m_clsTableAccess.SetVal("NOTSTARTEDTEXT", vVal)
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetGroups
' Description   : Populates the frmMain Listview with all records from the BatchJobs Table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetGroups()

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
    ' DJP SQL Server port - don't use local SQLAssist, use global Supervisor one
    sSQL = "SELECT " & m_sTableName & ".*, ComboBusinessType.ValueName BusinessAreaText " & _
           " FROM combovalue ComboBusinessType, " & m_sTableName & _
           " WHERE" & _
           " ComboBusinessType.GroupName = " & g_clsSQLAssistSP.FormatString(m_sBusinessAreaComboGroup) & _
           " AND ComboBusinessType.ValueID =  BusinessArea"

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    ValidateRecordset rs, "CaseTrack"
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetTasksForGroup(vGroupID As Variant)
    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select * from " & m_sTableName & " Where GroupID = " & g_clsSQLAssistSP.FormatString(vGroupID)
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    m_clsTableAccess.SetRecordSet rs
    

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
