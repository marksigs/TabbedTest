VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AvailableTemplatesTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : AvailableTemplatesTable
' Description   : Contains all data access functions for the AvailableTemplates Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private Const m_sTableName = "AVAILABLETEMPLATES"
Private Const m_sHostTableName = "HOSTTEMPLATE"
Private Const m_sStageTableName = "STAGE"

Private Const m_sApplicationStageComboGroup = "ApplicationStage"
Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "HOSTTEMPLATEID"
    
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

Public Function GetUnassignedStages(sTemplateID As String) As ADODB.Recordset

    On Error GoTo Failed
    Dim sSQL As String
    
    If Len(sTemplateID) > 0 Then
    
        sSQL = "Select * from " & m_sStageTableName & " Where StageID " & _
               "Not in (Select StageID from " & m_sTableName & " where HostTemplateID = " & g_clsSQLAssistSP.FormatString(sTemplateID) & _
               " ) And DeleteFlag = 0"
    Else
        sSQL = "Select * from " & m_sStageTableName & " Where DeleteFlag = 0"
    
    End If
    
    Set GetUnassignedStages = g_clsDataAccess.GetTableData(, sSQL)
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub GetSelectedStages(sTemplateID As String, Optional bIncludStageName As Boolean = False)

    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    If bIncludStageName Then
'        ' DJP SQL Server port, put Outer Join in Stored Proc
'        Dim Param As ADODB.Parameter
'        Dim colParam As Collection
'
'        ' First parameter, the GroupName
'        Set Param = New ADODB.Parameter
'        Set colParam = New Collection
'
'        With Param
'            .Name = "TemplateID"
'            .Type = adBSTR
'            .Direction = adParamInput
'            .Value = sTemplateID
'            ' SQL Server requires the length
'            .Size = Len(.Value)
'        End With
'
'        colParam.Add Param
'
'        ' Get the data
'        Set rs = g_clsDataAccess.ExecuteStoredProcRS(spGetSelectedStages, colParam)
        sSQL = "Select " & m_sTableName & ".*,  StageName from " & m_sTableName & ", " & m_sStageTableName & _
               " Where " & m_sTableName & ".HostTemplateID = " & g_clsSQLAssistSP.FormatString(sTemplateID) & " And " & m_sStageTableName & ".Stageid = " & _
                m_sTableName & ".Stageid"
            
    Else
        sSQL = "SELECT " & _
                    "* " & _
                "FROM " & _
                    m_sTableName & " " & _
                "WHERE " & _
                    "HostTemplateID = " & g_clsSQLAssistSP.FormatString(sTemplateID)
    End If
    Set rs = g_clsDataAccess.GetTableData(, sSQL, m_clsTableAccess.GetDatabase)
    
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Friend Function GetStageID() As Variant
    GetStageID = m_clsTableAccess.GetVal("STAGEID")
End Function
Friend Function SetStageID(vVal As Variant) As Boolean
    SetStageID = m_clsTableAccess.SetVal("STAGEID", vVal)
End Function
Friend Function GetTemplateID() As Variant
    GetTemplateID = m_clsTableAccess.GetVal("HOSTTEMPLATEID")
End Function
Friend Function SetTemplateID(vVal As Variant) As Boolean
    SetTemplateID = m_clsTableAccess.SetVal("HOSTTEMPLATEID", vVal)
End Function
Friend Function GetStageName() As Variant
    GetStageName = m_clsTableAccess.GetVal("STAGENAME")
End Function
Friend Function SetStageName(vVal As Variant) As Boolean
    SetStageName = m_clsTableAccess.SetVal("STAGENAME", vVal)
End Function
Public Sub GetSelectedStagesFromComboGroup(sTemplateID As String, Optional bIncludeName As Boolean = True)
    
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    If bIncludeName Then
        sSQL = "Select " & m_sTableName & ".*, ValueName StageName from " & _
                m_sTableName & ", ComboValue  StageName " & _
                "Where HostTemplateID = " & sTemplateID & " and " & m_sTableName & _
                ".StageID  = StageName.ValueID AND StageName.GroupName = " & g_clsSQLAssistSP.FormatString(m_sApplicationStageComboGroup)
    Else
        sSQL = "select * from " & m_sTableName & " Where HostTemplateID = " & sTemplateID
    End If
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    m_clsTableAccess.SetRecordSet rs
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
