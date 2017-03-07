VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppStageTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ApplicationStageTable
' Description   : Contains functions to read/write/search the ApplicationStage table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

' Constants
Private Const m_sTableName = "APPLICATIONSTAGE"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "APPLICATIONNUMBER"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
Public Function GetApplicationNumber() As Variant
    GetApplicationNumber = m_clsTableAccess.GetVal("APPLICATIONNUMBER")
End Function
Public Function GetAppFactFindNumber() As Variant
    GetAppFactFindNumber = m_clsTableAccess.GetVal("MAXFACTFIND")
End Function
Public Function GetAppStageName() As Variant
    GetAppStageName = m_clsTableAccess.GetVal("STAGENAME")
End Function
Public Function GetAppStageNumber() As Variant
    GetAppStageNumber = m_clsTableAccess.GetVal("STAGENUMBER")
End Function
Friend Function SetApplicationNumber(vVal As Variant) As Boolean
    SetApplicationNumber = m_clsTableAccess.SetVal("APPLICATIONNUMBER", CVar(vVal))
End Function
Friend Function SetDate(vVal As Variant)
    SetDate = m_clsTableAccess.SetVal("DATETIME", CVar(vVal))
End Function
Friend Function SetAppFactFindNumber(vVal As Variant) As Boolean
    SetAppFactFindNumber = m_clsTableAccess.SetVal("APPLICATIONFACTFINDNUMBER", CVar(vVal))
End Function
Friend Function SetAppStageNumber(vVal As Variant) As Boolean
    SetAppStageNumber = m_clsTableAccess.SetVal("STAGENUMBER", CVar(vVal))
End Function
Friend Function SetAppStageName(vVal As Variant) As Boolean
    SetAppStageName = m_clsTableAccess.SetVal("STAGENAME", CVar(vVal))
End Function
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
    colListLine.Add GetApplicationNumber()
    colListLine.Add GetAppStageName()
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetStage
' Description   : Sets the applicationnumber, stage number and stage name for the current
'                 record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetStage(sAppNo As String, sStageNo As String, sStageName As String)
    On Error GoTo Failed
    Dim sFunctionName As String
    Dim sSearch As String
    Dim sAppFactFindNo As String
    
    sFunctionName = "FindApplication"
    
    If Len(sAppNo) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + " Application number is empty"
    End If
    
    sSearch = "SELECT MAX (APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER) ""MAXFACTFIND"" FROM " + _
    " APPLICATIONFACTFIND WHERE APPLICATIONFACTFIND.APPLICATIONNUMBER = " + _
    g_clsSQLAssistSP.FormatString(sAppNo)
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    
    If m_clsTableAccess.RecordCount() > 0 Then
        ' Should be just the one
        sAppFactFindNo = GetAppFactFindNumber()
        m_clsTableAccess.CloseRecordSet
        
        If Len(sAppFactFindNo) > 0 Then
            g_clsFormProcessing.CreateNewRecord m_clsTableAccess
                    
            ' Set all the details
            SetApplicationNumber sAppNo
            SetAppFactFindNumber sAppFactFindNo
            SetAppStageNumber sStageNo
            SetAppStageName sStageName
            SetDate Now
            m_clsTableAccess.Update
        End If
    
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + " Application not found"
    End If
        
    m_clsTableAccess.CloseRecordSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function DoesStageExist(sAppNo As String, sStage As String)
    On Error GoTo Failed
    Dim bExists As Boolean
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim sFunctionName As String
    
    sFunctionName = "DoesStageExist"
    
    bExists = False
    
    If Len(sAppNo) = 0 Or Len(sStage) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": App number and stage must be set"
    End If
    
    sSearch = "SELECT APPLICATIONNUMBER FROM " + m_sTableName + " WHERE " + _
    "APPLICATIONNUMBER = " + g_clsSQLAssistSP.FormatString(sAppNo) + " AND STAGENUMBER = " + sStage + " AND " + _
    "APPLICATIONFACTFINDNUMBER = 1"
    
    Set rs = g_clsDataAccess.ExecuteCommand(sSearch)
    
    If Not rs Is Nothing Then
        If rs.RecordCount() > 0 Then
            bExists = True
        End If
    End If
    
    DoesStageExist = bExists
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
