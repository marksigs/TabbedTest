VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IncomeMultipleSetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : IncomeMultipleSetTable
' Description   : Contains all access routines and SQL for the IncomeMultipleSet table.
'                 This class implements both TableAccess for generic routines..
' Change history
' Prog      Date        Description
' AW      13/05/02      Created
' GD      23/05/02      BMIDS0012
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
' Constants
Private Const m_sTableName = "INCOMEMULTIPLESET"

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "IncomeMultiplierCode"
    
    'If g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False) Then
    '    colFields.Add m_sVersionField
    'End If
    
    m_clsTableAccess.SetKeyMatchFields colFields

End Sub

Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetIncMultCode()
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIncMultCode() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("IncomeMultiplierCode")
    GetIncMultCode = vVal
End Function

Friend Function GetIncMultDesc() As String
    GetIncMultDesc = CStr(m_clsTableAccess.GetVal("IncomeMultiplierDescription"))
End Function

Friend Function GetSingleIncMult() As String
    GetSingleIncMult = CStr(m_clsTableAccess.GetVal("SingleIncomeMultiple"))
End Function

Friend Function GetJointIncMult() As String
    GetJointIncMult = CStr(m_clsTableAccess.GetVal("JointIncomeMultiple"))
End Function

Friend Function GetHighEarnerMult() As String
    GetHighEarnerMult = CStr(m_clsTableAccess.GetVal("HighestEarnerMultiple"))
End Function
Friend Function GetLowEarnerMult() As String
    GetLowEarnerMult = CStr(m_clsTableAccess.GetVal("LowestEarnerMultiple"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIncMultCode(sVal As Variant) As Boolean
    SetIncMultCode = m_clsTableAccess.SetVal("IncomeMultiplierCode", CVar(sVal))
End Function

Friend Function SetIncMultDesc(sIncMultDesc As String) As Boolean
    SetIncMultDesc = m_clsTableAccess.SetVal("IncomeMultiplierDescription", CVar(sIncMultDesc))
End Function

Friend Function SetSingleIncMult(sSingleIncMult As String) As Boolean
    SetSingleIncMult = m_clsTableAccess.SetVal("SingleIncomeMultiple", CVar(sSingleIncMult))
End Function

Friend Function SetJointIncMult(sJointIncMult As String) As Boolean
    SetJointIncMult = m_clsTableAccess.SetVal("JointIncomeMultiple", CVar(sJointIncMult))
End Function

Friend Function SetHighEarnerMult(sHighEarnMult As String) As Boolean
    SetHighEarnerMult = m_clsTableAccess.SetVal("HighestEarnerMultiple", CVar(sHighEarnMult))
End Function

Friend Function SetLowEarnerMult(sLowEarnMult As String) As Boolean
    SetLowEarnerMult = m_clsTableAccess.SetVal("LowestEarnerMultiple", CVar(sLowEarnMult))
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

'Private Function TableAccess_DoUpdateSets() As Boolean
'    TableAccess_DoUpdateSets = m_clsTableAccess.DoUpdateSets()
'End Function
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

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    m_clsTableAccess = clsTable
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
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
    colListLine.Add GetIncMultDesc
    colListLine.Add GetSingleIncMult()
    colListLine.Add GetJointIncMult()
    colListLine.Add GetHighEarnerMult()
    colListLine.Add GetLowEarnerMult()
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

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")

    On Error GoTo Failed
    Dim sMessage As String
    Dim inum As Integer
    Dim colValues  As New Collection
    Set colValues = m_clsTableAccess.GetKeyMatchValues
    
    Dim clsMortgageProductTable As MortgageProductTable
    Set clsMortgageProductTable = New MortgageProductTable
    
    Set colValues = clsMortgageProductTable.DoesProductExistForThisItem("INCOMEMULTIPLIERCODE", colValues.Item(1))
    
    If (Not colValues Is Nothing) And (colValues.Count > 0) Then
        sMessage = ""
        For inum = 1 To colValues.Count
            sMessage = sMessage & "(" & inum & ") " & colValues.Item(inum) & vbCrLf
        Next
        
        MsgBox "Record can not be deleted. The following Mortgage Products exist for the selected Income Multiple set : " & vbCrLf & sMessage, vbCritical, "Cannot Delete"
        g_clsErrorHandling.RaiseError , "Record can not be deleted. The following Mortgage Products exist for the selected Income Multiple set : " & vbCrLf & sMessage, ErrorSystem
    Else
        m_clsTableAccess.DeleteRecords sDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetIncomeMultipleSet() As Boolean
'GD BMIDS0012
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    
 
    sSearch = "SELECT CONVERT(nvarchar,INCOMEMULTIPLIERCODE) + ' - ' + INCOMEMULTIPLIERDESCRIPTION 'CONCATFIELD',INCOMEMULTIPLIERCODE FROM INCOMEMULTIPLESET order by INCOMEMULTIPLIERCODE"
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    m_clsTableAccess.SetRecordSet rs

    If Not rs Is Nothing Then
        bRet = True
    Else
        bRet = False
    End If
End Function
