VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboValueGroupTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "COMBOGROUP"
Private m_sSearch As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History:
'
'Prog   Date        AQR     Description
'MV     15/01/2003  BM0085  Added a New function CreateComboAuditRecords
'MV     28/02/2003  BM0402  Amended DeleteComboGroup
'MV     04/04/2003  BM0402  Amended DeleteComboGroup
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateComboAuditRecords(ByVal sGroupName As String, Optional strDBName As String) As Boolean
    
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim ComboValidationRS As ADODB.Recordset
    Dim sValueId  As String
    Dim sValidationType As String
    Dim sValueName As String
    Dim ComboValueRS  As ADODB.Recordset
    
'Stage1:
    
    sSQL = ""
    sSQL = sSQL + " SELECT * FROM "
    sSQL = sSQL + " COMBOVALIDATION WHERE  GROUPNAME = '" + sGroupName + "' "
    
    Set ComboValidationRS = g_clsDataAccess.GetTableData("COMBOVALIDATION", sSQL, strDBName)
       
    If ComboValidationRS.RecordCount > 0 Then
        
        While Not ComboValidationRS.EOF
             
            sValueId = ComboValidationRS.fields("VALUEID").Value
            sValidationType = ComboValidationRS.fields("VALIDATIONTYPE").Value
            
            frmNewCombo.CreateComboValidationAuditRecord sGroupName, sValueId, g_sSupervisorUser, sValidationType, "", "D", strDBName
            
            ComboValidationRS.MoveNext
            
        Wend
        
    End If
    
'Stage2:

    sSQL = ""
    sSQL = sSQL + " SELECT * FROM COMBOVALUE WHERE  GROUPNAME = '" + sGroupName + "' "
    
    Set ComboValueRS = g_clsDataAccess.GetTableData("COMBOVALUE", sSQL, strDBName)
        
    If ComboValueRS.RecordCount > 0 Then
        
        While Not ComboValueRS.EOF
             
            sValueId = ComboValueRS.fields("VALUEID").Value
            sValueName = ComboValueRS.fields("VALUENAME").Value
            
            frmNewCombo.CreateComboValueAuditRecord sGroupName, sValueId, g_sSupervisorUser, sValueName, "", "D", strDBName
            
            ComboValueRS.MoveNext
            
        Wend
        
    End If
    
    Set ComboValidationRS = Nothing
    Set ComboValueRS = Nothing
    
    CreateComboAuditRecords = True
            
    Exit Function
        
Failed:

    Set ComboValidationRS = Nothing
    Set ComboValueRS = Nothing
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "GroupName"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Friend Function GetGroupName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("GroupName")
    GetGroupName = CStr(vVal)
End Function

Friend Function GetGroupNote() As String
    GetGroupNote = CStr(m_clsTableAccess.GetVal("GroupNote"))
End Function

Friend Function SetGroupName(sGroupName As String) As Boolean
    SetGroupName = m_clsTableAccess.SetVal("GroupName", CVar(sGroupName))
End Function

Friend Function SetGroupNote(sGroupNote As String) As Boolean
    SetGroupNote = m_clsTableAccess.SetVal("GroupNote", CVar(sGroupNote))
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

Public Function TableAccess_GetUpdateValues() As Collection
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetGroupName()
    colListLine.Add GetGroupNote()
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

Public Sub DeleteComboGroup(Optional sTargetDatabase As String)
    On Error GoTo Failed
    Dim colMatchValues As Collection
    
    If Len(sTargetDatabase) = 0 Then
        sTargetDatabase = m_clsTableAccess.GetDatabase()
    End If
    
    'MV - 18/01/2003 - BM0085
    Dim sGroupName  As String
    Dim bRet As Boolean
    
    'g_clsDataAccess.BeginTrans
    ' First, any validation records
    Dim clsTableAccess As TableAccess
    Dim clsValidation As New ComboValidationTable
    Set clsTableAccess = clsValidation
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    'MV -  28/02/2003 - BM0402
    sGroupName = colMatchValues(1)
    bRet = CreateComboAuditRecords(sGroupName, sTargetDatabase)
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Error in Creating Combo Audit Records"
        Exit Sub
    End If
    
    clsValidation.DeleteValidation sTargetDatabase
    
    ' Next, Value records
    Set clsTableAccess = New ComboValueTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase

    ' Last, the group name
    Set clsTableAccess = New ComboValueGroupTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.DeleteRecords sTargetDatabase
    'g_clsDataAccess.CommitTrans
    Exit Sub
Failed:
    'g_clsDataAccess.RollbackTrans
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

