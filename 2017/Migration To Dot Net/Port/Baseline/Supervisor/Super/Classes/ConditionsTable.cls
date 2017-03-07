VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConditionsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ConditionsTable.cls
' Description   : Contains all data access routines for the Conditions table.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SQL Server port, use global SQLAssist not local one
' DJP       27/06/01    SQL Server port
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const m_sComboGroup = "ConditionType"
Private Const m_sTableName = "CONDITIONS"

' Private data
Private m_sQuestionReference As String
Private m_clsTableAccess As TableAccess

Implements TableAccess
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' DJP       05/03/03    BM0282 Added Conditions search facility.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM Change history
' TW        02/01/2007  EP2_640 Deleted Conditions still being shown on frmMain.lvListView
'                               Deletions not being promoted correctly
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "CONDITIONREFERENCE"
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
    'TableAccess_GetSearch = m_sSearch
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

    colListLine.Add GetConditionReference
    colListLine.Add GetConditionName
    colListLine.Add GetConditionDescription
    colListLine.Add GetConditionTypeText
    colListLine.Add GetBooleanFromNumber(GetConditionEditFlag)
    colListLine.Add GetBooleanFromNumber(GetConditionFreeFormatFlag)
    colListLine.Add GetConditionRuleRef
    colListLine.Add GetBooleanFromNumber(GetConditionDeleteFlag)
    
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

Public Function GetConditionReference() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONDITIONREFERENCE")
    GetConditionReference = CStr(vVal)
End Function

Public Function GetConditionName() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONDITIONNAME")
    GetConditionName = CStr(vVal)
End Function

Public Function GetConditionDescription() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONDITIONDESCRIPTION")
    GetConditionDescription = CStr(vVal)
End Function

Public Function GetConditionType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONDITIONTYPE")
    GetConditionType = CStr(vVal)
End Function

Public Function GetConditionChannelID() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CHANNELID")
    GetConditionChannelID = vVal
End Function

Public Function GetConditionTypeText() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TYPETEXT")
    GetConditionTypeText = CStr(vVal)
End Function

Public Function GetConditionEditFlag() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EDITABLEIND")
    GetConditionEditFlag = CStr(vVal)
End Function

Public Function GetConditionFreeFormatFlag() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FREEFORMATIND")
    GetConditionFreeFormatFlag = CStr(vVal)
End Function

Public Function GetConditionalIndicator() As Integer
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CONDITIONIND")
    GetConditionalIndicator = CInt(vVal)
End Function

Public Function GetConditionRuleRef() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("RULEREFERENCE")
    GetConditionRuleRef = CStr(vVal)
End Function

Public Function GetConditionDeleteFlag() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("DELETEFLAG")
    GetConditionDeleteFlag = CStr(vVal)
End Function

Public Function GetComboFields() As Collection
    Dim colFields As New Collection
    colFields.Add "CHANNELID"
    colFields.Add "CHANNELNAME"
    Set GetComboFields = colFields
End Function

Friend Function SetConditionReferenceID(sReference As String) As Boolean
    SetConditionReferenceID = m_clsTableAccess.SetVal("CONDITIONREFERENCE", CVar(sReference))
End Function

Friend Function SetConditionName(sName As String) As Boolean
    SetConditionName = m_clsTableAccess.SetVal("CONDITIONNAME", sName)
End Function

Friend Function SetConditionDescription(sDescription As String) As Boolean
    SetConditionDescription = m_clsTableAccess.SetVal("CONDITIONDESCRIPTION", sDescription)
End Function

Friend Function SetConditionType(nType As Integer) As Boolean
    SetConditionType = m_clsTableAccess.SetVal("CONDITIONTYPE", CVar(nType))
End Function

Friend Function SetConditionEditInd(nEditInd As Integer) As Boolean
    SetConditionEditInd = m_clsTableAccess.SetVal("EDITABLEIND", CVar(nEditInd))
End Function

Friend Function SetConditionFormatInd(nFormatInd As Integer) As Boolean
    SetConditionFormatInd = m_clsTableAccess.SetVal("FREEFORMATIND", CVar(nFormatInd))
End Function

Friend Function SetConditionChannelID(sChannelID As String) As Boolean
    SetConditionChannelID = m_clsTableAccess.SetVal("CHANNELID", sChannelID)
End Function

Friend Function SetConditionRuleReference(sRuleReference As String) As Boolean
    SetConditionRuleReference = m_clsTableAccess.SetVal("RULEREFERENCE", sRuleReference)
End Function

Friend Function SetConditionDeleteFlag(nDeleteFlag As Integer) As Boolean
    SetConditionDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDeleteFlag))
End Function

Friend Function SetConditionalIndicator(nConditionalInd As Integer) As Boolean
    SetConditionalIndicator = m_clsTableAccess.SetVal("CONDITIONIND", CVar(nConditionalInd))
End Function

Public Function GetNextConditionRef() As String
    On Error GoTo Failed
    Dim cmd As ADODB.Command
    Dim clsDataAccess As DataAccess
    Dim Param As ADODB.Parameter
    Dim colResult As Collection
    Dim colParam As Collection
    
    Set Param = New ADODB.Parameter
    Set colParam = New Collection
    Set colResult = New Collection

    Set cmd = g_clsDataAccess.GetActiveCommand()

    With Param
        .Name = "NextNumber"
        .Type = adNumeric
        .Direction = adParamOutput
        .Precision = 10
    End With
    
    colParam.Add Param
    
    ' DJP SQL Server port, remove hard coded proc name
    Set colResult = g_clsDataAccess.ExecuteStoredProcedure(spGetNextConditionsRef, colParam)
    
    GetNextConditionRef = colResult(1)
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
        
End Function

Public Sub GetConditions(Optional sRef As String, Optional sName As String, Optional sType As String, Optional sChannelID As String)
    Dim sSearch As String
    On Error GoTo Failed
    
    ' SQL Server port, use global SQLAssist not local one
    sSearch = "    SELECT " & _
                    "C.*, " & _
                    "CV.VALUENAME TypeText " & _
           "    FROM " & _
                    "COMBOVALUE CV, " & _
                    "CONDITIONS C " & _
           " WHERE" & _
           "    CV.GROUPNAME = " & g_clsSQLAssistSP.FormatString(m_sComboGroup) & _
           "    AND CONDITIONTYPE = CV.VALUEID AND DELETEFLAG <> 1" ' TW 02/01/2007 EP2_640
    
    ' DJP BM0282 Add search conditions
    If Len(sRef) > 0 Then
        sSearch = sSearch & "   AND UPPER(C.CONDITIONREFERENCE) " & g_clsSQLAssistSP.GetLikeOperator & UCase(sRef)
    End If

    If Len(sName) > 0 Then
        sSearch = sSearch & "   AND UPPER(C.CONDITIONNAME) " & g_clsSQLAssistSP.GetLikeOperator & UCase(sName)
    End If
    
    If Len(sType) > 0 Then
        sSearch = sSearch & "   AND UPPER(C.CONDITIONTYPE) = " & UCase(sType)
    End If
    
    If Len(sChannelID) > 0 Then
        sSearch = sSearch & "   AND UPPER(C.CHANNELID) = " & UCase(sChannelID)
    End If
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub DeleteCondition()
    On Error GoTo Failed
    
    GetCurrentRecord
' TW 02/01/2007 EP2_640
    If m_clsTableAccess.RecordCount > 0 Then
' TW 02/01/2007 EP2_640 End
        SetConditionDeleteFlag 1
        m_clsTableAccess.Update
' TW 02/01/2007 EP2_640
    End If
' TW 02/01/2007 EP2_640 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetCurrentRecord()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    
    Set rs = m_clsTableAccess.GetTableData(POPULATE_KEYS)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

