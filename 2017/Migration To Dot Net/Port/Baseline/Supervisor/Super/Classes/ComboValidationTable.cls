VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboValidationTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ComboValidation
' Description   : Contains accesss routines for the ComboValidation table.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Implements BandedTable

' Constants
Private Const m_sTableName = "COMBOVALIDATION"

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private Sub Class_Initialize()
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "GroupName"
    colFields.Add "ValueID"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "GroupName"
    colUpdateFields.Add "ValueID"

    Set colDupKeys = New Collection
    
    colDupKeys.Add "GROUPNAME"
    colDupKeys.Add "VALUEID"
    colDupKeys.Add "VALIDATIONTYPE"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys


    m_clsBandedTable.SetUpdateFields colUpdateFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetGroupName() As String
    GetGroupName = CStr(m_clsTableAccess.GetVal("GroupName"))
End Function
Friend Function GetValueID() As String
    GetValueID = CStr(m_clsTableAccess.GetVal("ValueID"))
End Function
Friend Function GetValueName() As String
    GetValueName = CStr(m_clsTableAccess.GetVal("VALUENAME"))
End Function
Friend Function GetValidationType() As String
    GetValidationType = CStr(m_clsTableAccess.GetVal("ValidationType"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetGroupName(sGroupName As String) As Boolean
    SetGroupName = m_clsTableAccess.SetVal("GroupName", CVar(sGroupName))
End Function
Friend Function SetValueID(sValueId As String) As Boolean
    SetValueID = m_clsTableAccess.SetVal("ValueID", CVar(sValueId))
End Function
Friend Function SetValidationType(sValidationType As String) As Boolean
    SetValidationType = m_clsTableAccess.SetVal("ValidationType", CVar(sValidationType))
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
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Private Sub BandedTable_SetTableAccess(clsTableAccess As TableAccess)
    m_clsBandedTable.SetTableAccess clsTableAccess
End Sub
Private Sub BandedTable_SetUpdateFields(colFields As Collection)
    m_clsBandedTable.SetUpdateFields colFields
End Sub
Private Sub BandedTable_SetUpdateSets(Optional bUpdate As Boolean = True)
    m_clsBandedTable.SetUpdateSets bUpdate
End Sub
Private Sub BandedTable_SetUpdateValues(colValues As Collection)
    m_clsBandedTable.SetUpdateValues colValues
End Sub
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub DeleteValidation(Optional sTargetDatabase As String)
    On Error GoTo Failed
    Dim sCommand As String
    Dim colMatchValues As Collection
    Dim sGroup As String
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues()
    
    If colMatchValues.Count = 1 Then
        ' First value should be the group name
        sGroup = colMatchValues(1)

        sCommand = "DELETE FROM " + m_sTableName + " where GROUPNAME = '" + sGroup + "'"
        g_clsDataAccess.ExecuteCommand sCommand, sTargetDatabase
    
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "Delete Validation records"
    End If

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
Public Sub GetValueFromValidation(sGroup As String, sValidation As String)
    On Error GoTo Failed
    Dim sSearch As String
    sGroup = UCase(g_clsSQLAssistSP.FormatString(sGroup))
    sValidation = UCase(g_clsSQLAssistSP.FormatString(sValidation))
    
    sSearch = "SELECT COMBOVALUE.*, COMBOVALIDATION.VALIDATIONTYPE FROM COMBOVALIDATION, COMBOVALUE WHERE UPPER(COMBOVALIDATION.GROUPNAME )= " + sGroup + " AND " + _
    " COMBOVALIDATION.VALIDATIONTYPE = " + sValidation + " AND UPPER(COMBOVALUE.GROUPNAME) = " + sGroup + " AND " + _
    " COMBOVALIDATION.VALUEID = COMBOVALUE.VALUEID"
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetComboGroup
' Description   :   Returns a Recordset containing combogroup records, including validation types
'                   sGroup Contains the name of the group to use
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetComboGroupWithValidationType(sGroup As String) As ADODB.Recordset
    
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim Param As ADODB.Parameter
    Dim colParam As Collection
    
    ' First parameter, the GroupName
    Set Param = New ADODB.Parameter
    Set colParam = New Collection
    
    With Param
        .Name = "GroupName"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = sGroup
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With
    
    colParam.Add Param
        
    ' DJP SQL Server port
    ' Get the data
    Set rs = g_clsDataAccess.ExecuteStoredProcRS(spGetComboWithValidation, colParam)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            m_clsTableAccess.SetRecordSet rs
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Combo Group " & sGroup
        End If
    End If
    
    Set GetComboGroupWithValidationType = rs
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValidationTypeFromCombo
' Description   : Moves the record set nMove places, and returns the validation type for that record
'                 Used in DBCombo for frmEditBatch
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetValidationTypeFromCombo(nMove As Long) As String
    On Error GoTo Failed
    Dim sValidationType As String
    Dim rs As ADODB.Recordset
    
    
    Set rs = m_clsTableAccess.GetRecordSet
    rs.Move (nMove - 1), 1
        
    sValidationType = GetValidationType
    GetValidationTypeFromCombo = sValidationType
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValueIDForValidationType
' Description   : Given a Validation Type and a ComboGroup, returns the valueID for that Record
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetValueIDForValidationType(sValidationType As String, sGroup As String) As Long
    On Error GoTo Failed
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim vValueID As Variant
    Dim sTableName As String
    
    'JD MAR40 initialise vValueID so if one is not found there is not an error
    vValueID = 0
    sTableName = "ComboValidation"
    sSQL = "SELECT ValueID FROM " & sTableName & " WHERE GroupName = " & g_clsSQLAssistSP.FormatString(sGroup) & _
           " AND ValidationType = " & g_clsSQLAssistSP.FormatString(sValidationType)
           
    Set rs = g_clsDataAccess.GetTableData(sTableName, sSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            vValueID = rs.fields("ValueID")
        End If
    End If
    
    GetValueIDForValidationType = vValueID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetSingleValueFromValidation(sGroup As String, sValidation As String) As Long
    On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim nValueID As Long
    
    sGroup = g_clsSQLAssistSP.FormatString(sGroup)
    
    sValidation = g_clsSQLAssistSP.FormatString(sValidation)
    
    sSearch = "SELECT VALUEID FROM " & m_sTableName & " WHERE ValidationType = " & _
              sValidation & " AND GroupName = " & sGroup
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    nValueID = 0
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            nValueID = rs.fields("VALUEID")
        End If
    End If
    
    GetSingleValueFromValidation = nValueID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

Public Function GetComboValidationAsCollection(sGroup As String, sValueId As String) As Collection

    On Error GoTo Failed
    Dim col As Collection
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim nCnt As Integer
    Dim sValidationType As String
    'Set rs = adodb.Recordset
    
    Set col = New Collection
        
    sSQL = "Select ValidationType from " & m_sTableName & " Where GroupName = " & g_clsSQLAssistSP.FormatString(sGroup) & _
           " And ValueID = " & sValueId
           
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            For nCnt = 1 To rs.RecordCount
                sValidationType = rs.fields("ValidationType")
                col.Add sValidationType
                rs.MoveNext
            Next
            Set GetComboValidationAsCollection = col
        End If
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValueIDFromValidationType
' Description   : Returns all records matching the Valueid in the specified group
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetValidationTypeForValueID(sGroup As String, sValueId As String)
    On Error GoTo Failed
    
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select " & _
                "* " & _
           "From " & _
                m_sTableName & " " & _
           "Where " & _
                "GroupName = " & g_clsSQLAssistSP.FormatString(sGroup) & " And " & _
                "ValueID = " & g_clsSQLAssistSP.FormatString(sValueId)
                
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
