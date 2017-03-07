VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdProdCondTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdEmpEligTable
' Description   : Contains all data access methods and SQL queries for the MORTGAGEPRODUCTCONDITIONS
'                 table.
' BMIDS Change history
' Prog      Date        Description
' GD       21/05/02     BMIDS Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
' Constants
Private Const m_sTableName = "MORTGAGEPRODUCTCONDITION"
'Private Const m_sVersionField As String = "PRODUCTVERSIONNUMBER"
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName

    colFields.Add "MortgageProductCode"
    colFields.Add "StartDate"
    'colFields.Add "ConditionReference"
    
    ' If versioning exists, we'll need to add that too...
    'If g_clsVersion.DoesVersioningExist() Then
        'colFields.Add m_sVersionField
    'End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetProductCode() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MortgageProductCode")
    GetProductCode = CStr(vVal)
End Function
Friend Function GetStartDate() As String
    Dim vVal As Variant
        
    vVal = m_clsTableAccess.GetVal("StartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function

Friend Function GetConditionReference() As String
    GetConditionReference = CStr(m_clsTableAccess.GetVal("ConditionReference"))
End Function
Friend Function GetConditionName(sCondRef As String) As String
'GD BMIDS. Gets Conditions.ConditionName where ConditionReference = DATA PASSED IN
    On Error GoTo Failed
    Dim sSearch1 As String
    Dim sSearch2 As String

    Dim rsConditions As ADODB.Recordset
    If Len(sCondRef) > 0 Then 'Check that empty string isn't passed in
        sSearch1 = "SELECT CONDITIONNAME FROM CONDITIONS WHERE DELETEFLAG <> 1   AND CONDITIONREFERENCE = " & sCondRef & " AND CONDITIONTYPE ="
        sSearch2 = "(SELECT VALUEID FROM COMBOVALIDATION WHERE GROUPNAME = 'ConditionType' AND VALIDATIONTYPE = 'P')"
        
        Set rsConditions = g_clsDataAccess.GetTableData("CONDITIONS", sSearch1 & sSearch2)
        If rsConditions.RecordCount > 0 Then
            rsConditions.MoveFirst 'Should only ever have 1 record returned
            GetConditionName = rsConditions.fields(0).Value
        Else
            GetConditionName = ""
        End If
    Else
        GetConditionName = ""
    End If
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
  
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllConditions
' Description   : Returns the list of ALL CONDITIONS with validation type 'P', which have no Delete Flag set
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAllConditions()
'GD BMIDS0012 Get ALL of CONDITIONS table where DELETEFLAG not set and ConditionType has validation type of 'P'
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSearch1 As String
    Dim sSearch2 As String
    
    Dim colFields As Collection
    Dim rsConditions As ADODB.Recordset

    'SQL to get data from CONDITIONS table
    sSearch1 = "SELECT CONDITIONREFERENCE,CONDITIONNAME FROM CONDITIONS WHERE DELETEFLAG <>1 AND CONDITIONTYPE = "
    
    'SQL to get the valueid from for ConditionType with ValidationType of 'P'
    sSearch2 = "(SELECT VALUEID FROM COMBOVALIDATION WHERE GROUPNAME = 'ConditionType' AND VALIDATIONTYPE = 'P')"
    sSearch = sSearch1 & sSearch2

    Set rsConditions = g_clsDataAccess.GetTableData("CONDITIONS", sSearch)
    
    m_clsTableAccess.SetRecordSet rsConditions
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
Public Sub GetMortgageProductCondition(sMortgageProductCode, sStartDate)
'GD BMIDS0012 Get All MORTGAGEPRODUCTCONDITIONs for DATA PASSED IN
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSearch1 As String
    Dim sSearch2 As String
    
    'Dim sProdSearch As String
    Dim colFields As Collection
    Dim rsConditions As ADODB.Recordset

    
    sSearch = "SELECT * FROM MORTGAGEPRODUCTCONDITION WHERE MORTGAGEPRODUCTCODE = '" & sMortgageProductCode & "' AND STARTDATE = '" & sStartDate & "'"

    Set rsConditions = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTCONDITION", sSearch)
    
    m_clsTableAccess.SetRecordSet rsConditions
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetProductCode(sProductCode As String) As Boolean
    SetProductCode = m_clsTableAccess.SetVal("MortgageProductCode", CVar(sProductCode))

    ' Versioning, if it exists
    'g_clsVersion.HandleVersioning Me

End Function
Friend Function SetStartDate(sStartDate As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("StartDate", CVar(sStartDate))
End Function
Friend Function SetConditionReference(sStatus As String) As Boolean
    SetConditionReference = m_clsTableAccess.SetVal("ConditionReference", CVar(sStatus))
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
    'colListLine.Add GetStatusText()
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
Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
