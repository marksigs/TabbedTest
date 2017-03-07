VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MPMigRateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MigRateTable
' Description   : Contains all access routines and SQL for the MigRate table
'
' Change history
' Prog      Date        Description
' AW        21/05/02    BM087 - Created
' AW        28/05/02    BMIDS00016 - Add target database parameter to DoesProductExistForThisItem()
' JD        30/03/2005  BMIDS982 - changed message text from 'mig' to 'HLC'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess
Implements BandedTable

' Constants
Private Const m_sTableName = "MPMIGRATEBAND"
Private Const m_sVersionField As String = "MPMIGSETVERSIONNUMBER"

' Private data
Private m_sSearch As String
Private m_sFirstMatch As String
Private m_bVersionExists As Boolean
Private m_clsBandedTable As BandedTable
Private m_clsTableAccess As TableAccess
Private m_sAdminFeeStartDateField As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetMPMigRateSet() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MPMigRateSet")
    GetMPMigRateSet = CStr(vVal)
End Function
Friend Function GetEffectiveDate() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EffectiveDate")
    GetEffectiveDate = CStr(vVal)
End Function
Friend Function GetEndDate() As String
    GetEndDate = CStr(m_clsTableAccess.GetVal("EndDate"))
End Function
Friend Function GetLowerLTV() As String
    GetLowerLTV = CStr(m_clsTableAccess.GetVal("LowerLTV"))
End Function
Friend Function GetHigherLTV() As String
    GetHigherLTV = CStr(m_clsTableAccess.GetVal("HigherLTV"))
End Function
Friend Function GetLowerLoanAmount() As String
    GetLowerLoanAmount = CStr(m_clsTableAccess.GetVal("LowerLoanAmount"))
End Function
Friend Function GetHigherLoanAmount() As String
    GetHigherLoanAmount = CStr(m_clsTableAccess.GetVal("HigherLoanAmount"))
End Function
Friend Function GetMinPrem() As String
    GetMinPrem = CStr(m_clsTableAccess.GetVal("MinimumPremium"))
End Function
Friend Function GetIndemnityRate() As String
    GetIndemnityRate = CStr(m_clsTableAccess.GetVal("IndemnityRate"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetFeeSet(sRateSet As String) As Boolean
    SetFeeSet = m_clsTableAccess.SetVal("MPMigRateSet", CVar(sRateSet))
    
    g_clsVersion.HandleVersioning Me
End Function
Friend Function SetEffectiveDate(sEffectiveDate As String) As Boolean
    SetEffectiveDate = m_clsTableAccess.SetVal("EffectiveDate", CVar(sEffectiveDate))
End Function
Friend Function SetEndDate(sEndDate As String) As Boolean
    SetEndDate = m_clsTableAccess.SetVal("TypeOfApplication", CVar(sEndDate))
End Function
Friend Function SetLowerLTV(sLowerLTV As String) As Boolean
    SetLowerLTV = m_clsTableAccess.SetVal("TypeOfApplicationText", CVar(sLowerLTV))
End Function
Friend Function SetHigherLTV(sHigherLTV As String) As Boolean
    SetHigherLTV = m_clsTableAccess.SetVal("TypeOfApplicationText", CVar(sHigherLTV))
End Function
Friend Function SetLowerLoanAmount(sLowerLoanAmount As String) As Boolean
    SetLowerLoanAmount = m_clsTableAccess.SetVal("Location", CVar(sLowerLoanAmount))
End Function
Friend Function SetHigherLoanAmount(sHigherLoanAmount As String) As Boolean
    SetHigherLoanAmount = m_clsTableAccess.SetVal("Location", CVar(sHigherLoanAmount))
End Function
Friend Function SetMinPrem(sSetMinPrem As String) As Boolean
    SetMinPrem = m_clsTableAccess.SetVal("Location", CVar(sSetMinPrem))
End Function
Friend Function SetIndemRate(sIndemRate As String) As Boolean
    SetIndemRate = m_clsTableAccess.SetVal("Location", CVar(sIndemRate))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Banded functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetVersionField() As String
    GetVersionField = m_sVersionField
End Function
Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim colFields As New Collection
    Dim colUpdateFields As New Collection
    Dim colDupKeys As Collection
    Dim bVersionExists As Boolean
    
    m_bVersionExists = g_clsDataAccess.DoesFieldExist(m_sTableName, m_sVersionField, , False)
    
'    m_sFirstMatch = "select * from MPMigRateBand A where A.HIGHERLTV = " + _
'    "(select min(HIGHERLTV) from  MPMigRateBand B " + _
'        "where A.MPMIGRATESET = B.MPMIGRATESET and A.EffectiveDate = B.EffectiveDate) and A.HIGHERLOANAMOUNT = " + _
'    "(select min(HIGHERLOANAMOUNT) from MPMigRateBand C " + _
'        "where a.EffectiveDate = c.EffectiveDate and a.MPMIGRATESET = C.MPMIGRATESET and " + _
'        "A.HIGHERLTV = C.HIGHERLTV)"
        
    If m_bVersionExists Then
        m_sFirstMatch = m_sFirstMatch + " AND " + m_sVersionField + " = " & g_sVersionNumber
    End If
        
        
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
        
    m_clsTableAccess.Create m_sTableName ', m_sFirstMatch

    colFields.Add "MPMIGRATESET"
    'colFields.Add "EffectiveDate"
    
    If m_bVersionExists Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields

    m_clsBandedTable.SetTableAccess m_clsTableAccess
    
    colUpdateFields.Add "MPMIGRATESET"
    colUpdateFields.Add "EffectiveDate"

    m_clsBandedTable.SetUpdateFields colUpdateFields

    Set colDupKeys = New Collection
    colDupKeys.Add "EffectiveDate"
    colDupKeys.Add "HIGHERLTV"
    colDupKeys.Add "HIGHERLOANAMOUNT"
    
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetPopulateNull()
    m_sSearch = "Select * from " + m_sTableName + " where Name is null"
End Sub
Public Sub SetPopulateType(enumType As PopulateType)
    m_clsTableAccess.SetPopulateType enumType
End Sub
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

Private Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
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
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
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

Private Function TableAccess_GetUpdateValues() As Collection
    Set TableAccess_GetUpdateValues = m_clsTableAccess.GetUpdateValues
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
    m_clsTableAccess.SetPopulateType enumType
End Sub


Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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
Public Function TableAccess_GetUpdateFields() As Collection
    Dim colKeys As New Collection
    
    colKeys.Add "AdministrationFeeSet"
    colKeys.Add m_sAdminFeeStartDateField '"AdministrationFeeStartDate"

    Set TableAccess_GetUpdateFields = colKeys
End Function
Public Sub GetHeaderData(colHeader As Collection)
    colHeader.Add GetMPMigRateSet()
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    If bIncludeHeader = True Then
        GetHeaderData colListLine
    End If
    
'    colListLine.Add GetTypeOfApplicationText()
'    colListLine.Add GetLocationText()
'    colListLine.Add GetAmount()

End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    'Modified by Sunil 07-03-2002
    'Modified by Sunil 06-02-2002
    'In order to check whether mortgage products exist for the selected
    'base rate
    On Error GoTo Failed
    Dim sMessage As String
    Dim inum As Integer
    Dim colValues  As New Collection
    Set colValues = m_clsTableAccess.GetKeyMatchValues
    
    Dim clsMortgageProductTable As MortgageProductTable
    Set clsMortgageProductTable = New MortgageProductTable
    
    Set colValues = clsMortgageProductTable.DoesProductExistForThisItem("MPMIGRATESET", colValues.Item(1), sDatabase)
    
    If (Not colValues Is Nothing) And (colValues.Count > 0) Then
        sMessage = ""
        For inum = 1 To colValues.Count
            sMessage = sMessage & "(" & inum & ") " & colValues.Item(inum) & vbCrLf
        Next
        
        MsgBox "Record can not be deleted. The following Mortgage Products exist for the selected HLC rate set : " & vbCrLf & sMessage, vbCritical, "Cannot Delete"
        g_clsErrorHandling.RaiseError , "Record can not be deleted. The following Mortgage Products exist for the selected HLC rate set : " & vbCrLf & sMessage, ErrorSystem
    Else
        m_clsTableAccess.DeleteRecords sDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteMigRateBands()
    On Error GoTo Failed
    ' First, delete the admin fee values
    m_clsTableAccess.DeleteRecords
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

