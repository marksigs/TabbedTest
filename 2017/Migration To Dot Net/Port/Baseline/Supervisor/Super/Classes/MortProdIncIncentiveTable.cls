VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncIncentiveTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncIncentiveTable
' Description   : Manages access to the Inclusive and Exclusive Incentives tables
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       03/12/01    SYS2912 SQL Server locking problem, update GetIncentives and make this
'                       a Banded table class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Implements BandedTable

' Private data
Private m_sSearch As String
Private m_sTableName As String

Private m_clsTableAccess As TableAccess
Private m_clsBandedTable As BandedTable
Private m_IncentiveType As IncentiveType

' Constants
Private Const m_sVersionField As String = "INCENTIVEVERSIONNUMBER"

' Enums
Public Enum IncentiveType
    Inclusive
    Exclusive
    InclusiveAndExclusive
End Enum
Friend Sub SetType(Incentive As IncentiveType)
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim colFields As New Collection
    Dim colDeleteFields As New Collection
    
    m_IncentiveType = Incentive
    
    If Incentive = Exclusive Then
        m_sTableName = "ExclusiveIncentive"
    ElseIf Incentive = Inclusive Then
        m_sTableName = "InclusiveIncentive"
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Invalid incentive type " & Incentive
    End If
    
    colDeleteFields.Add "IncentiveGUID"
    colFields.Add "MortgageProductCode"
    colFields.Add "Startdate"
    
    ' If versioning exists, we'll need to add that too...
    If g_clsVersion.DoesVersioningExist() Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields
    m_clsTableAccess.SetDeleteFields colDeleteFields
    m_clsTableAccess.Create m_sTableName

    m_clsBandedTable.SetUpdateFields colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DJP SQL Server port, return variant
Friend Function GetIncentiveGUID() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("IncentiveGUID")

    ' DJP SQL Server port
    GetIncentiveGUID = vVal
End Function
Friend Function GetMortgageProductCode() As String
    GetMortgageProductCode = m_clsTableAccess.GetVal("MortgageProductCode")
End Function
Friend Function GetStartDate() As String
    Dim vVal As Variant
        
    vVal = m_clsTableAccess.GetVal("StartDate")
        
    If Not IsNull(vVal) Then
        GetStartDate = Format(vVal, "Short Date")
    End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIncentiveGUID() As Variant
    Dim clsGUID As New GuidAssist
    Dim vVal As Variant
    Dim bRet As Boolean
    
    vVal = CVar(clsGUID.CreateGUID())
    
    bRet = m_clsTableAccess.SetVal("IncentiveGUID", vVal)

    If bRet = True Then
        SetIncentiveGUID = vVal
    Else
        SetIncentiveGUID = ""
    End If


End Function
Friend Function SetMortgageProductCode(sCode As String) As Boolean
    SetMortgageProductCode = m_clsTableAccess.SetVal("MortgageProductCode", CVar(sCode))
    
    ' Versioning, if it exists
    g_clsVersion.HandleVersioning Me

End Function
Friend Function SetStartDate(sStartDate As Variant) As Boolean
    SetStartDate = m_clsTableAccess.SetVal("StartDate", CVar(sStartDate))
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsTableAccess = New TableAccess
    Set m_clsBandedTable = New BandedTable
    m_clsBandedTable.SetTableAccess m_clsTableAccess
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_clsTableAccess.SetOrderCriteria sCriteria
End Sub

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
Public Sub GetIncentives(incType As IncentiveType, colIncentiveGUIDs As Collection)
    On Error GoTo Failed
    Dim sSearch As String
    Dim sDatabase As String
    Dim sSubSearch As String
    Dim sIncentiveTable As String
    Dim clsTableAccess As TableAccess
    
    If incType = Exclusive Then
        sIncentiveTable = "ExclusiveIncentive"
    ElseIf incType = Inclusive Then
        sIncentiveTable = "InclusiveIncentive"
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Invalid Incentive Type: " & incType
    End If
        
    SetType incType

    sDatabase = m_clsTableAccess.GetDatabase
        
    sSearch = "SELECT * FROM " & _
                    sIncentiveTable & " " & _
                "WHERE " & _
                    "INCENTIVEGUID IN "
    
    Dim vIncentiveGUID As Variant
    Dim sIncentiveGUID As String
    Dim nThisGUID As Long
    Dim nCount As Long
    
    sSubSearch = "("
    
    nCount = colIncentiveGUIDs.Count
    nThisGUID = 1
    
    For Each vIncentiveGUID In colIncentiveGUIDs
        sIncentiveGUID = g_clsSQLAssistSP.GuidToString(CStr(vIncentiveGUID))
        sIncentiveGUID = g_clsSQLAssistSP.FormatGuid(sIncentiveGUID)
        
        sSubSearch = sSubSearch & sIncentiveGUID
        If nThisGUID < nCount Then
            sSubSearch = sSubSearch & ", "
        End If
        nThisGUID = nThisGUID + 1
    Next
    
    sSubSearch = sSubSearch & ")"
    sSearch = sSearch & sSubSearch
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub BandedTable_DoUpdateSets()
    m_clsBandedTable.DoUpdateSets
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
