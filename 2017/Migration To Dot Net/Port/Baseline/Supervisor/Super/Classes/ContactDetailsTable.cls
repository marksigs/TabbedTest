VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactDetailsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ContactDetailsTable
' Description   : Contails all SQL and data access routines for the ContactDetails table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' STB       10/12/01    SYS2550 - Generate GUIDs work if variants are empty as well as missing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM History
' EP1033  HMA  26/07/2006  Add ContactOtherTitle
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_bContactDetailsExists As Boolean

' Constants
Private Const m_sTableName = "CONTACTDETAILS"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "ContactDetailsGUID"
    m_clsTableAccess.SetKeyMatchFields colFields

    m_bContactDetailsExists = g_clsDataAccess.DoesTableExist(m_clsTableAccess.GetTable())
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DJP SQL Server port. Return variant, not string
Friend Function GetContactDetailsGUID() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactDetailsGUID")
    ' DJP SQL Server port
    GetContactDetailsGUID = vVal
End Function
Friend Function GetContactForname() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactForename")
    GetContactForname = CStr(vVal)
End Function
Friend Function GetContactSurname() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactSurname")
    GetContactSurname = CStr(vVal)
End Function
Friend Function GetContactTitle() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactTitle")
    GetContactTitle = CStr(vVal)
End Function
Friend Function GetContactType() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactType")
    GetContactType = CStr(vVal)
End Function
Friend Function GetEmailAddress() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("EmailAddress")
    GetEmailAddress = CStr(vVal)
End Function
Friend Function GetFaxNumber() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FaxNumber")
    GetFaxNumber = CStr(vVal)
End Function
Friend Function GetTelExtNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TelephoneExtensionNumber")
    GetTelExtNo = CStr(vVal)
End Function
Friend Function GetTelNo() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("TelephoneNumber")
    GetTelNo = CStr(vVal)
End Function
'EP1033 Get Other Title
Friend Function GetOtherTitle() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactTitleOther")
    GetOtherTitle = CStr(vVal)
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetContactDetailsGUID(Optional vVal As Variant) As Variant
    On Error GoTo Failed
    Dim bRet As Boolean
    
    If IsMissing(vVal) Or (TypeName(vVal) = "Empty") Then
        Dim clsGUID As New GuidAssist
        vVal = CVar(clsGUID.CreateGUID())
    End If
    
    m_clsTableAccess.SetVal "ContactDetailsGUID", vVal
    
    SetContactDetailsGUID = vVal
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Friend Function SetContactForname(vVal As Variant) As Boolean
   SetContactForname = m_clsTableAccess.SetVal("ContactForename", vVal)
End Function
Friend Function SetContactSurname(vVal As Variant) As Boolean
   SetContactSurname = m_clsTableAccess.SetVal("ContactSurname", vVal)
End Function
Friend Function SetContactTitle(vVal As Variant) As Boolean
   SetContactTitle = m_clsTableAccess.SetVal("ContactTitle", vVal)
End Function
Friend Function SetContactType(vVal As Variant) As Boolean
   SetContactType = m_clsTableAccess.SetVal("ContactType", vVal)
End Function
Friend Function SetEmailAddress(vVal As Variant) As Boolean
   SetEmailAddress = m_clsTableAccess.SetVal("EmailAddress", vVal)
End Function
Friend Function SetFaxNumber(vVal As Variant) As Boolean
   SetFaxNumber = m_clsTableAccess.SetVal("FaxNumber", vVal)
End Function
Friend Function SetTelExtNo(vVal As Variant) As Boolean
   SetTelExtNo = m_clsTableAccess.SetVal("TelephoneExtensionNumber", vVal)
End Function
Friend Function SetTelNo(vVal As Variant) As Boolean
   SetTelNo = m_clsTableAccess.SetVal("TelephoneNumber", vVal)
End Function
'EP1033 Set Other Title
Friend Function SetOtherTitle(vVal As Variant) As Boolean
   SetOtherTitle = m_clsTableAccess.SetVal("ContactTitleOther", vVal)
End Function
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
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

Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub

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

'Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, optional bCheck as boolean = true ) As Variant
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
    If m_bContactDetailsExists Then
        m_clsTableAccess.Update
    End If
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

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
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
