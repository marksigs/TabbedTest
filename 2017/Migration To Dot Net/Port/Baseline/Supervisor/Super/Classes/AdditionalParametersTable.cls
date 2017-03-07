VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdditionalParametersTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "MortgageLenderParametersTable"
Private m_sSearch As String
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "ContactDetailsGUID"
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOrganisationID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("OrganisationID")
    GetOrganisationID = CStr(vVal)
End Function
Friend Function GetACCInterestInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ContactForname")
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetContactDetailsGUID(vVal As Variant) As Boolean
   SetContactDetailsGUID = m_clsTableAccess.SetVal("ContactDetailsGUID", vVal)
End Function
Friend Function SetContactForname(vVal As Variant) As Boolean
   SetContactForname = m_clsTableAccess.SetVal("ContactForname", vVal)
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
   SetTelExtNo = m_clsTableAccess.SetVal("TelExtNo", vVal)
End Function
Friend Function SetTelNo(vVal As Variant) As Boolean
   SetTelNo = m_clsTableAccess.SetVal("TelNo", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_AddRow() As Boolean
    TableAccess_AddRow = m_clsTableAccess.AddRow()
End Function
Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub
Private Function TableAccess_DeleteRow(colMatchData As Collection) As Boolean
    TableAccess_DeleteRow = m_clsTableAccess.DeleteRow(colMatchData)
End Function
Private Function TableAccess_DoesRecordExist(colMatchData As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData)
End Function

Private Function TableAccess_GetDate(sField As String) As Variant
    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function

Private Function TableAccess_GetKeyMatchFields()
    TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function

Private Function TableAccess_GetKeyMatchValues() As Variant

End Function

Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function

'Private Function TableAccess_DoUpdateSets() As Boolean
'    TableAccess_DoUpdateSets = m_clsTableAccess.DoUpdateSets()
'End Function
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS) As ADODB.Recordset
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function

Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function
Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function
Private Function TableAccess_IsDataValid() As Boolean
    TableAccess_IsDataValid = m_clsTableAccess.IsDataValid()
End Function
Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function
Private Function TableAccess_MoveFirst() As Boolean
    TableAccess_MoveFirst = m_clsTableAccess.MoveFirst()
End Function
Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
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
Private Function TableAccess_Update() As Boolean
    TableAccess_Update = m_clsTableAccess.Update()
End Function
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

Public Sub TableAccess_Create(sTable As String)
    m_clsTableAccess.Create (sTable)
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub



Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End FunctionPrivate Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    TableAccess_ApplyFilter = m_clsTableAccess.ApplyFilter(sSearch)
End Sub

Private Sub TableAccess_CancelFilter()
    TableAccess_CancelFilter = m_clsTableAccess.CancelFilter
End Sub