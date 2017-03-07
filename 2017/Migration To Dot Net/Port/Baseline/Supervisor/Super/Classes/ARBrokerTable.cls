VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ARBrokerTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   ARBrokerTable
' Description   :   Contains all access functions for the ARBroker view.
' Change history
' Change history
' Prog     Date         Description
' PB       08/11/2006   EP2_15 - Created
' TW       16/01/2007   EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_objTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "ARBroker"
Private Const DELETE_FLAG_SET As Integer = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_objTableAccess = New TableAccess
    
    m_objTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "INTRODUCERID"
    m_objTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_objTableAccess.SetOrderCriteria "Surname"


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Friend Function GetIntroducerID() As String
    GetIntroducerID = CStr(m_objTableAccess.GetVal("IntroducerID"))
End Function
Friend Function GetTitle() As String
    GetTitle = CStr(m_objTableAccess.GetVal("TITLETEXT"))
End Function
Friend Function GetAddress() As String
    GetAddress = CStr(m_objTableAccess.GetVal("Address"))
End Function

Friend Function GetForename() As String
    GetForename = CStr(m_objTableAccess.GetVal("USERFORENAME"))
End Function
Friend Function GetInitials() As String
    GetInitials = CStr(m_objTableAccess.GetVal("INITIALS"))
End Function
Friend Function GetFSARef() As String
    GetFSARef = CStr(m_objTableAccess.GetVal("FSAARFirmRef"))
End Function
Friend Function GetUserName() As String
' TW 16/01/2007 EP2_859
    GetUserName = CStr(m_objTableAccess.GetVal("USERNAME"))
End Function

Friend Function GetSurname() As String
    GetSurname = CStr(m_objTableAccess.GetVal("USERSURNAME"))
End Function

Friend Function GetUserID() As String
    GetUserID = CStr(m_objTableAccess.GetVal("USERID"))
End Function
Friend Function GetEmailAddress() As String
    GetEmailAddress = CStr(m_objTableAccess.GetVal("EMAILADDRESS"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIntroducerID(sValue As Variant) As Boolean
    SetIntroducerID = m_objTableAccess.SetVal("INTRODUCERID", CVar(sValue))
End Function
Friend Function SetUserID(sValue As Variant) As Boolean
    SetUserID = m_objTableAccess.SetVal("USERID", CVar(sValue))
End Function
Friend Function SetTitle(sValue As Variant) As Boolean
    SetTitle = m_objTableAccess.SetVal("TITLE", CVar(sValue))
End Function
Friend Function SetForename(sValue As Variant) As Boolean
    SetForename = m_objTableAccess.SetVal("FORENAME", CVar(sValue))
End Function
Friend Function SetInitials(sValue As Variant) As Boolean
    SetInitials = m_objTableAccess.SetVal("INITIALS", CVar(sValue))
End Function
Friend Function SetSurname(sValue As Variant) As Boolean
    SetSurname = m_objTableAccess.SetVal("SURNAME", CVar(sValue))
End Function
Friend Function SetEmailAddress(sValue As Variant) As Boolean
    SetEmailAddress = m_objTableAccess.SetVal("EMAILADDRESS", CVar(sValue))
End Function

Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_objTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_objTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_objTableAccess.GetIsDeleted()
End Function

Private Sub TableAccess_SetDeleteFields(colFields As Collection)
    m_objTableAccess.SetDeleteFields colFields
End Sub

Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_objTableAccess.GetTableDataFromSearch sSearch
End Sub

Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_objTableAccess.GetLength(sField)
End Function
Private Sub TableAccess_AddRow()
    m_objTableAccess.AddRow
End Sub
Private Sub TableAccess_CloseRecordSet()
    m_objTableAccess.CloseRecordSet
End Sub
Private Sub TableAccess_DeleteAllRows()
    m_objTableAccess.DeleteAllRows
End Sub

Private Sub TableAccess_DeleteRow(colMatchData As Collection)
    m_objTableAccess.DeleteRow colMatchData
End Sub

Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    TableAccess_GetDate = m_objTableAccess.GetDate(sField)
End Function

Private Function TableAccess_GetKeyMatchFields()
    Set TableAccess_GetKeyMatchFields = m_objTableAccess.GetKeyMatchFields()
End Function

Private Function TableAccess_GetKeyMatchValues() As Variant
    Set TableAccess_GetKeyMatchValues = m_objTableAccess.GetKeyMatchValues()
End Function

Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_objTableAccess.GetRecordSet()
End Function
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Set TableAccess_GetTableData = m_objTableAccess.GetTableData(enumPopulateType, enumClassOption)
End Function
Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_objTableAccess.GetUpdated()
End Function
Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_objTableAccess.GetVal(sField)
End Function
Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
        m_objTableAccess.ValidateData sDesc
End Sub
Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_objTableAccess.MatchRecords(colFields, colMatchData)
End Function
Private Sub TableAccess_MoveFirst()
    m_objTableAccess.MoveFirst
End Sub
Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_objTableAccess.MoveNext()
End Function
Private Function TableAccess_RecordCount() As Integer
    TableAccess_RecordCount = m_objTableAccess.RecordCount()
End Function

Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean
    TableAccess_SetDate = m_objTableAccess.SetDate(sField, vVal)
End Function

Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)
    m_objTableAccess.SetKeyMatchFields colFields
End Sub

Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)
    m_objTableAccess.SetKeyMatchValues colValues
End Sub

Private Function TableAccess_SetNextRow() As Boolean
    TableAccess_SetNextRow = m_objTableAccess.SetNextRow()
End Function
Private Sub TableAccess_SetPopulateType(enumType As PopulateType)
    
End Sub
Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_objTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetUpdated()
    m_objTableAccess.SetUpdated
End Sub
Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
    TableAccess_SetVal = m_objTableAccess.SetVal(sField, vVal)
End Function
Private Sub TableAccess_Update()
    m_objTableAccess.Update
End Sub
Private Sub TableAccess_SetPopulateNull()
    m_objTableAccess.SetPopulateNull
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
    m_objTableAccess.Create sTable, sFirstMatch
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetIntroducerID
    colListLine.Add GetTitle
    colListLine.Add GetForename
    colListLine.Add GetInitials
    colListLine.Add GetSurname
    colListLine.Add GetEmailAddress
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_objTableAccess = clsTable
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_objTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_objTableAccess.SetDatabase sDatabaseKey
End Sub

Public Function TableAccess_GetDatabase() As String
    TableAccess_GetDatabase = m_objTableAccess.GetDatabase()
End Function

Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_objTableAccess.SetOrderCriteria sCriteria
End Sub

Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_objTableAccess.SetDuplicateKeys colDupKeys
End Sub
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_objTableAccess.GetDuplicateKeys()
End Function


Private Sub TableAccess_SetCollectionField(sField As String)
    m_objTableAccess.SetCollectionField sField
End Sub
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_objTableAccess.GetCollectionField()
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_objTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_objTableAccess.CancelFilter
End Sub

