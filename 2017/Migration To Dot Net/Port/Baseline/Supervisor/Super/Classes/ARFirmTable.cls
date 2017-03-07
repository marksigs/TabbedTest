VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ARFirmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   ARFirmTable
' Description   :   Contains all access functions for the ARFirm table.
' Change history
' Prog     Date         Description
' TW       17/10/2006   EP2_15 - Created
' TW       16/01/2007   EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "ARFIRM"
Private Const DELETE_FLAG_SET As Integer = 1

Private m_sFirstMatch As String

Friend Function GetUnitId() As String
    GetUnitId = CStr(m_clsTableAccess.GetVal("UnitId"))
End Function

Friend Function SetUnitID(sValue As Variant) As Boolean
    SetUnitID = m_clsTableAccess.SetVal("UnitId", CVar(sValue))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "ARFIRMID"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "ARFirmName"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetARFirmID() As String
    GetARFirmID = CStr(m_clsTableAccess.GetVal("ARFirmId"))
End Function
Friend Function GetFSAARFirmRef() As String
    GetFSAARFirmRef = CStr(m_clsTableAccess.GetVal("FSAARFirmRef"))
End Function
Friend Function GetPrincipalFirmID() As String
    GetPrincipalFirmID = CStr(m_clsTableAccess.GetVal("PrincipalFirmID"))
End Function
Friend Function GetARFirmName() As String
    GetARFirmName = CStr(m_clsTableAccess.GetVal("ARFirmName"))
End Function
Friend Function GetFirmStatus() As String
    GetFirmStatus = CStr(m_clsTableAccess.GetVal("FirmStatus"))
End Function
Friend Function GetListingStatus() As String
    GetListingStatus = CStr(m_clsTableAccess.GetVal("ListingStatus"))
End Function
Friend Function GetAddressLine1() As String
    GetAddressLine1 = CStr(m_clsTableAccess.GetVal("AddressLine1"))
End Function
Friend Function GetAddress() As String
    GetAddress = CStr(m_clsTableAccess.GetVal("Address"))
End Function

Friend Function GetAddressLine2() As String
    GetAddressLine2 = CStr(m_clsTableAccess.GetVal("AddressLine2"))
End Function
Friend Function GetAddressLine3() As String
    GetAddressLine3 = CStr(m_clsTableAccess.GetVal("AddressLine3"))
End Function
Friend Function GetAddressLine4() As String
    GetAddressLine4 = CStr(m_clsTableAccess.GetVal("AddressLine4"))
End Function
Friend Function GetAddressLine5() As String
    GetAddressLine5 = CStr(m_clsTableAccess.GetVal("AddressLine5"))
End Function
Friend Function GetAddressLine6() As String
    GetAddressLine6 = CStr(m_clsTableAccess.GetVal("AddressLine6"))
End Function
Friend Function GetProcFeeOnlineInd() As String
    GetProcFeeOnlineInd = CStr(m_clsTableAccess.GetVal("ProcFeeOnlineInd"))
End Function
Friend Function GetPostcode() As String
    GetPostcode = CStr(m_clsTableAccess.GetVal("PostCode"))
End Function

Friend Function GetTelephoneCountryCode() As String
    GetTelephoneCountryCode = CStr(m_clsTableAccess.GetVal("TelephoneCountryCode"))
End Function
Friend Function GetTelephoneAreaCode() As String
    GetTelephoneAreaCode = CStr(m_clsTableAccess.GetVal("TelephoneAreaCode"))
End Function
Friend Function GetTelephoneNumber() As String
    GetTelephoneNumber = CStr(m_clsTableAccess.GetVal("TelephoneNumber"))
End Function
Friend Function GetFaxCountryCode() As String
    GetFaxCountryCode = CStr(m_clsTableAccess.GetVal("FaxCountryCode"))
End Function
Friend Function GetFaxAreaCode() As String
    GetFaxAreaCode = CStr(m_clsTableAccess.GetVal("FaxAreaCode"))
End Function
Friend Function GetFaxNumber() As String
    GetFaxNumber = CStr(m_clsTableAccess.GetVal("FaxNumber"))
End Function
Friend Function GetLastUpdatedDate() As String
    GetLastUpdatedDate = CStr(m_clsTableAccess.GetVal("LastUpdatedDate"))
End Function
Friend Function GetLastUpdatedBy() As String
    GetLastUpdatedBy = CStr(m_clsTableAccess.GetVal("LastUpdatedBy"))
End Function
Friend Function GetStatusDate() As String
    GetStatusDate = CStr(m_clsTableAccess.GetVal("StatusDate"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetARFirmID(sValue As Variant) As Boolean
    SetARFirmID = m_clsTableAccess.SetVal("ARFirmId", CVar(sValue))
End Function
Friend Function SetFSAARFirmRef(sValue As Variant) As Boolean
    SetFSAARFirmRef = m_clsTableAccess.SetVal("FSAARFirmRef", CVar(sValue))
End Function
Friend Function SetPrincipalFirmID(sValue As Variant) As Boolean
    SetPrincipalFirmID = m_clsTableAccess.SetVal("PrincipalFirmID", CVar(sValue))
End Function
Friend Function SetARFirmName(sValue As Variant) As Boolean
    SetARFirmName = m_clsTableAccess.SetVal("ARFirmName", CVar(sValue))
End Function
Friend Function SetFirmStatus(sValue As Variant) As Boolean
    SetFirmStatus = m_clsTableAccess.SetVal("FirmStatus", CVar(sValue))
End Function
Friend Function SetListingStatus(sValue As Variant) As Boolean
    SetListingStatus = m_clsTableAccess.SetVal("ListingStatus", CVar(sValue))
End Function
Friend Function SetAddressLine1(sValue As Variant) As Boolean
    SetAddressLine1 = m_clsTableAccess.SetVal("AddressLine1", CVar(sValue))
End Function
Friend Function SetAddressLine2(sValue As Variant) As Boolean
    SetAddressLine2 = m_clsTableAccess.SetVal("AddressLine2", CVar(sValue))
End Function
Friend Function SetAddressLine3(sValue As Variant) As Boolean
    SetAddressLine3 = m_clsTableAccess.SetVal("AddressLine3", CVar(sValue))
End Function
Friend Function SetAddressLine4(sValue As Variant) As Boolean
    SetAddressLine4 = m_clsTableAccess.SetVal("AddressLine4", CVar(sValue))
End Function
Friend Function SetAddressLine5(sValue As Variant) As Boolean
    SetAddressLine5 = m_clsTableAccess.SetVal("AddressLine5", CVar(sValue))
End Function
Friend Function SetAddressLine6(sValue As Variant) As Boolean
    SetAddressLine6 = m_clsTableAccess.SetVal("AddressLine6", CVar(sValue))
End Function
Friend Function SetProcFeeOnlineInd(sValue As Variant) As Boolean
    SetProcFeeOnlineInd = m_clsTableAccess.SetVal("ProcFeeOnlineInd", CVar(sValue))
End Function
Friend Function SetPostcode(sValue As Variant) As Boolean
    SetPostcode = m_clsTableAccess.SetVal("PostCode", CVar(sValue))
End Function

Friend Function SetTelephoneCountryCode(sValue As Variant) As Boolean
    SetTelephoneCountryCode = m_clsTableAccess.SetVal("TelephoneCountryCode", CVar(sValue))
End Function
Friend Function SetTelephoneAreaCode(sValue As Variant) As Boolean
    SetTelephoneAreaCode = m_clsTableAccess.SetVal("TelephoneAreaCode", CVar(sValue))
End Function
Friend Function SetTelephoneNumber(sValue As Variant) As Boolean
    SetTelephoneNumber = m_clsTableAccess.SetVal("TelephoneNumber", CVar(sValue))
End Function
Friend Function SetFaxCountryCode(sValue As Variant) As Boolean
    SetFaxCountryCode = m_clsTableAccess.SetVal("FaxCountryCode", CVar(sValue))
End Function
Friend Function SetFaxAreaCode(sValue As Variant) As Boolean
    SetFaxAreaCode = m_clsTableAccess.SetVal("FaxAreaCode", CVar(sValue))
End Function
Friend Function SetFaxNumber(sValue As Variant) As Boolean
    SetFaxNumber = m_clsTableAccess.SetVal("FaxNumber", CVar(sValue))
End Function
Friend Function SetLastUpdatedDate(sValue As Variant) As Boolean
    SetLastUpdatedDate = m_clsTableAccess.SetVal("LastUpdatedDate", CVar(sValue))
End Function
Friend Function SetLastUpdatedBy(sValue As Variant) As Boolean
    SetLastUpdatedBy = m_clsTableAccess.SetVal("LastUpdatedBy", CVar(sValue))
End Function
Friend Function SetStatusDate(sValue As Variant) As Boolean
    SetStatusDate = m_clsTableAccess.SetVal("StatusDate", CVar(sValue))
End Function
Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
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
    colListLine.Add GetARFirmID()
' TW 16/01/2007 EP2_859
    colListLine.Add GetUnitId()
' TW 16/01/2007 EP2_859 End
    colListLine.Add GetFSAARFirmRef()
    colListLine.Add GetARFirmName()
' TW 16/01/2007 EP2_859
    colListLine.Add GetAddress()
' TW 16/01/2007 EP2_859 End
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
