VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntroducerTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   IntroducerTable
' Description   :   Contains all access functions for the Introducer table.
' Change history
' Change history
' Prog     Date         Description
' TW       17/10/2006   EP2_15 - Created
' TW       15/06/2007   DBM405 - Add "No Cross Selling Agreement" processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "INTRODUCER"
Private Const DELETE_FLAG_SET As Integer = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "INTRODUCERID"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "USERID"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIntroducerID() As String
     GetIntroducerID = CStr(m_clsTableAccess.GetVal("IntroducerID"))
End Function
Friend Function GetNationalInsuranceNumber() As String
    GetNationalInsuranceNumber = CStr(m_clsTableAccess.GetVal("NationalInsuranceNumber"))
End Function
Friend Function GetNoCrossSellInd() As String
    GetNoCrossSellInd = CStr(m_clsTableAccess.GetVal("NoCrossSellInd"))
End Function
Friend Function GetUserID() As String
    GetUserID = CStr(m_clsTableAccess.GetVal("UserId"))
End Function

Friend Function GetEmailAddress() As String
    GetEmailAddress = CStr(m_clsTableAccess.GetVal("EmailAddress"))
End Function
Friend Function GetIntroducerType() As String
    GetIntroducerType = CStr(m_clsTableAccess.GetVal("IntroducerType"))
End Function
Friend Function GetARIndicator() As String
    GetARIndicator = CStr(m_clsTableAccess.GetVal("ARIndicator"))
End Function
Friend Function GetIntroducerStatus() As String
    GetIntroducerStatus = CStr(m_clsTableAccess.GetVal("IntroducerStatus"))
End Function
Friend Function GetListingStatus() As String
    GetListingStatus = CStr(m_clsTableAccess.GetVal("ListingStatus"))
End Function

Friend Function GetBuildingOrHouseName() As String
    GetBuildingOrHouseName = CStr(m_clsTableAccess.GetVal("BuildingOrHouseName"))
End Function
Friend Function GetBuildingOrHouseNumber() As String
    GetBuildingOrHouseNumber = CStr(m_clsTableAccess.GetVal("BuildingOrHouseNumber"))
End Function
Friend Function GetFlatNumber() As String
    GetFlatNumber = CStr(m_clsTableAccess.GetVal("FlatNumber"))
End Function
Friend Function GetStreet() As String
    GetStreet = CStr(m_clsTableAccess.GetVal("Street"))
End Function
Friend Function GetDistrict() As String
    GetDistrict = CStr(m_clsTableAccess.GetVal("District"))
End Function
Friend Function GetTown() As String
    GetTown = CStr(m_clsTableAccess.GetVal("Town"))
End Function
Friend Function GetCounty() As String
    GetCounty = CStr(m_clsTableAccess.GetVal("County"))
End Function
Friend Function GetCountry() As String
    GetCountry = CStr(m_clsTableAccess.GetVal("Country"))
End Function
Friend Function GetPostcode() As String
    GetPostcode = CStr(m_clsTableAccess.GetVal("Postcode"))
End Function
Friend Function GetNumberOfLoginFailures() As String
    GetNumberOfLoginFailures = CStr(m_clsTableAccess.GetVal("NumberOfLoginFailures"))
End Function
Friend Function GetMarketingDataOptOut() As String
    GetMarketingDataOptOut = CStr(m_clsTableAccess.GetVal("MarketingDataOptOut"))
End Function
Friend Function GetLastupdateddate() As String
    GetLastupdateddate = CStr(m_clsTableAccess.GetVal("LastUpdatedDate"))
End Function
Friend Function GetLastUpdatedBy() As String
    GetLastUpdatedBy = CStr(m_clsTableAccess.GetVal("LastUpdatedBy"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIntroducerID(sValue As Variant) As Boolean
    SetIntroducerID = m_clsTableAccess.SetVal("IntroducerID", CVar(sValue))
End Function
Friend Function SetNationalInsuranceNumber(sValue As Variant) As Boolean
    SetNationalInsuranceNumber = m_clsTableAccess.SetVal("NationalInsuranceNumber", CVar(sValue))
End Function
Friend Function SetNoCrossSellInd(sValue As Variant) As Boolean
    SetNoCrossSellInd = m_clsTableAccess.SetVal("NoCrossSellInd", CVar(sValue))
End Function
Friend Function SetUserID(sValue As Variant) As Boolean
    SetUserID = m_clsTableAccess.SetVal("UserId", CVar(sValue))
End Function

Friend Function SetEmailAddress(sValue As Variant) As Boolean
    SetEmailAddress = m_clsTableAccess.SetVal("EmailAddress", CVar(sValue))
End Function
Friend Function SetIntroducerType(sValue As Variant) As Boolean
    SetIntroducerType = m_clsTableAccess.SetVal("IntroducerType", CVar(sValue))
End Function
Friend Function SetARIndicator(sValue As Variant) As Boolean
    SetARIndicator = m_clsTableAccess.SetVal("ARIndicator", CVar(sValue))
End Function
Friend Function SetIntroducerStatus(sValue As Variant) As Boolean
    SetIntroducerStatus = m_clsTableAccess.SetVal("IntroducerStatus", CVar(sValue))
End Function
Friend Function SetListingStatus(sValue As Variant) As Boolean
    SetListingStatus = m_clsTableAccess.SetVal("ListingStatus", CVar(sValue))
End Function
Friend Function SetBuildingOrHouseName(sValue As Variant) As Boolean
    SetBuildingOrHouseName = m_clsTableAccess.SetVal("BuildingOrHouseName", CVar(sValue))
End Function
Friend Function SetBuildingOrHouseNumber(sValue As Variant) As Boolean
    SetBuildingOrHouseNumber = m_clsTableAccess.SetVal("BuildingOrHouseNumber", CVar(sValue))
End Function
Friend Function SetFlatNumber(sValue As Variant) As Boolean
    SetFlatNumber = m_clsTableAccess.SetVal("FlatNumber", CVar(sValue))
End Function
Friend Function SetStreet(sValue As Variant) As Boolean
    SetStreet = m_clsTableAccess.SetVal("Street", CVar(sValue))
End Function
Friend Function SetDistrict(sValue As Variant) As Boolean
    SetDistrict = m_clsTableAccess.SetVal("District", CVar(sValue))
End Function
Friend Function SetTown(sValue As Variant) As Boolean
    SetTown = m_clsTableAccess.SetVal("Town", CVar(sValue))
End Function
Friend Function SetCounty(sValue As Variant) As Boolean
    SetCounty = m_clsTableAccess.SetVal("County", CVar(sValue))
End Function
Friend Function SetCountry(sValue As Variant) As Boolean
    SetCountry = m_clsTableAccess.SetVal("Country", CVar(sValue))
End Function
Friend Function SetPostcode(sValue As Variant) As Boolean
    SetPostcode = m_clsTableAccess.SetVal("Postcode", CVar(sValue))
End Function
Friend Function SetNumberOfLoginFailures(sValue As Variant) As Boolean
    SetNumberOfLoginFailures = m_clsTableAccess.SetVal("NumberOfLoginFailures", CVar(sValue))
End Function
Friend Function SetMarketingDataOptOut(sValue As Variant) As Boolean
    SetMarketingDataOptOut = m_clsTableAccess.SetVal("MarketingDataOptOut", CVar(sValue))
End Function
Friend Function SetLastupdateddate(sValue As Variant) As Boolean
    SetLastupdateddate = m_clsTableAccess.SetVal("LastUpdatedDate", CVar(sValue))
End Function
Friend Function SetLastUpdatedBy(sValue As Variant) As Boolean
    SetLastUpdatedBy = m_clsTableAccess.SetVal("LastUpdatedBy", CVar(sValue))
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
Public Function GetTaskField() As String
    GetTaskField = "TASKNAME"
End Function
Public Function GetComboFields() As Collection
    On Error GoTo Failed
    Dim colFields As Collection
    
    Set colFields = New Collection
    
    colFields.Add "TASKNAME"
    colFields.Add "TASKID"
    
    Set GetComboFields = colFields
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
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

Private Function TableAccess_SetDate(sField As String, vval As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vval)
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
Private Function TableAccess_SetVal(sField As Variant, vval As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vval)
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
    colListLine.Add GetIntroducerID()
    colListLine.Add GetEmailAddress()
    colListLine.Add GetNationalInsuranceNumber()
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
