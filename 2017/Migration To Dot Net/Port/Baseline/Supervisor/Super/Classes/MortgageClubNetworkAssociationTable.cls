VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageClubNetAssocTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   MortgageClubNetAssocTable
' Description   :   Contains all access functions for the MortgageClubNetworkAssociation table.
' Change history
' Change history
' Prog     Date         Description
' TW       17/10/2006   EP2_15 - Created
' TW       07/12/2006   EP2_153 - Field PaymentType renamed to PaymentMethod
' TW       07/12/2006   EP2_359 - E2CR24_26 - Changes for new field "Identifier"
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "MORTGAGECLUBNETWORKASSOCIATION"
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
    colFields.Add "ClubNetworkAssociationId"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "MORTGAGECLUBNETWORKASSOCNAME"


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetClubNetworkAssociationID() As String
    GetClubNetworkAssociationID = CStr(m_clsTableAccess.GetVal("ClubNetworkAssociationId"))
End Function
Friend Function GetMortgageClubNetworkAssocDescription() As String
    GetMortgageClubNetworkAssocDescription = CStr(m_clsTableAccess.GetVal("MORTGAGECLUBNETWORKASSOCDESCRIPTION"))
End Function
Friend Function GetAddress() As String
    GetAddress = CStr(m_clsTableAccess.GetVal("ADDRESS"))
End Function
Friend Function GetMortgageClubNetworkAssocName() As String
    GetMortgageClubNetworkAssocName = CStr(m_clsTableAccess.GetVal("MORTGAGECLUBNETWORKASSOCNAME"))
End Function

Friend Function GetBankAccountName() As String
    GetBankAccountName = CStr(m_clsTableAccess.GetVal("BankAccountName"))
End Function
Friend Function GetBankAccountNumber() As String
    GetBankAccountNumber = CStr(m_clsTableAccess.GetVal("BankAccountNumber"))
End Function
Friend Function GetAssociationFeeRate() As String
    GetAssociationFeeRate = CStr(m_clsTableAccess.GetVal("AssociationFeeRate"))
End Function
Friend Function GetAssociationFeeAmount() As String
    GetAssociationFeeAmount = CStr(m_clsTableAccess.GetVal("AssociationFeeAmount"))
End Function

Friend Function GetBankAccountBranchName() As String
    GetBankAccountBranchName = CStr(m_clsTableAccess.GetVal("BankAccountBranchName"))
End Function

Friend Function GetBankSortCode() As String
    GetBankSortCode = CStr(m_clsTableAccess.GetVal("BankSortCode"))
End Function
Friend Function GetBankWizardIndicator() As String
    GetBankWizardIndicator = CStr(m_clsTableAccess.GetVal("BankWizardIndicator"))
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
Friend Function GetIdentifier() As String
    GetIdentifier = CStr(m_clsTableAccess.GetVal("Identifier"))
End Function
Friend Function GetPostCode() As String
    GetPostCode = CStr(m_clsTableAccess.GetVal("PostCode"))
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
Friend Function GetVolProcFeeAdjustment() As String
    GetVolProcFeeAdjustment = CStr(m_clsTableAccess.GetVal("VolProcFeeAdjustment"))
End Function
Friend Function GetProcLoadingInd() As String
    GetProcLoadingInd = CStr(m_clsTableAccess.GetVal("ProcLoadingInd"))
End Function
Friend Function GetPackagerIndicator() As String
    GetPackagerIndicator = CStr(m_clsTableAccess.GetVal("PackagerIndicator"))
End Function
Friend Function GetLastUpdatedDate() As String
    GetLastUpdatedDate = CStr(m_clsTableAccess.GetVal("LastUpdatedDate"))
End Function
Friend Function GetLastUpdatedBy() As String
    GetLastUpdatedBy = CStr(m_clsTableAccess.GetVal("LastUpdatedBy"))
End Function
Friend Function GetPaymentMethod() As String
' TW 07/12/2006 EP2_153
    'GetPaymentType = CStr(m_clsTableAccess.GetVal("PaymentType"))
    GetPaymentMethod = CStr(m_clsTableAccess.GetVal("PaymentMethod"))
' TW 07/12/2006 EP2_153 End
End Function
Friend Function GetAgreedProcFeeRate() As String
    GetAgreedProcFeeRate = CStr(m_clsTableAccess.GetVal("AgreedProcFeeRate"))
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetClubNetworkAssociationId(sValue As Variant) As Boolean
    SetClubNetworkAssociationId = m_clsTableAccess.SetVal("ClubNetworkAssociationId", CVar(sValue))
End Function
Friend Function SetMortgageClubNetworkAssocDescription(sValue As Variant) As Boolean
    SetMortgageClubNetworkAssocDescription = m_clsTableAccess.SetVal("MORTGAGECLUBNETWORKASSOCDESCRIPTION", CVar(sValue))
End Function
Friend Function SetMortgageClubNetworkAssocName(sValue As Variant) As Boolean
    SetMortgageClubNetworkAssocName = m_clsTableAccess.SetVal("MORTGAGECLUBNETWORKASSOCNAME", CVar(sValue))
End Function
Friend Function SetBankAccountName(sValue As Variant) As Boolean
    SetBankAccountName = m_clsTableAccess.SetVal("BankAccountName", CVar(sValue))
End Function
Friend Function SetBankAccountNumber(sValue As Variant) As Boolean
    SetBankAccountNumber = m_clsTableAccess.SetVal("BankAccountNumber", CVar(sValue))
End Function
Friend Function SetAssociationFeeRate(sValue As Variant) As Boolean
    SetAssociationFeeRate = m_clsTableAccess.SetVal("AssociationFeeRate", CVar(sValue))
End Function
Friend Function SetAssociationFeeAmount(sValue As Variant) As Boolean
    SetAssociationFeeAmount = m_clsTableAccess.SetVal("AssociationFeeAmount", CVar(sValue))
End Function

Friend Function SetBankAccountBranchName(sValue As Variant) As Boolean
    SetBankAccountBranchName = m_clsTableAccess.SetVal("BankAccountBranchName", CVar(sValue))
End Function

Friend Function SetBankSortCode(sValue As Variant) As Boolean
    SetBankSortCode = m_clsTableAccess.SetVal("BankSortCode", CVar(sValue))
End Function
Friend Function SetBankWizardIndicator(sValue As Variant) As Boolean
    SetBankWizardIndicator = m_clsTableAccess.SetVal("BankWizardIndicator", CVar(sValue))
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
Friend Function SetIdentifier(sValue As Variant) As Boolean
    SetIdentifier = m_clsTableAccess.SetVal("Identifier", CVar(sValue))
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
Friend Function SetVolProcFeeAdjustment(sValue As Variant) As Boolean
    SetVolProcFeeAdjustment = m_clsTableAccess.SetVal("VolProcFeeAdjustment", CVar(sValue))
End Function
Friend Function SetProcLoadingInd(sValue As Variant) As Boolean
    SetProcLoadingInd = m_clsTableAccess.SetVal("ProcLoadingInd", CVar(sValue))
End Function
Friend Function SetPackagerIndicator(sValue As Variant) As Boolean
    SetPackagerIndicator = m_clsTableAccess.SetVal("PackagerIndicator", CVar(sValue))
End Function
Friend Function SetLastUpdatedDate(sValue As Variant) As Boolean
    SetLastUpdatedDate = m_clsTableAccess.SetVal("LastUpdatedDate", CVar(sValue))
End Function
Friend Function SetLastUpdatedBy(sValue As Variant) As Boolean
    SetLastUpdatedBy = m_clsTableAccess.SetVal("LastUpdatedBy", CVar(sValue))
End Function
Friend Function SetPaymentMethod(sValue As Variant) As Boolean
' TW 07/12/2006 EP2_153
'    SetPaymentType = m_clsTableAccess.SetVal("PaymentType", CVar(sValue))
    SetPaymentMethod = m_clsTableAccess.SetVal("PaymentMethod", CVar(sValue))
' TW 07/12/2006 EP2_153 End
End Function
Friend Function SetAgreedProcFeeRate(sValue As Variant) As Boolean
    SetAgreedProcFeeRate = m_clsTableAccess.SetVal("AgreedProcFeeRate", CVar(sValue))
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
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType, enumClassOption)
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
    colListLine.Add GetClubNetworkAssociationID()
    colListLine.Add GetMortgageClubNetworkAssocName()
' TW 16/01/2007 EP2_859
'    colListLine.Add GetMortgageClubNetworkAssocDescription()
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
