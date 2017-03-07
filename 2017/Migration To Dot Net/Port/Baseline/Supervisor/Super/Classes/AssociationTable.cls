VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssociationTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : AssociationTable
' Description   : Contains functions to read/write/search the association table
' TW        18/12/2006  Remove the use of ASSOCIATION view from Supervisor
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "MORTGAGECLUBNETWORKASSOCIATION"
Private m_sSearch As String
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "ASSOCIATIONID"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
Public Function GetAssociationId() As Variant
    GetAssociationId = m_clsTableAccess.GetVal("CLUBNETWORKASSOCIATIONID")
End Function
Public Function SetAssociationId(strAssociationId As String)
    m_clsTableAccess.SetVal "ASSOCIATIONID", strAssociationId
End Function
Public Function GetDescription() As Variant
    GetDescription = m_clsTableAccess.GetVal("MORTGAGECLUBNETWORKASSOCDESCRIPTION")
End Function
Public Sub SetDescription(strDescription As String)
    m_clsTableAccess.SetVal "DESCRIPTION", strDescription
End Sub
Public Function GetName() As Variant
    GetName = m_clsTableAccess.GetVal("MORTGAGECLUBNETWORKASSOCNAME")
End Function
Public Sub SetName(strName As String)
    m_clsTableAccess.SetVal "NAME", strName
End Sub
Public Function GetBuildingHouseName() As String
    GetBuildingHouseName = m_clsTableAccess.GetVal("BUILDINGORHOUSENAME")
End Function
Public Sub SetBuildingHouseName(strBuildingHouseName As String)
    m_clsTableAccess.SetVal "BUILDINGHOUSENAME", strBuildingHouseName
End Sub
Public Function GetBuildingHouseNumber() As String
    GetBuildingHouseNumber = m_clsTableAccess.GetVal("BUILDINGORHOUSENUMBER")
End Function
Public Sub SetBuildingHouseNumber(strBuildingHouseNumber As String)
    m_clsTableAccess.SetVal "BUILDINGHOUSENUMBER", strBuildingHouseNumber
End Sub
Public Function GetFlatNumber() As String
    GetFlatNumber = m_clsTableAccess.GetVal("FLATNUMBER")
End Function
Public Sub SetFlatNumber(strFlatNumber As String)
    m_clsTableAccess.SetVal "FLATNUMBER", strFlatNumber
End Sub
Public Function GetStreet() As String
    GetStreet = m_clsTableAccess.GetVal("STREET")
End Function
Public Sub SetStreet(strStreet As String)
    m_clsTableAccess.SetVal "STREET", strStreet
End Sub
Public Function GetDistrict() As String
    GetDistrict = m_clsTableAccess.GetVal("DISTRICT")
End Function
Public Sub SetDistrict(strDistrict As String)
    m_clsTableAccess.SetVal "DISTRICT", strDistrict
End Sub
Public Function GetTown() As String
    GetTown = m_clsTableAccess.GetVal("TOWN")
End Function
Public Sub SetTown(strTown As String)
    m_clsTableAccess.SetVal "TOWN", strTown
End Sub
Public Function GetCounty() As String
    GetCounty = m_clsTableAccess.GetVal("COUNTY")
End Function
Public Sub SetCounty(strCounty As String)
    m_clsTableAccess.SetVal "COUNTY", strCounty
End Sub
Public Function GetCountry() As String
    GetCountry = m_clsTableAccess.GetVal("COUNTRY")
End Function
Public Sub SetCountry(strCountry As String)
    m_clsTableAccess.SetVal "COUNTRY", strCountry
End Sub
Public Function GetPostCode() As String
    GetPostCode = m_clsTableAccess.GetVal("POSTCODE")
End Function
Public Sub SetPostcode(strPostCode As String)
    m_clsTableAccess.SetVal "POSTCODE", strPostCode
End Sub
Public Function GetTelephoneCountryCode() As String
    GetTelephoneCountryCode = m_clsTableAccess.GetVal("TELEPHONECOUNTRYCODE")
End Function
Public Sub SetTelephoneCountryCode(strTelephoneCountryCode As String)
    m_clsTableAccess.SetVal "TELEPHONECOUNTRYCODE", strTelephoneCountryCode
End Sub
Public Function GetTelephoneAreaCode() As String
    GetTelephoneAreaCode = m_clsTableAccess.GetVal("TELEPHONEAREACODE")
End Function
Public Sub SetTelephoneAreaCode(strTelephoneAreaCode As String)
    m_clsTableAccess.SetVal "TELEPHONEAREACODE", strTelephoneAreaCode
End Sub
Public Function GetTelephoneNumber() As String
    GetTelephoneNumber = m_clsTableAccess.GetVal("TELEPHONENUMBER")
End Function
Public Sub SetTelephoneNumber(strTelephoneNumber As String)
    m_clsTableAccess.SetVal "TELEPHONENUMBER", strTelephoneNumber
End Sub
Public Function GetFaxCountryCode() As String
    GetFaxCountryCode = m_clsTableAccess.GetVal("FAXCOUNTRYCODE")
End Function
Public Sub SetFaxCountryCode(strFaxCountryCode As String)
    m_clsTableAccess.SetVal "FAXCOUNTRYCODE", strFaxCountryCode
End Sub
Public Function GetFaxAreaCode() As String
    GetFaxAreaCode = m_clsTableAccess.GetVal("FAXAREACODE")
End Function
Public Sub SetFaxAreaCode(strFaxAreaCode As String)
    m_clsTableAccess.SetVal "FAXAREACODE", strFaxAreaCode
End Sub
Public Function GetFaxNumber() As String
    GetFaxNumber = m_clsTableAccess.GetVal("FAXNUMBER")
End Function
Public Sub SetFaxNumber(strFaxNumber As String)
    m_clsTableAccess.SetVal "FAXNUMBER", strFaxNumber
End Sub
Public Function GetAddress() As String
    GetAddress = m_clsTableAccess.GetVal("ADDRESS")
End Function
Public Function GetAccountName() As String
    GetAccountName = m_clsTableAccess.GetVal("ACCOUNTNAME")
End Function

Public Sub SetAccountName(strAccountName As String)
    m_clsTableAccess.SetVal "ACCOUNTNAME", strAccountName
End Sub
Public Function GetAccountNumber() As String
    GetAccountNumber = m_clsTableAccess.GetVal("ACCOUNTNUMBER")
End Function
Public Sub SetAccountNumber(strAccountNumber)
    m_clsTableAccess.SetVal "ACCOUNTNUMBER", strAccountNumber
End Sub
Public Function GetBranchName() As String
    GetBranchName = m_clsTableAccess.GetVal("BRANCHNAME")
End Function
Public Sub SetBranchName(strBranchName)
    m_clsTableAccess.SetVal "BRANCHNAME", strBranchName
End Sub
Public Function GetSortCode() As String
    GetSortCode = m_clsTableAccess.GetVal("SORTCODE")
End Function
Public Sub SetSortCode(strSortCode)
    m_clsTableAccess.SetVal "SORTCODE", strSortCode
End Sub
Public Function GetDetailsValidated() As String
    GetDetailsValidated = m_clsTableAccess.GetVal("DETAILSVALIDATED")
End Function
Public Sub SetDetailsValidated(strDetailsValidated)
    m_clsTableAccess.SetVal "DETAILSVALIDATED", strDetailsValidated
End Sub
Public Function GetAgreedProcFeeRate() As String
    GetAgreedProcFeeRate = m_clsTableAccess.GetVal("AGREEDPROCFEERATE")
End Function
Public Sub SetAgreedProcFeeRate(strAgreedProcFeeRate)
    m_clsTableAccess.SetVal "AGREEDPROCFEERATE", strAgreedProcFeeRate
End Sub
Public Function GetVolProcFeeAdjustment() As String
    GetVolProcFeeAdjustment = m_clsTableAccess.GetVal("VOLPROCFEEADJUSTMENT")
End Function
Public Sub SetVolProcFeeAdjustment(strVolProcFeeAdjustment)
    m_clsTableAccess.SetVal "VOLPROCFEEADJUSTMENT", strVolProcFeeAdjustment
End Sub
Public Function GetProcLoadingIndicator() As String
    GetProcLoadingIndicator = m_clsTableAccess.GetVal("PROCLOADINGINDICATOR")
End Function
Public Sub SetProcLoadingIndicator(strProcLoadingIndicator)
    m_clsTableAccess.SetVal "PROCLOADINGINDICATOR", strProcLoadingIndicator
End Sub
Public Function GetProcFeePaymentMethod() As String
    GetProcFeePaymentMethod = m_clsTableAccess.GetVal("PROCFEEPAYMENTMETHOD")
End Function
Public Sub SetProcFeePaymentMethod(strProcFeePaymentMethod)
    m_clsTableAccess.SetVal "PROCFEEPAYMENTMETHOD", strProcFeePaymentMethod
End Sub
'
'
' ... ADD MORE SUBS HERE
'
'
'
'
'

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
    colListLine.Add GetAssociationId
    colListLine.Add GetDescription
    colListLine.Add GetName
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActiveApplications
' Description   : Returns all active applications. An active application is one that doesn't have
'                 an entry in the applicationstage table for Cancelled or Complete. This function
'                 will return the lead application name and address, too.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetActiveAssociations(Optional sAssId As String, Optional sName As String, Optional sDescription As String, Optional bUseExistingSearch As Boolean = False)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
       
    sSearch = "SELECT Distinct ASSOCIATIONID, DESCRIPTION, NAME From ASSOCIATION "
    
    If Len(sAssId) > 0 Then
        sSearch = sSearch + "Where ASSOCIATION.ASSOCIATIONID = " & sAssId
    End If
    
    If Len(sName) > 0 Then
        sSearch = sSearch + " AND ASSOCIATION.NAME LIKE " + g_clsFormProcessing.FormatWildcardString(sName)
    End If
    
    If Len(sDescription) > 0 Then
        sSearch = sSearch + " AND CV.SURNAME LIKE " + g_clsFormProcessing.FormatWildcardString(sDescription)
    End If
    
    sSearch = sSearch + " ORDER BY ASSOCIATION.ASSOCIATIONID"

    m_clsTableAccess.GetTableDataFromSearch sSearch, bUseExistingSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


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

Public Function LoadRecord() As Boolean
    '
    Dim objRS As ADODB.Recordset
    Dim strSQL
    '
    
End Function
