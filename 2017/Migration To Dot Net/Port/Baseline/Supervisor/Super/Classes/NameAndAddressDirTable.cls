VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NameAndAddressDirTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : NameAndAddressDirTable
' Description   : Contains all access routines for the NameAndAddressDirectory table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Update
'                       SetContactDetailsGUID and SetAddressGUID to take optional GUIDS. If the GUID
'                       is not present, create one.
' AW        08/01/02    SYS3560 - Added GetOtherSystemNumber()
' STB       13/02/02    SYS4054 Added branchname field.
' STB       15/05/02    SYS4558 Added OtherSystemSuffix field.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_bExists As Boolean

' Constants
Private Const m_sTableName = "NAMEANDADDRESSDIRECTORY"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS HISTORY
'
' Prog   Date        Description
' DB     04/11/02    Re-ordered the third party list view.
' DJP    24/02/03    BM0138 Add Town for GetNamesAndAddresses
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "DirectoryGUID"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    m_bExists = g_clsDataAccess.DoesTableExist(m_clsTableAccess.GetTable())
End Sub

Friend Function GetDirectoryGUID() As Variant
    Dim vGUID As Variant
    vGUID = m_clsTableAccess.GetVal("DirectoryGUID")
    ' DJP SQL Server port
    GetDirectoryGUID = vGUID
End Function

Friend Function GetBankSortCode() As Variant
    GetBankSortCode = m_clsTableAccess.GetVal("NAMEANDADDRESSBANKSORTCODE")
End Function

Friend Function GetOrganisationName() As Variant
    GetOrganisationName = m_clsTableAccess.GetVal("OrganisationName")
End Function

Friend Function GetNameAndAddressType() As Variant
    GetNameAndAddressType = m_clsTableAccess.GetVal("NameAndAddressType")
End Function

Friend Function GetNameAndAddressTypeText() As Variant
    GetNameAndAddressTypeText = m_clsTableAccess.GetVal("NameAndAddressTypeText")
End Function

Friend Function GetCompanyName() As Variant
    GetCompanyName = m_clsTableAccess.GetVal("CompanyName")
End Function

Friend Function GetHeadOfficeIndicator() As Variant
    GetHeadOfficeIndicator = m_clsTableAccess.GetVal("HeadOfficeIndicator")
End Function

Friend Function GetActiveFrom() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("NameAndAddressActiveFrom")
    
    If Not IsNull(vVal) Then
        GetActiveFrom = Format(vVal, "Short Date")
    End If

End Function

Friend Function GetActiveTo() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("NameAndAddressActiveTo")
    
    If Not IsNull(vVal) Then
        GetActiveTo = Format(vVal, "Short Date")
    End If

End Function

Friend Function GetDXID() As Variant
    GetDXID = m_clsTableAccess.GetVal("DXID")
End Function

Friend Function GetDXLocation() As Variant
    GetDXLocation = m_clsTableAccess.GetVal("DXLocation")
End Function

Friend Function GetNotes() As Variant
    GetNotes = m_clsTableAccess.GetVal("Notes")
End Function

Friend Function GetAddressGUID() As Variant
    Dim vGUID As Variant
    vGUID = m_clsTableAccess.GetVal("AddressGUID")
    ' DJP SQL Server port
    GetAddressGUID = vGUID
End Function

Friend Function GetContactDetailsGUID() As Variant
    Dim vGUID As Variant
    ' DJP SQL Server port

    vGUID = m_clsTableAccess.GetVal("ContactDetailsGUID")
    GetContactDetailsGUID = vGUID
End Function

Friend Function GetOrganisationType(Optional ctrl As Object = Nothing, Optional ctrlLabel As Object) As Variant
    GetOrganisationType = m_clsTableAccess.GetData("ORGANISATIONTYPE", ctrl, ctrlLabel)
End Function

Friend Function GetOrganisationTypeText() As Variant
    GetOrganisationTypeText = m_clsTableAccess.GetVal("OrganisationTypeText")
End Function

Friend Function GetOtherSystemNumber() As Variant
    GetOtherSystemNumber = m_clsTableAccess.GetVal("OtherSystemNumber")
End Function

Friend Function GetOtherSystemSuffix() As Variant
    GetOtherSystemSuffix = m_clsTableAccess.GetVal("OtherSystemSuffix")
End Function

Friend Function GetBranchName() As Variant
    GetBranchName = m_clsTableAccess.GetVal("BranchName")
End Function

Friend Function SetDirectoryGUID() As Variant
    Dim vVal As Variant
    Dim bRet As Boolean
    
    Dim clsGUID As New GuidAssist
    
    vVal = clsGUID.CreateGUID()
    
    bRet = m_clsTableAccess.SetVal("DIRECTORYGUID", vVal)

    ' DJP SQL Server port - return in database format
    If bRet = True Then
        SetDirectoryGUID = m_clsTableAccess.GetVal("DIRECTORYGUID")
    End If
End Function

Friend Function SetCompanyName(vVal As Variant) As Boolean
    SetCompanyName = m_clsTableAccess.SetVal("CompanyName", vVal)
End Function

Friend Function SetBankSortCode(vVal As Variant) As Boolean
    SetBankSortCode = m_clsTableAccess.SetVal("NAMEANDADDRESSBANKSORTCODE", vVal)
End Function

Friend Function SetOrganisationName(vVal As Variant) As Boolean
    SetOrganisationName = m_clsTableAccess.SetVal("OrganisationName", vVal)
End Function

Friend Function SetNameAndAddressType(vVal As Variant) As Boolean
    SetNameAndAddressType = m_clsTableAccess.SetVal("NameAndAddressType", vVal)
End Function

Friend Function SetHeadOfficeIndicator(vVal As Variant) As Boolean
    SetHeadOfficeIndicator = m_clsTableAccess.SetVal("HeadOfficeIndicator", vVal)
End Function

Friend Function SetActiveFrom(vVal As Variant) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("NameAndAddressActiveFrom", vVal)
End Function

Friend Function SetActiveTo(vVal As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("NameAndAddressActiveTo", vVal)
End Function

Friend Function SetDXID(vVal As Variant) As Boolean
    SetDXID = m_clsTableAccess.SetVal("DXID", vVal)
End Function

Friend Function SetDXLocation(vVal As Variant) As Boolean
    SetDXLocation = m_clsTableAccess.SetVal("DXLocation", vVal)
End Function

Friend Function SetNotes(vVal As Variant) As Boolean
    SetNotes = m_clsTableAccess.SetVal("Notes", vVal)
End Function

Friend Function SetAddressGUID(Optional vVal As Variant) As Variant
    On Error GoTo Failed
    Dim clsGUID As GuidAssist
    
    If IsMissing(vVal) Then
        Set clsGUID = New GuidAssist
        vVal = clsGUID.CreateGUID
    End If
    
    m_clsTableAccess.SetVal "AddressGUID", vVal
    
    SetAddressGUID = GetAddressGUID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Friend Function SetContactDetailsGUID(Optional vVal As Variant) As Variant
    On Error GoTo Failed
    Dim clsGUID As GuidAssist
    
    If IsMissing(vVal) Then
        Set clsGUID = New GuidAssist
        vVal = clsGUID.CreateGUID
    End If
    
    m_clsTableAccess.SetVal "ContactDetailsGUID", vVal
    
    SetContactDetailsGUID = GetContactDetailsGUID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Friend Function SetOrganisationType(vVal As Variant) As Boolean
    m_clsTableAccess.SetData "OrganisationType", vVal '   SetVal("OrganisationType", vVal)
    SetOrganisationType = True
End Function

Friend Function SetBranchName(ByVal vVal As Variant) As Boolean
    SetBranchName = m_clsTableAccess.SetVal("BranchName", vVal)
End Function

Friend Function SetOtherSystemNumber(ByVal vVal As Variant) As Boolean
    SetOtherSystemNumber = m_clsTableAccess.SetVal("OtherSystemNumber", vVal)
End Function

Friend Function SetOtherSystemSuffix(ByVal vVal As Variant) As Boolean
    SetOtherSystemSuffix = m_clsTableAccess.SetVal("OtherSystemSuffix", vVal)
End Function

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
    If m_bExists Then
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
    Dim sCompanyName As String
    Dim sAddressType As String
    
    sCompanyName = GetCompanyName()
    sAddressType = GetNameAndAddressTypeText()
    
    'DB BMIDS00720 - Made company name the first column
    colListLine.Add sCompanyName
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
    'colListLine.Add GetOrganisationTypeText()
    colListLine.Add sAddressType
    colListLine.Add sCompanyName + ", " + sAddressType, OBJECT_DESCRIPTION_KEY
    colListLine.Add GetTown
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTown
' Description   : Returns the Town for a Third Party
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetTown() As Variant
    GetTown = m_clsTableAccess.GetVal("TOWN")
End Function

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Public Sub GetNamesAndAddresses()
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    

    sSearch = " SELECT " & _
                    "DIR.DirectoryGUID, " & _
                    "DIR.CompanyName, " & _
                    "DIR.NAMEANDADDRESSACTIVEFROM, " & _
                    "DIR.NAMEANDADDRESSACTIVETO, " & _
                    "CV.ValueName NameAndAddressTypeText, " & _
                    "DIR.BRANCHNAME, " & _
                    "ADDR.TOWN "
    
    If g_clsDataAccess.GetDatabaseType = INDEX_SQL_SERVER Then
        sSearch = sSearch & _
        "FROM " & _
            m_sTableName & " DIR LEFT OUTER JOIN ADDRESS ADDR ON " & _
            "DIR.ADDRESSGUID = ADDR.ADDRESSGUID, " & _
            "ComboValue CV " & _
        "WHERE " & _
            "DIR.NAMEANDADDRESSTYPE = CV.ValueID AND " & _
            "CV.GroupName = 'ThirdPartyType' AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyValuer & " AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyLegalRep & " AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyValuer
    Else
        ' DJP BM0318 Oracle untested
        sSearch = sSearch & _
        "FROM " & _
            m_sTableName & " DIR, " & _
            " ADDRESS ADDR,  " & _
            "ComboValue CV " & _
        "WHERE " & _
            "DIR.ADDRESSGUID = ADDR.ADDRESSGUID (+), " & _
            "DIR.NAMEANDADDRESSTYPE = CV.ValueID AND " & _
            "CV.GroupName = 'ThirdPartyType' AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyValuer & " AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyLegalRep & " AND " & _
            "DIR.NAMEANDADDRESSTYPE <> " & ThirdPartyValuer
    End If
    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)
    
    ValidateRecordset rs, "NameAndAddressDirectory"
    m_clsTableAccess.SetRecordSet rs

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteDirectory(Optional sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim sContactDetails As String
    Dim sAddressDetails As String
    Dim colMatchData As Collection
    Dim sCommand As String
    
    Set colMatchData = New Collection
    m_clsTableAccess.SetDatabase sTargetDatabase
    Set rs = m_clsTableAccess.GetTableData()
    m_clsTableAccess.ValidateData
    
    If m_clsTableAccess.RecordCount() > 0 Then
        m_clsTableAccess.MoveFirst
        
        Do While Not rs.EOF
            sContactDetails = GetContactDetailsGUID()
            sAddressDetails = GetAddressGUID()
        
            rs.Delete
            m_clsTableAccess.Update
        
            ' Now the ContactDetails
            sCommand = "DELETE FROM CONTACTDETAILS WHERE CONTACTDETAILSGUID = " + g_clsSQLAssistSP.FormatString(sContactDetails)
            g_clsDataAccess.ExecuteCommand sCommand, sTargetDatabase
        
            ' Now the Address
            sCommand = "DELETE FROM ADDRESS WHERE ADDRESSGUID = " + g_clsSQLAssistSP.FormatString(sAddressDetails)
            g_clsDataAccess.ExecuteCommand sCommand, sTargetDatabase
            rs.MoveNext
        Loop
    End If
    
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

