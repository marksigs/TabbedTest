VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PanelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PanelTable
' Description   : Contains all access routines and SQL for the Panel table. Implements TableAccess
' Change history
' Prog      Date        Description
' DJP       13/06/01    SQL Server port
' STB       13/11/01    Telephone details are deleted along with Third Party.
' STB       11/01/02    SYS3581 Promoting deleted thirdparties was failing as sTargegDatabase
'                       was not being set on clsNameAndAddress.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' TK        30/11/2005  MAR81 Amended GetPanelFromSearch, TableAccess_GetRowOfData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM History
' Prog      Date        Description
' TW        02/01/2007  EP2_640 - Promotion failures
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Constants
Private Const m_sTableName = "PANEL"

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess

Private m_enumSearchType As ThirdPartyType
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS HISTORY
'
' Prog   Date        Description
' DB     04/11/02    Re-ordered the third party list view.
' DB     04/11/02    Re-ordered the third party list view.
' DJP    21/02/02    BM0318 - Third party searches
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "DirectoryGUID"
    
    m_clsTableAccess.SetKeyMatchFields colFields
    m_enumSearchType = ThirdPartyInvalid
End Sub

Friend Function GetDirectoryGUID() As Variant
    GetDirectoryGUID = m_clsTableAccess.GetVal("DirectoryGUID")
End Function

Friend Function GetPanelID() As Variant
    GetPanelID = m_clsTableAccess.GetVal("PanelID")
End Function

Friend Function GetAssociatedPanelID() As Variant
    GetAssociatedPanelID = m_clsTableAccess.GetVal("AssociatedPanelID")
End Function

Friend Function GetPaymentMethod() As Variant
    GetPaymentMethod = m_clsTableAccess.GetVal("PaymentMethod")
End Function

Friend Function GetNameAndAddressTypeText() As Variant
    GetNameAndAddressTypeText = m_clsTableAccess.GetVal("NameAndAddressTypeText")
End Function

Friend Function GetCompanyName() As Variant
    GetCompanyName = m_clsTableAccess.GetVal("CompanyName")
End Function
'TK 29/11/2005 MAR81
Friend Function GetStatusText() As Variant
    GetStatusText = m_clsTableAccess.GetVal("StatusText")
End Function

Friend Function GetPanelLegalRepUserID() As Variant
    GetPanelLegalRepUserID = m_clsTableAccess.GetVal("UserID")
End Function

Friend Function GetOrganisationTypeText() As Variant
    GetOrganisationTypeText = m_clsTableAccess.GetVal("OrganisationTypeText")
End Function

Friend Function GetActiveFrom() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("NameAndAddressActiveFrom", , , False)
    
    If Not IsNull(vVal) Then
        GetActiveFrom = Format(vVal, "Short Date")
    End If

End Function

Friend Function GetActiveTo() As Variant
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetDate("NameAndAddressActiveTo", , , False)
    
    If Not IsNull(vVal) Then
        GetActiveTo = Format(vVal, "Short Date")
    End If

End Function

Friend Function SetDirectoryGUID(vGUID As Variant) As Boolean
    SetDirectoryGUID = m_clsTableAccess.SetVal("DirectoryGUID", vGUID)
End Function

Friend Function SetPanelID(vVal As Variant) As Boolean
    SetPanelID = m_clsTableAccess.SetVal("PanelID", vVal)
End Function

Friend Function SetAssociatedPanelID(vVal As Variant) As Boolean
    SetAssociatedPanelID = m_clsTableAccess.SetVal("AssociatedPanelID", vVal)
End Function

Friend Function SetPaymentMethod(vVal As Variant) As Boolean
    SetPaymentMethod = m_clsTableAccess.SetVal("PaymentMethod", vVal)
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

Public Function TableAccess_GetUpdateValues() As Collection
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    Dim sCompanyName As String
    Dim sAddressType As String
    Dim sStatus As String
    Dim sPanelLegalRepUserID As String
    
    sCompanyName = GetCompanyName()
    sAddressType = GetNameAndAddressTypeText()
    sStatus = GetStatusText() 'TK 29/11/2005 MAR81
    sPanelLegalRepUserID = GetPanelLegalRepUserID()
    'DB BMIDS00720 - Made company name the first column
    colListLine.Add sCompanyName
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
    'colListLine.Add GetOrganisationTypeText()
    colListLine.Add sAddressType

    ' DJP BM0318
    If m_enumSearchType <> ThirdPartyLocalType Then
        colListLine.Add GetPanelID()
    End If
    
    'TK 29/11/2005 MAR81
    colListLine.Add sStatus 'TK 29/11/2005 MAR81
    colListLine.Add sPanelLegalRepUserID
    
    ' DJP BM0318
    If m_enumSearchType <> ThirdPartyInvalid Then
        colListLine.Add GetTown
    End If
    
    colListLine.Add sCompanyName + ", " + sAddressType, OBJECT_DESCRIPTION_KEY
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Public Sub GetPanelValues(enumType As ThirdPartyCombo)
    Dim rs As ADODB.Recordset
    Dim Param As ADODB.Parameter
    Dim colParam As Collection
    
    ' First parameter, the GroupName
    Set Param = New ADODB.Parameter
    Set colParam = New Collection
    
    With Param
        .Name = "GroupName"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = "ThirdPartyType"
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With
    
    colParam.Add Param
    
    ' Second parameter, the type
    Set Param = New ADODB.Parameter
    
    With Param
        .Name = "Type"
        .Type = adNumeric
        .Direction = adParamInput
        .Value = enumType
        ' SQL Server requires the precision
        .Precision = 10
    End With
    
    colParam.Add Param
    
        
    ' DJP SQL Server port
    ' The name is different between SQL Server and Oracle
    
    ' Get the data
    Set rs = g_clsDataAccess.ExecuteStoredProcRS(spGetPanelDetails, colParam)
    
    ValidateRecordset rs, "Panel"
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

Public Sub DeletePanel()
    
    On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim colMatchFields As Collection
    Dim sCommand As String
    Dim sTargetDatabase As String
        
    sTargetDatabase = m_clsTableAccess.GetDatabase()
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues
    
    If colMatchValues.Count = 1 Then
        'g_clsDataAccess.BeginTrans
        ' Panel Valuer Valuation Type
        Set clsTableAccess = New PanelValuerTypeTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
    
        ' Panel Valuer
        Set clsTableAccess = New PanelValuerTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
        ' Bank Accounts
        Set clsTableAccess = New PanelBankAccountTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
        ' Panel Legal Rep
        Set clsTableAccess = New PanelLegalRepTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
        ' Panel
        Set clsTableAccess = New PanelTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.DeleteRecords sTargetDatabase
        
        ' Name and Address Directory
        Dim clsNameAndAddress As New NameAndAddressDirTable
        Set clsTableAccess = clsNameAndAddress
        
        clsTableAccess.SetDatabase sTargetDatabase
        clsTableAccess.SetKeyMatchValues colMatchValues
        
        'Prior to deleting the name and address directory, obtain the
        'ContactDetailsGUID so we can remove any telephone number records.
        clsTableAccess.GetTableData
        
' TW 02/01/2007 EP2_640
        If clsTableAccess.RecordCount > 0 Then
' TW 02/01/2007 EP2_640 End
        
            'Create a new collection to hold a new key value.
            Set colMatchValues = New Collection
            Set colMatchFields = New Collection
            
            'Add into the collection, the Contacts GUID from the address record.
            colMatchValues.Add clsNameAndAddress.GetContactDetailsGUID
            colMatchFields.Add "CONTACTDETAILSGUID"
            
            'Delete the address record.
            clsNameAndAddress.DeleteDirectory sTargetDatabase
            
            'Telephone Details.
            Set clsTableAccess = New ContactDetailsTelephoneTable
            
            clsTableAccess.SetKeyMatchFields colMatchFields
            clsTableAccess.SetKeyMatchValues colMatchValues
            clsTableAccess.DeleteRecords sTargetDatabase
' TW 02/01/2007 EP2_640
        End If
' TW 02/01/2007 EP2_640 End
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "Delete Panel"
    End If
    
    Exit Sub
Failed:
    'g_clsDataAccess.RollbackTrans
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetPanelFromSearch
' Description   : Returns Third Party details based on the Company Name, Panel ID and type
'                 of Third Party passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPanelFromSearch(sCompanyName As String, sPanelID As String, sTown As String, nType As Long)
    On Error GoTo Failed
    Dim sSearch As String
    
    If g_clsDataAccess.GetDatabaseType = INDEX_SQL_SERVER Then
    'TK 29/11/2005 MAR81
    sSearch = "  SELECT " & _
                    "NAD.DIRECTORYGUID, " & _
                    "NAD.COMPANYNAME, " & _
                    "NAD.NAMEANDADDRESSACTIVEFROM, " & _
                    "NAD.NAMEANDADDRESSACTIVETO, " & _
                    "CV.VALUENAME NAMEANDADDRESSTYPETEXT, " & _
                    "PAN.PANELID," & _
                    "ADDR.TOWN," & _
                    "(SELECT CV2.VALUENAME FROM COMBOVALUE CV2 WHERE CV2.VALUEID = PLR.STATUS AND CV2.GROUPNAME = 'LEGALREPSTATUS') STATUSTEXT," & _
                    "PLR.USERID " & _
                "FROM " & _
                    "NAMEANDADDRESSDIRECTORY NAD LEFT OUTER JOIN PANEL PAN ON " & _
                    "NAD.DIRECTORYGUID = PAN.DIRECTORYGUID " & _
                    "LEFT OUTER JOIN ADDRESS ADDR ON " & _
                    "NAD.ADDRESSGUID= ADDR.ADDRESSGUID " & _
                    "LEFT OUTER JOIN PANELLEGALREP PLR ON PLR.DIRECTORYGUID = NAD.DIRECTORYGUID, " & _
                    "COMBOVALUE CV " & _
                "WHERE " & _
                    "NAD.NAMEANDADDRESSTYPE = CV.VALUEID AND " & _
                    "NAD.NAMEANDADDRESSTYPE = " & CStr(nType) & " AND " & _
                    "UPPER(CV.GROUPNAME) = 'THIRDPARTYTYPE'"
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Oracle not implemented"
    End If
    
    If Len(sCompanyName) > 0 Then
        sSearch = sSearch & "   AND UPPER(NAD.COMPANYNAME) LIKE " & UCase(sCompanyName)
    End If
    
    If Len(sPanelID) > 0 Then
        sSearch = sSearch & "   AND UPPER(PAN.PANELID) = '" & UCase(sPanelID) & "' "
    End If
    
    If Len(sTown) > 0 Then
        sSearch = sSearch & "   AND UPPER(ADDR.TOWN) LIKE " & UCase(sTown)
    End If
    
    m_clsTableAccess.GetTableDataFromSearch sSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetSearchMode
' Description   : Sets the type of third party this class is associated with.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetSearchMode(enumSearchType As ThirdPartyType)
    m_enumSearchType = enumSearchType
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTown
' Description   : Returns the Town for a Third Party
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetTown() As Variant
    GetTown = m_clsTableAccess.GetVal("TOWN")
End Function

