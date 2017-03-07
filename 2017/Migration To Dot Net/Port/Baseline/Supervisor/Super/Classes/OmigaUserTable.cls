VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmigaUserTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   OmigaUserTable
' Description   :   Contains all data/functions associated with the OmigaUser table.
'                   Implements TableAccess for generic routines.
' Change history
' Prog      Date        Description
' DJP       30/11/00    Phase 2 Task Management, add GetActiveUsers
' DJP       26/11/01    SYS3130 - When deleting, delete OmigaUser table too.
' STB       07/01/02    SYS3581 - DeleteUser() now handles external/target databases.
' STB       23/01/02    SYS2957 - DeleteUser() Also deletes users supervisor access rights.
' STB       06/03/02    SYS4229 - Added ContactTelephoneDetails relationship to
'                       DeleteUser(). Also removed a bug where users with
'                       Supervisor permissions couldn't be deleted.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Constants
Private Const m_sTableName = "OMIGAUSER"

' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' DJP       06/03/03    BM0282 Added searching for users.
' HMA       04/02/2004  BMIDS678  Added Credit Check Access
'--------------------------------------------------------------------------------------
' MARS history
' Prog      Date        Description
' GHun      09/08/2005  MAR34 Work around timeout error
' GHun      21/10/2005  MAR264 added workgroup user
'--------------------------------------------------------------------------------------
' EPSOM history
' Prog      Date        Description
' TW        18/12/2006  EP2_568 Add functionality to select which Omiga Users are returned
' TW        02/01/2007  EP2_640 DeleteUser not checking all associated tables correctly
'--------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "UserID"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    m_clsTableAccess.SetCollectionField "USERID"

End Sub

Friend Function GetUserID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("UserID")
    GetUserID = CStr(vVal)
End Function

Friend Function GetUserFullName() As String
    GetUserFullName = m_clsTableAccess.GetVal("UserFullName")
End Function

Friend Function GetAccessTypeText() As String
    GetAccessTypeText = m_clsTableAccess.GetVal("AccessTypeText")
End Function

Friend Function GetAccessType() As String
    GetAccessType = m_clsTableAccess.GetVal("AccessType")
End Function

Friend Function GetActiveFrom() As String
    GetActiveFrom = m_clsTableAccess.GetDate("OmigaUserActiveFrom")
End Function

Friend Function GetActiveTo() As String
    GetActiveTo = m_clsTableAccess.GetDate("OmigaUserActiveTo")
End Function

Friend Function GetNotes() As String
    GetNotes = m_clsTableAccess.GetVal("Notes")
End Function

Friend Function GetChangePassword() As Boolean
    Dim nChange As Integer
    
    nChange = m_clsTableAccess.GetVal("ChangePasswordIndicator")
    
    If nChange = 1 Then
        GetChangePassword = True
    Else
        GetChangePassword = False
    End If
End Function

Friend Function GetWorkingHourType() As String
    GetWorkingHourType = m_clsTableAccess.GetVal("WorkingHourType")
End Function

' BMIDS678
Friend Function GetCreditCheckAccess() As Variant
    GetCreditCheckAccess = m_clsTableAccess.GetVal("CreditCheckAccess")
End Function

'MAR264 GHun
Friend Function GetWorkgroupUser() As Variant
    GetWorkgroupUser = m_clsTableAccess.GetVal("WorkgroupUser")
End Function
'MAR264 End

Friend Function SetUserID(sUserID As String) As Boolean
    SetUserID = m_clsTableAccess.SetVal("UserID", CVar(sUserID))
End Function

Friend Function SetAccessType(sAccessType As String) As Boolean
    SetAccessType = m_clsTableAccess.SetVal("AccessType", CVar(sAccessType))
End Function

Friend Function SetActiveFrom(sActiveFrom As Variant) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("OmigaUserActiveFrom", CVar(sActiveFrom))
End Function

Friend Function SetActiveTo(sActiveTo As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("OmigaUserActiveTo", CVar(sActiveTo))
End Function

Friend Function SetNotes(sNotes As String) As Boolean
    SetNotes = m_clsTableAccess.SetVal("Notes", CVar(sNotes))
End Function

Friend Function SetChangePassword(Optional bChange As Boolean = True) As Boolean
    Dim nChange As Integer
    
    If bChange = True Then
        nChange = 1
    Else
        nChange = 0
    End If
    
    SetChangePassword = m_clsTableAccess.SetVal("ChangePasswordIndicator", CVar(nChange))
End Function

Friend Function SetWorkingHourType(sType As String) As Boolean
    SetWorkingHourType = m_clsTableAccess.SetVal("WorkingHourType", CVar(sType))
End Function

' BMIDS678
Friend Function SetCreditCheckAccess(sType As Variant) As Boolean
    SetCreditCheckAccess = m_clsTableAccess.SetVal("CreditCheckAccess", CVar(sType))
End Function

'MAR264 GHun
Friend Function SetWorkgroupUser(sType As Variant) As Boolean
    SetWorkgroupUser = m_clsTableAccess.SetVal("WorkgroupUser", CVar(sType))
End Function
'MAR264 End

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
    colListLine.Add GetUserID()
    colListLine.Add GetUserFullName()
    'colListLine.Add GetAccessTypeText()
    colListLine.Add GetAccessTypeText()
    colListLine.Add GetActiveFrom()
    colListLine.Add GetActiveTo()
    
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Public Sub GetActiveUsers()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim sSearch As String
    Dim sSystemDate As String
    
    sSystemDate = g_clsSQLAssistSP.GetSystemDate()
    
    sSearch = "SELECT " & _
                    "USERID " & _
                "FROM " & _
                    "OMIGAUSER " & _
                "WHERE " & _
                    "OMIGAUSERACTIVEFROM <= " & sSystemDate & " AND " & _
                    "OMIGAUSERACTIVETO >= " & sSystemDate
     
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    ValidateRecordset rs, m_sTableName
    m_clsTableAccess.SetRecordSet rs
     
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetUsersFromUnit(sUnit As String)
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSystemDate As String
    Dim rs As ADODB.Recordset
    
    ' DJP SQL Server port
    sSystemDate = g_clsSQLAssistSP.TruncateDateColumn(g_clsSQLAssistSP.GetSystemDate())

    sSearch = "SELECT " & _
                    "USERID " & _
                "FROM " & _
                    "OMIGAUSER " & _
                "WHERE " & _
                    "USERID IN ( " & _
                        "SELECT " & _
                            "USERID " & _
                        "FROM " & _
                            "USERROLE " & _
                        "WHERE " & _
                            "UNITID = " & g_clsSQLAssistSP.FormatString(sUnit) & _
                            " AND " & g_clsSQLAssistSP.TruncateDateColumn("USERROLEACTIVEFROM") & " <= " & sSystemDate & " AND (" & _
                            g_clsSQLAssistSP.TruncateDateColumn("USERROLEACTIVETO") & " >= " & sSystemDate & " OR " & g_clsSQLAssistSP.TruncateDateColumn("USERROLEACTIVETO") & " IS NULL) )" & _
                    " AND " & g_clsSQLAssistSP.TruncateDateColumn("OMIGAUSERACTIVEFROM") & " <= " & sSystemDate & " AND (" & _
                    g_clsSQLAssistSP.TruncateDateColumn("OMIGAUSERACTIVETO") & " >= " & sSystemDate & _
                    " OR " & g_clsSQLAssistSP.TruncateDateColumn("OMIGAUSERACTIVETO") & " IS NULL)"
                    
    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

    ValidateRecordset rs, m_sTableName
    m_clsTableAccess.SetRecordSet rs
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub ValidateSuperUser(sUserID As String)
    On Error GoTo Failed
    Dim sSearch As String
    Dim sSysDate As String
    Dim rs As ADODB.Recordset

    ' DJP SQL Server port
    If Len(sUserID) > 0 Then
        sSysDate = g_clsSQLAssistSP.GetSystemDate
        sSysDate = g_clsSQLAssistSP.TruncateDateColumn(sSysDate)
        
        'MAR34 GHun use NoLock locking hint on ComboValidation table
        sSearch = "SELECT * " & _
                    "FROM " & _
                        "USERROLE " & _
                    "WHERE " & _
                        "USERID = " & g_clsSQLAssistSP.FormatString(sUserID) & " AND " & _
                        g_clsSQLAssistSP.TruncateDateColumn("USERROLEACTIVEFROM") & " <= " & sSysDate & " AND (" & _
                        g_clsSQLAssistSP.TruncateDateColumn("USERROLEACTIVETO") & " >= " & sSysDate & " OR USERROLEACTIVETO IS NULL) AND " & _
                        "ROLE IN ( " & _
                            "SELECT " & _
                                "VALUEID " & _
                            "FROM " & _
                                "COMBOVALIDATION WITH (NOLOCK) " & _
                            "WHERE " & _
                                "UPPER(GROUPNAME) = 'USERROLE' AND UPPER(VALIDATIONTYPE) = 'S')"


        Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSearch)

        ValidateRecordset rs, m_sTableName
        m_clsTableAccess.SetRecordSet rs
    
        If rs.RecordCount = 0 Then
            g_clsErrorHandling.RaiseError errUserNotSupervisor
        End If
    Else
        g_clsErrorHandling.RaiseError ErrUserIDIsEmpty
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' BM0282
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetUsers
' Description   : Returns Users based on the criteria passed in.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetUsers(Optional ByVal sUserID As String, Optional ByVal sSurName As String, Optional ByVal sForeName As String, Optional intReturnType As LV_DETAIL = LV_DETAIL_ALL)
    On Error GoTo Failed
    Dim sProc As String
    Dim rs As ADODB.Recordset
    Dim bFound As Boolean
    Dim colParam As Collection
    Dim Param As ADODB.Parameter
    
    Set Param = New ADODB.Parameter
    Set colParam = New Collection
    
    If Len(sUserID) > 0 Then
        sUserID = g_clsSQLAssistSP.FormatWildcardedString(sUserID, bFound)
    End If

    With Param
        .Name = "USERID"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = sUserID
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With

    colParam.Add Param
    Set Param = New ADODB.Parameter

    If Len(sSurName) > 0 Then
        sSurName = g_clsSQLAssistSP.FormatWildcardedString(sSurName, bFound)
    End If

    With Param
        .Name = "SURNAME"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = sSurName
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With

    colParam.Add Param
    Set Param = New ADODB.Parameter

    If Len(sForeName) > 0 Then
        sForeName = g_clsSQLAssistSP.FormatWildcardedString(sForeName, bFound)
    End If

    With Param
        .Name = "FORENAME"
        .Type = adBSTR
        .Direction = adParamInput
        .Value = sForeName
        ' SQL Server requires the length
        .Size = Len(.Value)
    End With

    colParam.Add Param

' TW 18/12/2006 EP2_568
    Set Param = New ADODB.Parameter
    With Param
        .Name = "INCLUDEINTRODUCERS"
        .Type = adInteger
        .Direction = adParamInput
        .Value = intReturnType
    End With

    colParam.Add Param
' TW 18/12/2006 EP2_568 End

    Set rs = g_clsDataAccess.ExecuteStoredProcRS(spGetUsers, colParam)
    
    ValidateRecordset rs, m_sTableName
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Function GetComboField() As String
    GetComboField = "UserID"
End Function

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteUser()
        
Dim sUserID As String
Dim vContactGUID As Variant
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection
Dim colContactMatchValues As Collection
Dim iManagement As ISecurityManager
Dim clsUserContactTable As OrgUserContactTable
' TW 02/01/2007 EP2_640
Dim rs As ADODB.Recordset
Dim sCommand As String
' TW 02/01/2007 EP2_640 End
    
    On Error GoTo Failed
    
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues()
    
    If colMatchValues.Count = 1 Then
        
' TW 02/01/2007 EP2_640
' Check to see if there are any USERHISTORY records for this user
        sUserID = colMatchValues(1)
        sCommand = "SELECT COUNT(1) FROM USERHISTORY WHERE USERID = '" & sUserID & "'"
        Set rs = g_clsDataAccess.ExecuteCommand(sCommand)
        If rs(0) > 0 Then
            MsgBox "User '" & sUserID & "' cannot be deleted as there are USERHISTORY entries for this user", vbExclamation
            Exit Sub
        End If
        
' Check to see if there are any INTRODUCER records for this user
        sCommand = "SELECT COUNT(1) FROM INTRODUCER WHERE USERID = '" & sUserID & "'"
        Set rs = g_clsDataAccess.ExecuteCommand(sCommand)
        If rs(0) > 0 Then
            MsgBox "User '" & sUserID & "' cannot be deleted as there are INTRODUCER entries for this user", vbExclamation
            Exit Sub
        End If
' TW 02/01/2007 EP2_640 End
        
        'User-to-Contact Mapping.
        Set clsUserContactTable = New OrgUserContactTable
        Set clsTableAccess = clsUserContactTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        
        'Load the contact mapping record and obtain the contact GUID foreign key.
        clsTableAccess.GetTableData
        
        'Only attempt to delete the contact records if we have a mapping.
        If clsTableAccess.RecordCount > 0 Then
            vContactGUID = clsUserContactTable.GetContactDetailsGUID
        
            'Delete the mapping record.
            clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
            clsTableAccess.DeleteRecords
        
            'Create a second key match collection for use with the contact GUID.
            Set colContactMatchValues = New Collection
            colContactMatchValues.Add vContactGUID
        
            'Contact Telephone Details.
            Set clsTableAccess = New ContactDetailsTelephoneTable
            clsTableAccess.SetKeyMatchValues colContactMatchValues
            clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
            clsTableAccess.DeleteRecords
            
            'Contact Details.
            Set clsTableAccess = New ContactDetailsTable
            clsTableAccess.SetKeyMatchValues colContactMatchValues
            clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
            clsTableAccess.DeleteRecords
        End If
        
' TW 02/01/2007 EP2_640
'        ' User History
'        Dim sCommand As String
'
'        sUserID = colMatchValues(1)
'        sCommand = "DELETE FROM USERHISTORY WHERE USERID = " + g_clsSQLAssistSP.FormatString(sUserID)
'        g_clsDataAccess.ExecuteCommand sCommand, (m_clsTableAccess.GetDatabase)
' TW 02/01/2007 EP2_640 End
        
' TW 02/01/2007 EP2_640
        ' User Audit
        sCommand = "DELETE FROM USERAUDIT WHERE USERID = " + g_clsSQLAssistSP.FormatString(sUserID)
        g_clsDataAccess.ExecuteCommand sCommand, (m_clsTableAccess.GetDatabase)
        
        ' User Role Audit
        sCommand = "DELETE FROM USERROLEAUDIT WHERE USERID = " + g_clsSQLAssistSP.FormatString(sUserID)
        g_clsDataAccess.ExecuteCommand sCommand, (m_clsTableAccess.GetDatabase)
        
        ' User Competency Audit
        sCommand = "DELETE FROM USERCOMPETENCYAUDIT WHERE USERID = " + g_clsSQLAssistSP.FormatString(sUserID)
        g_clsDataAccess.ExecuteCommand sCommand, (m_clsTableAccess.GetDatabase)
' TW 02/01/2007 EP2_640 End
        
        ' User Role
        Set clsTableAccess = New UserRoleTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
        
        ' Organisation User
        Set clsTableAccess = New OrgUserTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
        
        ' Qualification
        Set clsTableAccess = New QualificationTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
        
        ' User Competency
        Set clsTableAccess = New UserCompetencyTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
        
        ' Password
        Set clsTableAccess = New PasswordTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
        
        'If security exists, then remove all access.
        Set iManagement = g_clsSecurityMgr
        iManagement.RemoveAllAccess sUserID, AgentTypeUser
        
        ' Omiga User
        Set clsTableAccess = New OmigaUserTable
        clsTableAccess.SetKeyMatchValues colMatchValues
        clsTableAccess.SetDatabase m_clsTableAccess.GetDatabase
        clsTableAccess.DeleteRecords
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "Delete User"
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

Public Function DoesNameExist(sName As String)
    Dim colMatchValues As Collection
    Dim colMatchKeys As Collection
    Dim rs As ADODB.Recordset
    Set colMatchValues = New Collection
    Set colMatchKeys = New Collection
        
    colMatchValues.Add sName
    colMatchKeys.Add "NameField"
    
    Set rs = m_clsTableAccess.MatchRecords(colMatchKeys, colMatchValues)

    If rs.RecordCount > 0 Then
        MsgBox "Duplicate"
    End If

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

