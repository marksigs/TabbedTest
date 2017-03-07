VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OmigaUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : OmigaUser
' Description   : Handles all processing to do with the User tab of the Omiga
'                 User form.
' History       :
'
' Prog      Date        Description
' STB       28/11/2001  SYS2912 - Removed the frmMainUserDetails screen and
'                       ported its functionality to the frmEditUser screen.
' STB       06/03/2002  SYS4229 - Added contact telephone number functionality.
' STB       04/04/2002  SYS4229 - Bug when adding user after above change.
' CL        28/05/2002  SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS
' LD        20/10/2005  MAR242 NT Authentication
' GHun      21/10/2005  MAR264 Added Workgroup user
' RF        18/01/2006  MAR1000 Error on updating user when authentication mode is
'                               Windows Authentication - enable change password for
'                               system user.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Text Edit constants.
Private Const USER_ID                 As Long = 0
Private Const USER_INITIALS           As Long = 1
Private Const USER_FIRST_FORENAME     As Long = 2
Private Const USER_SURNAME            As Long = 3
Private Const USER_DATE_OF_BIRTH      As Long = 4
Private Const USER_ACTIVE_FROM        As Long = 5
Private Const USER_ACTIVE_TO          As Long = 6
Private Const USER_INTERMEDIARY_ID    As Long = 7
Private Const USER_MOTHER_MAIDEN_NAME As Long = 8
Private Const USER_PASSWORD           As Long = 9
Private Const USER_EMERGENCY_PASSWORD As Long = 10
Private Const USER_TEL_COUNTRYCODE    As Long = 11
Private Const USER_TEL_AREACODE       As Long = 12
Private Const USER_TEL_TELENUMBER     As Long = 13
Private Const USER_TEL_EXTENSION      As Long = 14

'The telephone usage validation type.
Private Const DEFAULT_TELEPHONE_USAGE As String = "W"

'Local Table Classes.
Private m_clsPassword As PasswordTable
Private m_clsUserTable As OmigaUserTable
Private m_clsOrgUserTable As OrgUserTable

'Tables required to hold a phone number.
Private m_clsContactTable As ContactDetailsTable
Private m_clsTelephoneTable As ContactDetailsTelephoneTable
Private m_clsUserContactTable As OrgUserContactTable

'Store the password at module-level when the class intialises. When save is
'called, we'll check to see if its been altered.
Private m_sPassword As String

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The current record's user ID.
Private m_sUserID As String

Private Type TypeUserAuditRecord
    UserId As String
    AuditDate As Date
    ChangeUser As String
    AccessType As Integer
    PrevAccessType As Integer
    PrevActiveFrom As String
    ActiveFrom As String
    PrevActiveTo As String
    ActiveTo As String
End Type

Private UserAuditRecord As TypeUserAuditRecord

Private Sub DoUserAuditTableUpdate()
    
    On Error GoTo Failed

    Dim sSQL As String
    
    sSQL = ""
    sSQL = "INSERT INTO USERAUDIT (USERID, AUDITDATE, CHANGEUSER,ACCESSTYPE,ACTIVEFROM,ACTIVETO,PREVACCESSTYPE,PREVACTIVEFROM,PREVACTIVETO)  " & _
            " VALUES ('" & UserAuditRecord.UserId & "', " & _
            g_clsSQLAssistSP.GetSystemDate & ", '" & _
            UserAuditRecord.ChangeUser & "', " & _
            UserAuditRecord.AccessType & ", "

    If Len(UserAuditRecord.ActiveFrom) > 0 Then
        sSQL = sSQL & g_clsSQLAssistSP.FormatDate(UserAuditRecord.ActiveFrom) & " , "
    Else
        sSQL = sSQL & " Null " & ","
    End If
    
    If Len(UserAuditRecord.ActiveTo) > 0 Then
        sSQL = sSQL & g_clsSQLAssistSP.FormatDate(UserAuditRecord.ActiveTo) & " , "
    Else
        sSQL = sSQL & " Null " & ","
    End If
    
    sSQL = sSQL & UserAuditRecord.PrevAccessType & " , "
    
    If Len(UserAuditRecord.PrevActiveFrom) > 0 Then
        sSQL = sSQL & g_clsSQLAssistSP.FormatDate(UserAuditRecord.PrevActiveFrom) & " , "
    Else
        sSQL = sSQL & " Null " & ","
    End If
    
    If Len(UserAuditRecord.PrevActiveTo) > 0 Then
        sSQL = sSQL & g_clsSQLAssistSP.FormatDate(UserAuditRecord.PrevActiveTo)
    Else
        sSQL = sSQL & " Null "
    End If
    
    sSQL = sSQL & " ) "
    
    g_clsDataAccess.ExecuteCommand sSQL
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetUserAuditTableData()
    
    UserAuditRecord.UserId = frmEditUser.txtUserDetails(USER_ID).Text
    UserAuditRecord.ChangeUser = g_sSupervisorUser
    UserAuditRecord.AccessType = m_clsUserTable.GetAccessType()
    UserAuditRecord.ActiveFrom = m_clsUserTable.GetActiveFrom()
    UserAuditRecord.ActiveTo = m_clsUserTable.GetActiveTo()
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated accoding to the
'                 underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    On Error GoTo Failed
    
    'Store the add/edit state of the form.
    m_bIsEdit = bIsEdit
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Store the password at module-level.
    m_sPassword = frmEditUser.txtPassword.Text
    
    'Copy the password into the confirm controls.
    frmEditUser.txtPasswordConfirm.Text = frmEditUser.txtPassword.Text
           
    'Populate the screen combo's.
    g_clsFormProcessing.PopulateCombo "Title", frmEditUser.cboTitle
    g_clsFormProcessing.PopulateCombo "UserAccessType", frmEditUser.cboAccessType
    
    'The working hours combo uses a different method for populating (for some reason...).
    PopulateWorkingHours

    'Load the related ContactTelephoneDetails record.
    LoadContactTelephoneRecord
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateWorkingHours()
    
    Dim colIDS As New Collection
    Dim colValues As New Collection
    Dim clsWorkingHours As New WorkingHourTypeTable
    
    On Error GoTo Failed
    
    'Obtain the values and ids collections from the WORKINGHOURSTYPE table.
    clsWorkingHours.GetWorkingHourTypesAsCollection colValues, colIDS
    
    'Populate the combo from those collections.
    frmEditUser.cboWorkingHours.SetListTextFromCollection colValues, colIDS
    
    'Select the first item if there were any found.
    If frmEditUser.cboWorkingHours.ListCount > 0 Then
        frmEditUser.cboWorkingHours.ListIndex = 0
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function SaveUserDetails() As Boolean
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    Dim colKeys As Collection
    
    'Store the user id at module level (from the screen control).
    m_sUserID = frmEditUser.txtUserDetails(USER_ID).Text
    
    'If we're adding, then set the colKeys now (otherwise promotions will fail).
    If m_bIsEdit = False Then
        'Create the collection now.
        Set colKeys = New Collection
        
        'Add the Key field ID into the collection.
        colKeys.Add m_sUserID
        
        'Set the collection against the underlying OmigaUser table.
        TableAccess(m_clsUserTable).SetKeyMatchValues colKeys
    End If
    
    'If there was a user id.
    If Len(m_sUserID) > 0 Then
        'Set the return code to True.
        bRet = True
        
        ' User ID
        m_clsUserTable.SetUserID m_sUserID
        
        ' Access Type
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboAccessType, vTmp, GET_CONTROL_VALUE
         m_clsUserTable.SetAccessType CStr(vTmp)
        
        ' Active From
        g_clsFormProcessing.HandleDate frmEditUser.txtUserDetails(USER_ACTIVE_FROM), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetActiveFrom vTmp
        
        ' Active To
        g_clsFormProcessing.HandleDate frmEditUser.txtUserDetails(USER_ACTIVE_TO), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetActiveTo vTmp
        
        ' Notes
        m_clsUserTable.SetNotes frmEditUser.TxtNotes.Text
        
        ' Working Hour Type
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboWorkingHours, vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetWorkingHourType CStr(vTmp)
        
        ' BMIDS678  Credit Check Access
        g_clsFormProcessing.HandleRadioButtons frmEditUser.optCreditCheckAccess(OPT_YES), frmEditUser.optCreditCheckAccess(OPT_NO), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetCreditCheckAccess vTmp
        
        'MAR264 GHun Workgroup user
        g_clsFormProcessing.HandleRadioButtons frmEditUser.optWorkgroupUser(OPT_YES), frmEditUser.optWorkgroupUser(OPT_NO), vTmp, GET_CONTROL_VALUE
        m_clsUserTable.SetWorkgroupUser vTmp
        'MAR264 End
    End If
        
    SaveUserDetails = bRet
    
    Exit Function

Failed:
    MsgBox "SaveUserDetails: Error is " + Err.DESCRIPTION
    SaveUserDetails = False
End Function

Private Function GetUserDetails() As Boolean
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    If Len(m_sUserID) > 0 Then
        bRet = True
        frmEditUser.txtUserDetails(USER_ID).Text = m_sUserID
        
        ' Access Type
        vTmp = m_clsUserTable.GetAccessType()
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboAccessType, vTmp, SET_CONTROL_VALUE
        
        ' Active From
        frmEditUser.txtUserDetails(USER_ACTIVE_FROM).Text = m_clsUserTable.GetActiveFrom()
        
        ' Active To
        frmEditUser.txtUserDetails(USER_ACTIVE_TO).Text = m_clsUserTable.GetActiveTo()
        
        ' Notes
        frmEditUser.TxtNotes.Text = m_clsUserTable.GetNotes()
                
        ' Working Hour Type
        vTmp = m_clsUserTable.GetWorkingHourType()
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboWorkingHours, vTmp, SET_CONTROL_VALUE

        ' BMIDS678  Credit Check Access
        vTmp = m_clsUserTable.GetCreditCheckAccess()
        g_clsFormProcessing.HandleRadioButtons frmEditUser.optCreditCheckAccess(OPT_YES), frmEditUser.optCreditCheckAccess(OPT_NO), vTmp, SET_CONTROL_VALUE

        'MAR264 GHun Workgroup user
        vTmp = m_clsUserTable.GetWorkgroupUser()
        g_clsFormProcessing.HandleRadioButtons frmEditUser.optWorkgroupUser(OPT_YES), frmEditUser.optWorkgroupUser(OPT_NO), vTmp, SET_CONTROL_VALUE
        'MAR264 End
        
    End If
    
    'MV - 16/12/2002 - BM0085
    UserAuditRecord.PrevAccessType = m_clsUserTable.GetAccessType()
    UserAuditRecord.PrevActiveFrom = m_clsUserTable.GetActiveFrom()
    UserAuditRecord.PrevActiveTo = m_clsUserTable.GetActiveTo()
    
    GetUserDetails = bRet
    
    Exit Function
    
Failed:
    MsgBox "GetUserDetails: Error is " + Err.DESCRIPTION
    GetUserDetails = False
End Function

Private Function SaveOrgUserDetails() As Boolean
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    bRet = True
    
    If Len(m_sUserID) > 0 Then
        'User ID.
        m_clsOrgUserTable.SetUserID m_sUserID
        
        ' Date of Birth
        g_clsFormProcessing.HandleDate frmEditUser.txtUserDetails(USER_DATE_OF_BIRTH), vTmp, GET_CONTROL_VALUE
        m_clsOrgUserTable.SetDateOfBirth vTmp
        
        ' Emergency Password
        m_clsOrgUserTable.SetEmergencyPassword frmEditUser.txtUserDetails(USER_EMERGENCY_PASSWORD).Text
        
        ' Mother Maiden Name
        m_clsOrgUserTable.SetMotherMaidenName frmEditUser.txtUserDetails(USER_MOTHER_MAIDEN_NAME).Text
        
        ' User Forename
        m_clsOrgUserTable.SetForeName frmEditUser.txtUserDetails(USER_FIRST_FORENAME).Text
        
        ' User Initials
        m_clsOrgUserTable.SetInitials frmEditUser.txtUserDetails(USER_INITIALS).Text
        
        ' User Surname
        m_clsOrgUserTable.SetUserSurname frmEditUser.txtUserDetails(USER_SURNAME).Text
        
        ' Title
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboTitle, vTmp, GET_CONTROL_VALUE
        m_clsOrgUserTable.SetUserTitle CStr(vTmp)
        
        'Country Code.
        m_clsTelephoneTable.SetCOUNTRYCODE frmEditUser.txtUserDetails(USER_TEL_COUNTRYCODE).Text
        
        'Area Code.
        m_clsTelephoneTable.SetAREA_CODE frmEditUser.txtUserDetails(USER_TEL_AREACODE).Text
        
        'Telephone Number.
        m_clsTelephoneTable.SetTELEPHONE frmEditUser.txtUserDetails(USER_TEL_TELENUMBER).Text
        
        'Extension Number.
        m_clsTelephoneTable.SetTELEPHONE_EXT frmEditUser.txtUserDetails(USER_TEL_EXTENSION).Text
    End If
    
    SaveOrgUserDetails = bRet
    
    Exit Function
    
Failed:
    MsgBox "SaveOrgUserDetails: Error is " + Err.DESCRIPTION
    SaveOrgUserDetails = False
End Function

Private Sub SavePassword()
    
    Dim sNewPassword As String
    Dim dMaxPassword As Double
    Dim sPasswordEncrypt As String
    Dim clsTableAccess As TableAccess
    
On Error GoTo Failed
    
    g_sFunctionName = "SavePassword"

    'Obtain a local working copy of the new password.
    sNewPassword = frmEditUser.txtPassword.Text
    
    If Len(m_sUserID) > 0 Then
        dMaxPassword = g_clsGlobalParameter.FindAmount("MaximumPasswords")

        ' How many passwords are there for this user?
        Set clsTableAccess = m_clsPassword

        If dMaxPassword > clsTableAccess.RecordCount Then
            g_clsFormProcessing.CreateNewRecord m_clsPassword
        Else
            ' Need to update the oldest one
            m_clsPassword.FindPassword m_sUserID
        End If
        
        m_clsPassword.SetUserID m_sUserID
        sPasswordEncrypt = Encrypt(sNewPassword)
        
        m_clsPassword.SetPassword sPasswordEncrypt
        m_clsPassword.SetCreationDate
        m_clsUserTable.SetChangePassword
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SavePassword - User ID is empty"
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function GetOrgUserDetails() As Boolean
On Error GoTo Failed

    Dim bRet As Boolean
    Dim vTmp As Variant
    
    bRet = False
    
    If Len(m_sUserID) > 0 Then
        bRet = True
        
        'Date of Birth.
        frmEditUser.txtUserDetails(USER_DATE_OF_BIRTH).Text = m_clsOrgUserTable.GetDateOfBirth()
        
        'Emergency Password.
        frmEditUser.txtUserDetails(USER_EMERGENCY_PASSWORD).Text = m_clsOrgUserTable.GetEmergencyPassword()
        
        'Mother Maiden Name.
        frmEditUser.txtUserDetails(USER_MOTHER_MAIDEN_NAME).Text = m_clsOrgUserTable.GetMotherMaidenName()
        
        'User Forename.
        frmEditUser.txtUserDetails(USER_FIRST_FORENAME).Text = m_clsOrgUserTable.GetForeName()
        
        'User Initials.
        frmEditUser.txtUserDetails(USER_INITIALS).Text = m_clsOrgUserTable.GetInitials()
        
        'User Surname.
        frmEditUser.txtUserDetails(USER_SURNAME).Text = m_clsOrgUserTable.GetSurName()
        
        'Title.
        vTmp = m_clsOrgUserTable.GetTitle()
        g_clsFormProcessing.HandleComboExtra frmEditUser.cboTitle, vTmp, SET_CONTROL_VALUE
        
        'Country Code.
        frmEditUser.txtUserDetails(USER_TEL_COUNTRYCODE).Text = m_clsTelephoneTable.GetCOUNTRYCODE
        
        'Area Code.
        frmEditUser.txtUserDetails(USER_TEL_AREACODE).Text = m_clsTelephoneTable.GetAREA_CODE
        
        'Telephone Number.
        frmEditUser.txtUserDetails(USER_TEL_TELENUMBER).Text = m_clsTelephoneTable.GetTELEPHONE
        
        'Extensions Number.
        frmEditUser.txtUserDetails(USER_TEL_EXTENSION).Text = m_clsTelephoneTable.GetTELEPHONE_EXT
    End If
    
    GetOrgUserDetails = bRet
    Exit Function

Failed:
    MsgBox "GetOrgUserDetails: Error is " + Err.DESCRIPTION
    GetOrgUserDetails = False
End Function

Public Function SetScreenFields()
    
    Dim bRet As Boolean
    
On Error GoTo Failed
    
    bRet = False

    'Only attempt to populate the screen if we have a user id.
    If Len(m_sUserID) > 0 Then
        bRet = GetUserDetails()
        
        If bRet = True Then
            bRet = GetOrgUserDetails()
        End If
    Else
        MsgBox "SetScreenFields: UserID is empty"
    End If
    
    'Indicate success or failure.
    SetScreenFields = bRet
    
    Exit Function
    
Failed:
    MsgBox "SetScreenFields: Error is " + Err.DESCRIPTION
    SetScreenFields = False
End Function

Public Sub SaveScreenData()
On Error GoTo Failed

    SaveUserDetails
    
    'MV - 16/12/2002 - BM0085
    SetUserAuditTableData
    
    SaveOrgUserDetails
    
    'If the password has altered, then save it, otherwise leave it be.
    If m_sPassword <> frmEditUser.txtPassword.Text Then
        ' RF 18/01/2006 MAR1000 Start
        If frmEditUser.GetSecurityCredentialsType = Omiga Or _
            frmEditUser.IsSystemUser() Then
        ' RF 18/01/2006 MAR1000 End
            
            SavePassword
            
        End If
    
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DoUpdates()
    
    On Error GoTo Failed
    
    'Update the OmigaUser table.
    TableAccess(m_clsUserTable).Update
    
    'Update the OrganisationUser table.
    TableAccess(m_clsOrgUserTable).Update

    'Update the password table
    ' RF 18/01/2006 MAR1000 Start
    If TableAccess(m_clsPassword).RecordCount > 0 Then
        TableAccess(m_clsPassword).Update
    End If
    ' RF 18/01/2006 MAR1000 End

    'Update the contact details and telephone tables (records may be cropped).
    SaveContactTelephoneRecord
    
    'MV - 16/12/2002 - BM0085
    DoUserAuditTableUpdate
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Public Function GetUserID() As Variant
    GetUserID = m_sUserID
End Function

Friend Sub SetUserID(sUserID As String)

    On Error GoTo Failed
    
    If Len(sUserID) > 0 Then
        m_sUserID = sUserID
    Else
        g_clsErrorHandling.RaiseError ErrUserIDIsEmpty
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetTableClass(ByRef clsTableAccess As TableAccess)
    Set m_clsUserTable = clsTableAccess
End Sub

Public Function ValidatePasswordFormat(ByVal vstrPassword As String) As Boolean

    On Error GoTo ValidatePasswordFormatVbErr
    
    '------------------------------------------------------------------------------------------
    ' check password length
    '------------------------------------------------------------------------------------------
    Dim bRet As Boolean
    Dim dblMinLen As Double, dblMaxLen As Double
    Dim sError As String
        
    bRet = True
    dblMinLen = g_clsGlobalParameter.FindAmount("PasswordMinimumLength")
        
    If Len(vstrPassword) < dblMinLen Then
        sError = "Password too short; minimum length is " & CStr(dblMinLen)
        Err.Raise SUPERVISOR_INVALID_PASSWORD
    End If
    
    dblMaxLen = g_clsGlobalParameter.FindAmount("PasswordMaximumLength")
        
    If Len(vstrPassword) > dblMaxLen Then
        sError = "Password too long; maximum length is " & CStr(dblMaxLen)
        Err.Raise SUPERVISOR_INVALID_PASSWORD
    End If
    
    '------------------------------------------------------------------------------------------
    ' check alpha and numeric chars
    '------------------------------------------------------------------------------------------
    
    Dim blnAlphaAndNumeric As Boolean
    
    blnAlphaAndNumeric = g_clsGlobalParameter.FindBoolean("PasswordAlphaNumeric")
        
    If blnAlphaAndNumeric = True Then
        If g_clsGeneralAssist.IsAlpha(vstrPassword) Then
            sError = "Password may not be all alpha characters"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
        
        If g_clsGeneralAssist.IsDigits(vstrPassword) = True Then
            sError = "Password may not be all digits"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    '------------------------------------------------------------------------------------------
    ' check for duplicated chars
    '------------------------------------------------------------------------------------------
    
    Dim blnDuplicatedChars As Boolean
    
    blnDuplicatedChars = g_clsGlobalParameter.FindBoolean("PasswordCharacterDuplication")
    
    If blnDuplicatedChars = True Then
        If g_clsGeneralAssist.HasDuplicatedChars(vstrPassword, False) = True Then
            sError = "Password may not contain duplicated characters"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    '------------------------------------------------------------------------------------------
    ' check the char set
    '------------------------------------------------------------------------------------------
    
    Dim blnUpperAlpha, blnLowerAlpha, blnNumeric, blnSpecialChars As Boolean
    
    blnUpperAlpha = g_clsGlobalParameter.FindBoolean("PasswordUpperAlpha")
        
    If blnUpperAlpha = False Then
        If g_clsGeneralAssist.ContainsUpperAlpha(vstrPassword) = True Then
            sError = "Password may not contain upper case alpha characters"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    blnLowerAlpha = g_clsGlobalParameter.FindBoolean("PasswordLowerAlpha")
        
    If blnLowerAlpha = False Then
        If g_clsGeneralAssist.ContainsLowerAlpha(vstrPassword) = True Then
            sError = "Password may not contain lower case alpha characters"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    blnNumeric = g_clsGlobalParameter.FindBoolean("PasswordNumeric")
        
    If blnNumeric = False Then
        If g_clsGeneralAssist.ContainsDigits(vstrPassword) = True Then
            sError = "Password may not contain digits"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    blnSpecialChars = g_clsGlobalParameter.FindBoolean("PasswordSpecialCharacters")
        
    If blnSpecialChars = False Then
        If g_clsGeneralAssist.ContainsSpecialChars(vstrPassword) = True Then
            sError = "Password may not contain special characters"
            Err.Raise SUPERVISOR_INVALID_PASSWORD
        End If
    End If
    
    ValidatePasswordFormat = bRet
    Exit Function

ValidatePasswordFormatVbErr:
    If Len(sError) > 0 Then
        MsgBox sError, vbCritical
    Else
        MsgBox "ValidatePassword - Error is: " + Err.DESCRIPTION

    End If
    frmEditUser.txtPassword.SetFocus
    ValidatePasswordFormat = False
End Function

Private Function ValidatePassword() As Boolean
    
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    Dim rs As ADODB.Recordset
    Dim sUserID As String
    Dim sPassword As String
    Dim sNewPassword As String
    Dim sPasswordConfirm As String
    
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsPassword
    bRet = True
    
    ' passwords are not mandatory when using NT authentication.
    ' RF 18/01/2006 MAR1000 Start
    'If frmEditUser.txtPassword.Enabled = True Then
    If frmEditUser.GetSecurityCredentialsType = Omiga Then
    ' RF 18/01/2006 MAR1000 End
    
        sUserID = frmEditUser.txtUserDetails(USER_ID).Text
        
        sNewPassword = frmEditUser.txtPassword.Text
        
        sPasswordConfirm = frmEditUser.txtPasswordConfirm.Text
        
        If Len(sUserID) > 0 Then
            If Len(sNewPassword) > 0 Then
                ' Make sure the confirm password is the same as the password
                
                bRet = ValidatePasswordFormat(sNewPassword)
                
                If bRet = True Then
                    If sNewPassword = sPasswordConfirm Then
                        colValues.Add sUserID
                        
                        clsTableAccess.SetKeyMatchValues colValues
                        
                        Set rs = clsTableAccess.GetTableData()
                
                        If Not rs Is Nothing Then
                            ' Now see if the password exists already
                            If rs.RecordCount > 0 Then
                                rs.MoveFirst
                                
                                While Not rs.EOF And bRet = True
                                    sPassword = m_clsPassword.GetPassword()
                                                                        
                                    If sPassword = Encrypt(sNewPassword) Then
                                        bRet = False
                                    Else
                                        rs.MoveNext
                                    End If
                                Wend
                                
                                If bRet = False Then
                                    frmEditUser.txtPassword.SetFocus
                                    g_clsErrorHandling.RaiseError errGeneralError, "Password has previously been used for user <" + sUserID + "> - please enter a unique password"
                                End If
                            End If
                        Else
                            g_clsErrorHandling.RaiseError errGeneralError, "ValidatePassword - Unable to populate Password table"
                        End If
                    Else
                        frmEditUser.txtPasswordConfirm.SetFocus
                        g_clsErrorHandling.RaiseError errGeneralError, "Password and Confirm Password must be the same"
                    End If
                End If
            Else
                g_clsErrorHandling.RaiseError errGeneralError, "ValidatePassword - Password is empty"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "ValidatePassword - UserID is empty"
        End If
    End If
    ValidatePassword = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.DisplayError
    ValidatePassword = False
End Function

Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    Dim sUserID As String
    Dim colKeys As Collection
    
    bRet = True
    bRet = ValidatePassword()

    'Ensure (if adding) that the record is unique.
    If (bRet = True) And (m_bIsEdit = False) Then
        'Create a keys collection.
        Set colKeys = New Collection
        
        'Add the user id into the collection.
        sUserID = frmEditUser.txtUserDetails(USER_ID).Text
        colKeys.Add sUserID
    
        bRet = Not TableAccess(m_clsUserTable).DoesRecordExist(colKeys)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "User ID " + sUserID + " already exists"
            bRet = False
        End If
    End If

    If bRet = True Then
        'Ensure the from and to dates are valid.
        bRet = g_clsValidation.ValidateActiveFromTo(frmEditUser.txtUserDetails(USER_ACTIVE_FROM), frmEditUser.txtUserDetails(USER_ACTIVE_TO))
    End If

    ValidateScreenData = bRet
    
End Function

Public Sub SetAddState()
    
    'Create a blank records for both OmigaUser and OrganisationUser.
    g_clsFormProcessing.CreateNewRecord m_clsUserTable
    g_clsFormProcessing.CreateNewRecord m_clsOrgUserTable

End Sub

Public Sub SetEditState()

    Dim dMinLen As Double
    Dim rs As ADODB.Recordset
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    'Disable the UserID control if we're editing an existing record as this is
    'used as a primary key.
    frmEditUser.txtUserDetails(USER_ID).Enabled = False

    'If editing, we want to just put some dummy characters into the password field.
    dMinLen = g_clsGlobalParameter.FindAmount("PasswordMinimumLength")

    '* means nothing - just an arbitary value.
    frmEditUser.txtPassword.Text = String(dMinLen, "*")
        
    'Load the desired record into the OmigaUser table object. Normally we'd
    'have set the keys collection against this table object. However, with
    'Users, the underlying table has already been created by the parent
    'form and passed to us with the keys collection already set.
    Set clsTableAccess = m_clsUserTable
    Set rs = clsTableAccess.GetTableData()
    
    'Ensure a record was loaded.
    clsTableAccess.ValidateData
    
    If clsTableAccess.RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errRecordNotFound, "OmigaUser"
    End If
    
    'Get a copy of the keys collection used to load record.
    Set colValues = clsTableAccess.GetKeyMatchValues()
    
    'Store the loaded User ID at module level.
    m_sUserID = m_clsUserTable.GetUserID()
            
    'Set the keys collection used for the OmigaUser table against the
    'OrganisationUser table.
    Set clsTableAccess = m_clsOrgUserTable
    clsTableAccess.SetKeyMatchValues colValues

    'Load a record into that table and get a recordset reference.
    Set rs = clsTableAccess.GetTableData()
    
    'Ensure a record was loaded.
    clsTableAccess.ValidateData
        
    If clsTableAccess.RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errRecordNotFound, "OrganisationUser"
    End If
    
End Sub

Private Sub Class_Initialize()
    
    Set m_clsPassword = New PasswordTable
    Set m_clsUserTable = New OmigaUserTable
    Set m_clsOrgUserTable = New OrgUserTable
    Set m_clsContactTable = New ContactDetailsTable
    Set m_clsTelephoneTable = New ContactDetailsTelephoneTable
    Set m_clsUserContactTable = New OrgUserContactTable
    
End Sub

Private Sub Class_Terminate()
    
    Set m_clsPassword = Nothing
    Set m_clsUserTable = Nothing
    Set m_clsOrgUserTable = Nothing
    Set m_clsContactTable = Nothing
    Set m_clsTelephoneTable = Nothing
    Set m_clsUserContactTable = Nothing
    
End Sub

Private Sub LoadContactTelephoneRecord()

    Dim lValueID As Long
    Dim clsGUID As GuidAssist
    Dim vContactGUID As Variant
    Dim colKeyMatchValues As Collection
    Dim clsComboTable As ComboValidationTable
    
    'USER-TO-CONTACT MAPPING RECORD.
    
    'Create a UserID keys collection.
    Set colKeyMatchValues = New Collection
    colKeyMatchValues.Add m_sUserID
    
    'Set the keys collection against the mapping table.
    TableAccess(m_clsUserContactTable).SetKeyMatchValues colKeyMatchValues
    
    'Load any mapping record.
    If m_bIsEdit = True Then
        TableAccess(m_clsUserContactTable).GetTableData POPULATE_KEYS
    Else
        TableAccess(m_clsUserContactTable).GetTableData POPULATE_EMPTY
    End If
            
    'If no record was found, create one (and generate a contact GUID).
    If TableAccess(m_clsUserContactTable).RecordCount = 0 Then
        'Create a blank mapping record.
        g_clsFormProcessing.CreateNewRecord m_clsUserContactTable
        
        'Create a helper class to generate GUIDs.
        Set clsGUID = New GuidAssist
        
        'Set the key field values.
        If m_bIsEdit = True Then
            m_clsUserContactTable.SetUserID m_sUserID
        End If
        
        m_clsUserContactTable.SetContactDetailsGUID CVar(clsGUID.CreateGUID())
    
    'If too many records were returned, moan.
    ElseIf TableAccess(m_clsUserContactTable).RecordCount > 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Too many contact-mapping records returned for user"
    End If
    
    'CONTACT RECORD.
    
    'Get a copy of the Contact GUID.
    vContactGUID = m_clsUserContactTable.GetContactDetailsGUID
    
    'Create a contact GUID keys collection.
    Set colKeyMatchValues = New Collection
    colKeyMatchValues.Add vContactGUID
    
    'Set the keys collection against the contact table.
    TableAccess(m_clsContactTable).SetKeyMatchValues colKeyMatchValues
    
    'Load the related contact details record.
    If m_bIsEdit = True Then
        TableAccess(m_clsContactTable).GetTableData POPULATE_KEYS
    Else
        TableAccess(m_clsContactTable).GetTableData POPULATE_EMPTY
    End If
    
    'If a record was not found, create one (and set the new GUID).
    If TableAccess(m_clsContactTable).RecordCount = 0 Then
        'Create a blank contact record.
        g_clsFormProcessing.CreateNewRecord m_clsContactTable
        
        'Set the key field values.
        m_clsContactTable.SetContactDetailsGUID vContactGUID
    
    'If too many records were returned, moan.
    ElseIf TableAccess(m_clsContactTable).RecordCount > 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Too many contact records returned for user"
    End If
    
    'CONTACT TELEPHONE RECORD.
    
    'Set a keys collection against the telephone table.
    TableAccess(m_clsTelephoneTable).SetKeyMatchValues colKeyMatchValues
    
    'Load the related telephone record.
    If m_bIsEdit Then
        TableAccess(m_clsTelephoneTable).GetTableData POPULATE_KEYS
    Else
        TableAccess(m_clsTelephoneTable).GetTableData POPULATE_EMPTY
    End If
    
    'If a record was not found, create one (and set the new GUID).
    If TableAccess(m_clsTelephoneTable).RecordCount = 0 Then
        'Create a help table to obtain the a combo ValueID.
        Set clsComboTable = New ComboValidationTable
        
        'The Usage validation type is always W(ork). Obtain the ValueID for it.
        lValueID = clsComboTable.GetValueIDForValidationType(DEFAULT_TELEPHONE_USAGE, "ContactTelephoneUsage")
                
        'Create a blank telephone record.
        g_clsFormProcessing.CreateNewRecord m_clsTelephoneTable
        
        'Set the key field values.
        m_clsTelephoneTable.SetCONTACTDETAILSTELEPHONEGUID vContactGUID
        m_clsTelephoneTable.SetTELEPHONE_SEQ_NUMBER 1
        m_clsTelephoneTable.SetType lValueID
    End If

End Sub

Private Sub SaveContactTelephoneRecord()
        
    If ShouldCropContactRecord() Then
        'Delete the contact and telephone records.
        TableAccess(m_clsTelephoneTable).DeleteAllRows
        TableAccess(m_clsContactTable).DeleteAllRows
        TableAccess(m_clsUserContactTable).DeleteAllRows
        
        'Update the tables child-then-parent.
        TableAccess(m_clsUserContactTable).Update
        TableAccess(m_clsTelephoneTable).Update
        TableAccess(m_clsContactTable).Update
    Else
        'If we're adding, then only now do we know the UserID to link the
        'contact record to.
        If m_bIsEdit = False Then
            m_clsUserContactTable.SetUserID m_sUserID
        End If
        
        'Update the tables parent-then-child.
        TableAccess(m_clsContactTable).Update
        TableAccess(m_clsTelephoneTable).Update
        TableAccess(m_clsUserContactTable).Update
    End If
    
End Sub

Private Function ShouldCropContactRecord() As Boolean

    Dim bFound As Boolean
        
    'Set bFound to true if any values have been entered.
    bFound = (frmEditUser.txtUserDetails(USER_TEL_COUNTRYCODE).Text <> "") Or bFound
    bFound = (frmEditUser.txtUserDetails(USER_TEL_AREACODE).Text <> "") Or bFound
    bFound = (frmEditUser.txtUserDetails(USER_TEL_TELENUMBER).Text <> "") Or bFound
    bFound = (frmEditUser.txtUserDetails(USER_TEL_EXTENSION).Text <> "") Or bFound

    ShouldCropContactRecord = (Not bFound)

End Function



