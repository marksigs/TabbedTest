VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : UserRole
' Description   : Handle all controls relevant to the User Roles tab.
' History       :
'
' Prog      Date        Description
' STB       28/11/2001  SYS2912 - Removed the frmMainUserDetails screen and
'                       ported its functionality to the frmEditUser screen.
' STB       07/03/2002  SYS4208 - Placed data verification check in.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM
' TW        13/04/2007  EP2_2375 - Error when trying to edit a user's "User Profile"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Column indexes.
Private Const FROM_COLUMN_IDX As Integer = 3
Private Const TO_COLUMN_IDX   As Integer = 4

'Collections to store various items.
Private m_colComboRole As New Collection
Private m_colComboUnits As New Collection
Private m_colComboRoleIDS As New Collection

'The underlying user role table object.
Private m_clsUserRole As New UserRoleTable

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The current record's user ID.
Private m_sUserID As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if the form is in an add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Speicific History:
'
'Prog   Date        AQR     Description
'MV     27/12/2002  BM0085  Modified DoUpdates(),Initialise()
'                           New DoUpdateUserRoleauditTableData(),SetInitialUserRoleAuditData();SetUserRoleAuditTableData()
'MV     07/01/2003  BM0085  Modified DoUpdateUserRoleauditTableData(),DoUpdates()
'MV     20/01/2003  BM0085  Amended DoUpdateUserRoleAuditTableData
'MV     28/01/2003  BM0303  Amended DoUpdates
'MV     05/03/2003  BM0321  Amended ValidateScreenData()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type TypeUserRoleAuditRecord

    UserId As String
    UnitId As String
    AuditDate As String
    ChangeUser As String
    PrevRoleActiveFrom As String
    PrevRoleActiveTo As String
    PrevRole As String
    RoleActiveFrom As String
    RoleActiveTo As String
    Role As String
    RowNum As String
    Flag As String

End Type

Private UserRoleAuditRecord As TypeUserRoleAuditRecord
Private InitialRowArray() As TypeUserRoleAuditRecord
Private FinalRowArray() As TypeUserRoleAuditRecord
Private m_clsUserRoleTable As UserRoleTable

Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_nCurrentIndex As Long



Public Function CreateUserRoleAudit(ByVal sUnitID As String, ByVal sUserID As String, ByVal sChangeUser As String, _
                                ByVal sPrevRoleActiveFrom As String, ByVal sPrevRoleActiveTo As String, _
                                ByVal sPrevRole As String, ByVal sRoleActiveFrom As String, ByVal sRoleActiveTo As String, _
                                ByVal sRole As String, ByVal sOperation As String) As Boolean
                                
        
    On Error GoTo Failed
    
    Dim sSQL As String
    
    sSQL = ""
    sSQL = sSQL + "INSERT INTO USERROLEAUDIT (UNITID,USERID,AUDITDATE,CHANGEUSER,PREVROLEACTIVEFROM,PREVROLEACTIVETO,PREVROLE,ROLEACTIVEFROM,ROLEACTIVETO,ROLE,OPERATION) "
    sSQL = sSQL + " VALUES ( "
    sSQL = sSQL + "'" & sUnitID & "', "
    sSQL = sSQL + "'" & sUserID & "', "
    sSQL = sSQL + g_clsSQLAssistSP.GetSystemDate & ", "
    sSQL = sSQL + "'" & sChangeUser & "', "
    
    If sPrevRoleActiveFrom = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sPrevRoleActiveFrom & "', "
    End If
    
    If sPrevRoleActiveTo = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sPrevRoleActiveTo & "', "
    End If
    
    If sPrevRole = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + sPrevRole & " , "
    End If
    
    If sRoleActiveFrom = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sRoleActiveFrom & "', "
    End If
    
    If sRoleActiveTo = "" Then
        sSQL = sSQL + "Null" + ","
    Else
        sSQL = sSQL + "'" & sRoleActiveTo & "', "
    End If
    
    If sRole = "" Then
        sSQL = sSQL + "Null" + ", "
    Else
        sSQL = sSQL + sRole & " ,"
    End If
     
    If sOperation = "" Then
        sSQL = sSQL + "Null" + " )"
    Else
        sSQL = sSQL + "'" + sOperation + "' )"
    End If
     
     
    g_clsDataAccess.ExecuteCommand sSQL
    
    CreateUserRoleAudit = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function

Public Sub DoUpdateUserRoleAuditTableData()
    
    On Error GoTo Failed
    
    Dim iInitialRows As Integer
    Dim iFinalRows As Integer
    Dim iCtr  As Integer
    Dim iCompareRows  As Integer
    Dim iTotalNewRecords As Integer
    Dim bRet  As Boolean
    Dim sSQL As String
         
'Stage1:
    
    'Set the Final Grid Data into an type Array
    SetUpFinalUserRoleAuditData
    
    On Error Resume Next
    iInitialRows = UBound(InitialRowArray, 1) - LBound(InitialRowArray, 1)
    iFinalRows = UBound(FinalRowArray, 1) - LBound(FinalRowArray, 1)
    On Error GoTo Failed
    
'Stage2:
    
    'Create an AuditRecord for an Editable Records
    
    If iFinalRows < iInitialRows Then
        iCompareRows = iFinalRows
    Else
        iCompareRows = iInitialRows
    End If
    
    For iCtr = 1 To iCompareRows
        
        If (InitialRowArray(iCtr).PrevRole <> FinalRowArray(iCtr).Role) Or _
        (InitialRowArray(iCtr).PrevRoleActiveTo <> FinalRowArray(iCtr).RoleActiveTo) Or _
        (InitialRowArray(iCtr).PrevRoleActiveFrom <> FinalRowArray(iCtr).RoleActiveFrom) Then
            
            bRet = CreateUserRoleAudit(FinalRowArray(iCtr).UnitId, FinalRowArray(iCtr).UserId, FinalRowArray(iCtr).ChangeUser, _
                    InitialRowArray(iCtr).PrevRoleActiveFrom, InitialRowArray(iCtr).PrevRoleActiveTo, _
                    InitialRowArray(iCtr).PrevRole, FinalRowArray(iCtr).RoleActiveFrom, _
                    FinalRowArray(iCtr).RoleActiveTo, FinalRowArray(iCtr).Role, "U")
            
            If Not bRet Then
                g_clsErrorHandling.DisplayError "Error in Creating Audit Record"
                Exit Sub
            End If
            
            
        End If
    
    Next
    
'Stage3:
    
    'Create an AuditRecord for an New Records
    
    'Finding are there any new rows added
    iTotalNewRecords = iFinalRows - iInitialRows
    
    If iTotalNewRecords > 0 Then
        
        For iCtr = (iInitialRows + 1) To iFinalRows
            
            bRet = CreateUserRoleAudit(FinalRowArray(iCtr).UnitId, FinalRowArray(iCtr).UserId, FinalRowArray(iCtr).ChangeUser, _
                    "", "", "", FinalRowArray(iCtr).RoleActiveFrom, FinalRowArray(iCtr).RoleActiveTo, _
                    FinalRowArray(iCtr).Role, "C")
            
            If bRet = False Then
                g_clsErrorHandling.DisplayError "Error in Creating Audit Record"
                Exit Sub
            End If
        
        Next iCtr
        
    End If
    
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

Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    On Error GoTo Failed
    m_ReturnCode = MSGFailure
    
    Set m_clsUserRoleTable = New UserRoleTable
    
    ' Set the listview headers
    SetListViewHeaders
    
    If bIsEdit Then
        SetEditState
    End If

    SetUpInitialUserRoleAuditData
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Public Sub SetEditState()
    
    On Error GoTo Failed
' TW 13/04/2007 EP2_2375
    Set m_colKeys = New Collection
    m_colKeys.Add m_sUserID
' TW 13/04/2007 EP2_2375 End
    
    TableAccess(m_clsUserRoleTable).SetKeyMatchValues m_colKeys

    ' Get the data from the database
    TableAccess(m_clsUserRoleTable).GetTableData

    ' Validate we have the record
    If TableAccess(m_clsUserRoleTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit User Roles - Unable to locate User Role"
    End If
    
    PopulateUserRoles
    
    'TxtUserID.Text = frmEditUser.m_sUserID
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetTableAccess(clsTableAccess As TableAccess)
    On Error GoTo Failed
    Dim sFunctionName As String

    sFunctionName = "SetTableAccess"
    
    Set clsTableAccess = m_clsUserRoleTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub PopulateUserRoles()
    On Error GoTo Failed
        
    
    Dim sUserID As String
    
    If Len(frmEditUser.m_sUserID) = 0 Then
        g_clsErrorHandling.RaiseError errKeysEmpty, TypeName(Me) & " Populate User Roles"
    End If
    
    ' Get the UserRoles
    m_clsUserRoleTable.GetUserRolesForUserID frmEditUser.m_sUserID
        
    g_clsFormProcessing.PopulateFromRecordset frmEditUser.lvUserRoles, m_clsUserRoleTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetListViewHeaders()

    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Unit ID"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Active From"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Active To"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 80
    lvHeaders.sName = "Role"
    headers.Add lvHeaders
    
    frmEditUser.lvUserRoles.AddHeadings headers
    frmEditUser.lvUserRoles.LoadColumnDetails TypeName(Me)
    
    frmEditUser.lvUserRoles.HideSelection = False
    
    frmEditUser.lvUserRoles.AllowColumnReorder = False
    frmEditUser.lvUserRoles.Sorted = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub





Public Function SetScreenFields() As Boolean
    
    'This routine is stubbed in this class as it uses a bound datagrid.
    SetScreenFields = True

End Function

Public Sub SaveScreenData()
    
    Dim iRecordIdx As Integer
    
    'If we're adding, then the UserID won't be set against each of the grid row/records.
    If m_bIsEdit = False Then
        'If there are any records then move to the first.
        If TableAccess(m_clsUserRole).RecordCount > 0 Then
            TableAccess(m_clsUserRole).MoveFirst
        
            'Loop through each record in the recordset and set the userid.
            For iRecordIdx = 1 To TableAccess(m_clsUserRole).RecordCount
                'If the key is empty then set it now.
                If m_clsUserRole.GetUserID = "" Then
                    m_clsUserRole.SetUserID m_sUserID
                End If
                
                'Move onto the next record.
                TableAccess(m_clsUserRole).MoveNext
            Next iRecordIdx
        End If
    End If
    
End Sub

Public Sub DoUpdates()
    
    Dim clsTable As TableAccess

    On Error GoTo Failed
    
    If TableAccess(m_clsUserRoleTable).RecordCount > 0 Then
        
        'Update the UserRole table.
        Set clsTable = m_clsUserRoleTable
        clsTable.Update
    
        DoUpdateUserRoleAuditTableData
    
    Else
        g_clsErrorHandling.RaiseError errGeneralError, " User Role must be entered"
    End If

    Exit Sub
    
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetUserID(ByVal sUserID As String)
    m_sUserID = sUserID
End Sub

Public Sub SetUpInitialUserRoleAuditData()
        
    On Error GoTo Failed

    Dim nRowNum  As Integer
    Dim rs As ADODB.Recordset
    
    If TableAccess(m_clsUserRoleTable).RecordCount > 0 Then
        
        ReDim Preserve InitialRowArray(TableAccess(m_clsUserRoleTable).RecordCount)
        
        Set rs = TableAccess(m_clsUserRoleTable).GetRecordSet
     
        If rs.RecordCount > 0 Then
        
            rs.MoveFirst
            nRowNum = 0
            Do While rs.EOF = False
        
                UserRoleAuditRecord.UserId = m_sUserID
                UserRoleAuditRecord.AuditDate = g_clsSQLAssistSP.GetSystemDate()
                UserRoleAuditRecord.ChangeUser = g_sSupervisorUser
                UserRoleAuditRecord.UnitId = rs.fields("UNITID").Value
 '               MsgBox UserRoleAuditRecord.UnitId
                If Not IsNull(rs.fields("USERROLEACTIVEFROM")) Then
                    UserRoleAuditRecord.PrevRoleActiveFrom = rs.fields("USERROLEACTIVEFROM").Value
                Else
                    UserRoleAuditRecord.PrevRoleActiveFrom = ""
                End If
                
                If Not IsNull(rs.fields("USERROLEACTIVETO")) Then
                    UserRoleAuditRecord.PrevRoleActiveTo = rs.fields("USERROLEACTIVETO").Value
                Else
                    UserRoleAuditRecord.PrevRoleActiveTo = ""
                End If
                
                If Not IsNull(rs.fields("USERROLETEXT")) Then
                    'UserRoleAuditRecord.PrevRole = rs.fields("USERROLETEXT").Value
                    UserRoleAuditRecord.PrevRole = g_clsCombo.GetValueIDFromValueName("UserRole", rs.fields("USERROLETEXT").Value)
                Else
                    UserRoleAuditRecord.PrevRole = ""
                End If
                
                nRowNum = nRowNum + 1
                UserRoleAuditRecord.RowNum = nRowNum
                InitialRowArray(nRowNum) = UserRoleAuditRecord
                
                rs.MoveNext
            
            Loop
        
        Else

            UserRoleAuditRecord.UserId = m_sUserID
            UserRoleAuditRecord.AuditDate = g_clsSQLAssistSP.GetSystemDate()
            UserRoleAuditRecord.ChangeUser = g_sSupervisorUser
            UserRoleAuditRecord.UnitId = ""
            UserRoleAuditRecord.PrevRoleActiveFrom = ""
            UserRoleAuditRecord.PrevRoleActiveTo = ""
            UserRoleAuditRecord.PrevRole = ""
            UserRoleAuditRecord.Flag = "NEW"
            UserRoleAuditRecord.RowNum = "0"
            ReDim Preserve InitialRowArray(0)
            InitialRowArray(0) = UserRoleAuditRecord

        End If
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function ValidateScreenData() As Boolean
    
    
    Dim bRet As Boolean
    bRet = True
    
    If frmEditUser.lvUserRoles.ListItems.Count <= 0 Then
        frmEditUser.SetActiveTab UserRoles, True
        g_clsErrorHandling.RaiseError errGeneralError, " User Roles must be entered"
    End If
        
    ValidateScreenData = bRet

End Function
Public Sub SetUpFinalUserRoleAuditData()
        
    Dim nRowNum  As Integer
    Dim rs As ADODB.Recordset
    
    If TableAccess(m_clsUserRoleTable).RecordCount > 0 Then
        
        ReDim Preserve FinalRowArray(TableAccess(m_clsUserRoleTable).RecordCount)
        
        Set rs = TableAccess(m_clsUserRoleTable).GetRecordSet
     
        If rs.RecordCount > 0 Then
        
            rs.MoveFirst
            nRowNum = 0
            Do While rs.EOF = False
        
                UserRoleAuditRecord.UserId = m_sUserID
                UserRoleAuditRecord.AuditDate = g_clsSQLAssistSP.GetSystemDate()
                UserRoleAuditRecord.ChangeUser = g_sSupervisorUser
                UserRoleAuditRecord.UnitId = rs.fields("UNITID").Value
                If Not IsNull(rs.fields("USERROLEACTIVEFROM")) Then
                    UserRoleAuditRecord.RoleActiveFrom = rs.fields("USERROLEACTIVEFROM").Value
                Else
                    UserRoleAuditRecord.RoleActiveFrom = ""
                End If
                
                If Not IsNull(rs.fields("USERROLEACTIVETO")) Then
                    UserRoleAuditRecord.RoleActiveTo = rs.fields("USERROLEACTIVETO").Value
                Else
                    UserRoleAuditRecord.RoleActiveTo = ""
                End If
                
                If Not IsNull(rs.fields("USERROLETEXT")) Then
                    UserRoleAuditRecord.Role = g_clsCombo.GetValueIDFromValueName("UserRole", rs.fields("USERROLETEXT").Value)
                    'UserRoleAuditRecord.Role = rs.fields("ROLE").Value
                Else
                    UserRoleAuditRecord.Role = ""
                End If
                
                nRowNum = nRowNum + 1
                UserRoleAuditRecord.RowNum = nRowNum
                FinalRowArray(nRowNum) = UserRoleAuditRecord
                
                rs.MoveNext
            
            Loop
        
        Else

            UserRoleAuditRecord.UserId = m_sUserID
            UserRoleAuditRecord.AuditDate = g_clsSQLAssistSP.GetSystemDate()
            UserRoleAuditRecord.ChangeUser = g_sSupervisorUser
            UserRoleAuditRecord.UnitId = ""
            UserRoleAuditRecord.PrevRoleActiveFrom = ""
            UserRoleAuditRecord.PrevRoleActiveTo = ""
            UserRoleAuditRecord.PrevRole = ""
            UserRoleAuditRecord.Flag = "NEW"
            UserRoleAuditRecord.RowNum = "0"
            ReDim Preserve FinalRowArray(0)
            FinalRowArray(0) = UserRoleAuditRecord

        End If
        
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub



Public Sub SetAddState()
    'Stub.
End Sub












Private Sub SetGridFields()
    
    

End Sub


