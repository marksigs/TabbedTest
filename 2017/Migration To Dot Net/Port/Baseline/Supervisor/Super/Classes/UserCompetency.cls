VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserCompetency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : UserCompetency
' Description   : Handle all controls relevant to the User Competency tab.
' History       :
'
' Prog      Date        Description
' STB       28/11/2001  SYS2912 - Removed the frmMainUserDetails screen and
'                       ported its functionality to the frmEditUser screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History:
'Prog   Date        AQR     Decription
'MV     18/12/2002  BM0085  Added SetUserCompetencyAuditTableData(),CreateUserCompetecyAuditRecord()
'                           Amended DoUpdates();IntialiseForms()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Mars Specific History:
' Prog   Date        AQR        Decription
' GHun   16/11/2005  MAR312     Added UpdateCompentencyCombo
' RF     18/01/2006  MAR1000    Added extra error handling
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Collections to store various items.
Private m_colComboIDS As New Collection
Private m_colComboValues As New Collection

'The underlying user role table object.
Private m_clsUserCompetency As New UserCompetencyTable  'MAR312 GHun

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The current record's user ID.
Private m_sUserID As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if the form is in an add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const USER_ID                 As Long = 0

Private Type TypeUserCompetencyAuditRecord
    UserId As String
    AuditDate As String
    ChangeUser As String
    CompetencyType As String
    CompetenctyActiveFrom As String
    PrevCompetencyType As String
    PrevCompetencyActiveFrom As String

End Type

Private UserCompetencyAuditRecord As TypeUserCompetencyAuditRecord
Private m_clsCompetency As CompetencyTable

Public Sub SetUserCompetencyAuditTableData()
    
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sField As FieldData
    
    If frmEditUser.dgCompetencies.Rows > 0 Then
        For iRow = 0 To frmEditUser.dgCompetencies.Rows - 1
            For iCol = 0 To frmEditUser.dgCompetencies.Columns.Count - 1
                If frmEditUser.dgCompetencies.Columns(iCol).DataField = "COMPETENCYTYPE" Then
                    UserCompetencyAuditRecord.PrevCompetencyType = frmEditUser.dgCompetencies.GetAtRowCol(iRow, iCol)
                End If
                If frmEditUser.dgCompetencies.Columns(iCol).DataField = "USERCOMPETENCYACTIVEFROM" Then
                    UserCompetencyAuditRecord.PrevCompetencyActiveFrom = frmEditUser.dgCompetencies.GetAtRowCol(iRow, iCol)
                End If
            Next iCol
        Next iRow
    Else
        UserCompetencyAuditRecord.PrevCompetencyType = ""
        UserCompetencyAuditRecord.PrevCompetencyActiveFrom = ""
    End If
    
End Sub

Public Sub CreateUserCompetecyAuditRecord()
    
    On Error GoTo Failed
    
    Dim sSQL  As String
    Dim iRow As Integer
    Dim iCol As Integer
    
    UserCompetencyAuditRecord.UserId = frmEditUser.txtUserDetails(USER_ID).Text
    UserCompetencyAuditRecord.ChangeUser = g_sSupervisorUser
    
    For iRow = 0 To frmEditUser.dgCompetencies.Rows - 1
        For iCol = 0 To frmEditUser.dgCompetencies.Columns.Count - 1
            If frmEditUser.dgCompetencies.Columns(iCol).DataField = "COMPETENCYTYPE" Then
                UserCompetencyAuditRecord.CompetencyType = frmEditUser.dgCompetencies.GetAtRowCol(iRow, iCol)
            End If
            If frmEditUser.dgCompetencies.Columns(iCol).DataField = "USERCOMPETENCYACTIVEFROM" Then
                UserCompetencyAuditRecord.CompetenctyActiveFrom = CStr(frmEditUser.dgCompetencies.GetAtRowCol(iRow, iCol))
            End If
        Next iCol
    Next iRow
    
    If (UserCompetencyAuditRecord.PrevCompetencyType <> UserCompetencyAuditRecord.CompetencyType) Or _
        (UserCompetencyAuditRecord.PrevCompetencyActiveFrom <> UserCompetencyAuditRecord.CompetenctyActiveFrom) Then
        
        sSQL = ""
        sSQL = sSQL + " INSERT INTO USERCOMPETENCYAUDIT ( USERID,AUDITDATE,CHANGEUSER,PREVCOMPETENCYTYPE,PREVCOMPETENCYACTIVEFROM,COMPETENCYTYPE,COMPETENCYACTIVEFROM) "
        sSQL = sSQL + " VALUES ( "
        sSQL = sSQL + " '" & UserCompetencyAuditRecord.UserId & "' , "
        sSQL = sSQL + g_clsSQLAssistSP.GetSystemDate & " , "
        sSQL = sSQL + " '" & UserCompetencyAuditRecord.ChangeUser & "' , "
        If UserCompetencyAuditRecord.PrevCompetencyType = "" Then
            sSQL = sSQL + "Null" & " , "
        Else
            sSQL = sSQL + UserCompetencyAuditRecord.PrevCompetencyType & " , "
        End If
        
        If UserCompetencyAuditRecord.PrevCompetencyActiveFrom = "" Then
            sSQL = sSQL + "Null" & " , "
        Else
            sSQL = sSQL + g_clsSQLAssistSP.FormatDate(UserCompetencyAuditRecord.PrevCompetencyActiveFrom) & ", "
        End If
        
        If UserCompetencyAuditRecord.CompetencyType = "" Then
            sSQL = sSQL + "Null" & " , "
        Else
            sSQL = sSQL + UserCompetencyAuditRecord.CompetencyType & " , "
        End If
        
        If UserCompetencyAuditRecord.CompetenctyActiveFrom = "" Then
            sSQL = sSQL + "Null" & " , "
        Else
            sSQL = sSQL + g_clsSQLAssistSP.FormatDate(UserCompetencyAuditRecord.CompetenctyActiveFrom) & ") "
        End If
        
        g_clsDataAccess.ExecuteCommand sSQL
        
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    Dim rs As ADODB.Recordset
    Dim colValues As Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    Set m_clsCompetency = New CompetencyTable
    
    If bIsEdit Then
        'Create a keys collection and add the user id into it.
        Set colValues = New Collection
        colValues.Add m_sUserID
        
        'Set the keys collection against the table object.
        Set clsTableAccess = m_clsUserCompetency
        clsTableAccess.SetKeyMatchValues colValues
        
        'Retrieve any available data for it.
        Set rs = clsTableAccess.GetTableData()
        clsTableAccess.ValidateData
        
    Else
        'Initialise the table,object.
        Set rs = TableAccess(m_clsUserCompetency).GetTableData(POPULATE_EMPTY)
    End If
    
    'Enable and bind the data grid to the recordset.
    frmEditUser.dgCompetencies.Enabled = True
    Set frmEditUser.dgCompetencies.DataSource = rs
    
    m_clsCompetency.GetCompetenciesAsCollection m_colComboValues, m_colComboIDS

    'Initialise the datagrid columns.
    SetGridFields
        
    'MV - 16/12/2002 - BM0085
    SetUserCompetencyAuditTableData
    
    If frmEditUser.dgCompetencies.Rows > 0 Then
        frmEditUser.dgCompetencies.AllowAdd = False
        frmEditUser.dgCompetencies.AllowDelete = False
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetGridFields()
    
    Dim bRet As Boolean
    Dim sCombo As String
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsCombo As New ComboUtils
    
    bRet = True

    ' User ID
    fields.sField = "UserID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sUserID
    fields.sError = ""
    fields.sTitle = ""
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.bDateField = False
    colFields.Add fields

    ' Competency Type
    fields.sField = "CompetencyType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Active From
    fields.sField = "UserCompetencyActiveFrom"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Active From Date must be entered"
    fields.sTitle = "Active From"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.bDateField = True
    colFields.Add fields

    ' Populate from the competencies setup from the competencies screen, then map the combo id
    ' to the combo value.
    ' Competency Type
    fields.sField = "CompetencyTypeText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Competency Type has to be entered"
    fields.sTitle = "Competency Type"
    Set fields.colComboValues = m_colComboValues
    Set fields.colComboIDS = m_colComboIDS
    fields.sOtherField = "CompetencyType"
    fields.bDateField = False
    colFields.Add fields

    frmEditUser.dgCompetencies.SetColumns colFields, "EditUserCompetency", "User Competencies"
End Sub

Public Function SetScreenFields() As Boolean
    
    'This routine is stubbed in this class as it uses a bound datagrid.
    SetScreenFields = True

End Function

Public Sub SaveScreenData()
On Error GoTo Failed
    
    Dim iRecordIdx As Integer
    
    'If we're adding, then the UserID won't be set against each of the grid row/records.
    If m_bIsEdit = False Then
        'If there are any records then move to the first.
        If TableAccess(m_clsUserCompetency).RecordCount > 0 Then
            TableAccess(m_clsUserCompetency).MoveFirst
        
            'Loop through each record in the recordset and set the userid.
            For iRecordIdx = 1 To TableAccess(m_clsUserCompetency).RecordCount
                'If the key is empty then set it now.
                If m_clsUserCompetency.GetUserID = "" Then
                    m_clsUserCompetency.SetUserID m_sUserID
                End If
                
                'Move onto the next record.
                TableAccess(m_clsUserCompetency).MoveNext
            Next iRecordIdx
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError _
        Err.Number, "Error in UserCompetency SaveScreenData: " & Err.DESCRIPTION
End Sub

Public Sub DoUpdates()
    
    Dim clsTable As TableAccess

    On Error GoTo Failed
    
    'Update the UserCompetency table.
    Set clsTable = m_clsUserCompetency
    clsTable.Update
    
    'MV - 16/12/2002 - BM0085
    CreateUserCompetecyAuditRecord
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError _
        Err.Number, "Error in UserCompetency DoUpdates: " & Err.DESCRIPTION
End Sub

Public Sub SetUserID(ByVal sUserID As String)
    m_sUserID = sUserID
End Sub

Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    
    bRet = True
    bRet = frmEditUser.dgCompetencies.ValidateRows()
    
    If bRet Then
        If frmEditUser.dgCompetencies.Rows = 0 Then
            frmEditUser.SetActiveTab CompetencyHistory, True
            g_clsErrorHandling.RaiseError errGeneralError, " User Competencies must be entered"
        End If
    
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsUserCompetency)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Competency and Start Date must be unique"
        End If
    
    End If
    
    ValidateScreenData = bRet
End Function

Public Sub SetAddState()
    'Stub.
End Sub

Public Sub SetEditState()
    'Stub.
End Sub

'MAR312 GHun
Public Sub UpdateCompentencyCombo(ByVal blnHasWorkgroupLimit As Boolean)
    Set m_colComboValues = Nothing
    Set m_colComboIDS = Nothing
    m_clsCompetency.GetCompetenciesAsCollection m_colComboValues, m_colComboIDS, blnHasWorkgroupLimit
    SetGridFields
End Sub
'MAR312 End
