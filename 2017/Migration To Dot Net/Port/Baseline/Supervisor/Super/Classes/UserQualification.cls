VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQualification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : UserQualification
' Description   : Handle all controls relevant to the Qualification tab.
' History       :
'
' Prog      Date        Description
' STB       28/11/2001  SYS2912 - Removed the frmMainUserDetails screen and
'                       ported its functionality to the frmEditUser screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Collections to store various items.
Private m_colComboIDS As Collection
Private m_colComboValues As Collection

'The underlying qualification table.
Private m_clsQualification As New QualificationTable

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The current record's user ID.
Private m_sUserID As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if the form is in an add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the add/edit state of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated accoding to the
'                 underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
        
    Dim rs As ADODB.Recordset
    Dim colValues As Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    If bIsEdit Then
        'Create a keys collection and add the user id into it.
        Set colValues = New Collection
        colValues.Add m_sUserID
        
        'Set the keys collection against the table object.
        Set clsTableAccess = m_clsQualification
        clsTableAccess.SetKeyMatchValues colValues
        
        'Retrieve any available data for it.
        Set rs = clsTableAccess.GetTableData()
        clsTableAccess.ValidateData
    Else
        'Initialise the table,object.
        Set rs = TableAccess(m_clsQualification).GetTableData(POPULATE_EMPTY)
    End If
    
    'Enable the data grid.
    frmEditUser.dgQualifications.Enabled = True
    
    'Bind the datagrid to the underlying recordset.
    Set frmEditUser.dgQualifications.DataSource = TableAccess(m_clsQualification).GetRecordSet
    
    g_clsCombo.FindComboGroup "UserQualification", m_colComboValues, m_colComboIDS
    
    'Initialise the datagrid columns.
    SetGridFields

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Initialises the datagrid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    
    Dim bRet As Boolean
    Dim sCombo As String
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsCombo As New ComboUtils
    
    bRet = True

    ' Qualification Text
    fields.sField = "QualificationTypeText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Qualification must be entered"
    fields.sTitle = "User Qualification"
    fields.sOtherField = "QualificationType"

    If m_colComboValues.Count > 0 And m_colComboIDS.Count > 0 Then
        Set fields.colComboValues = m_colComboValues
        Set fields.colComboIDS = m_colComboIDS
        colFields.Add fields
    Else
        MsgBox "Unable to locate combo group" + sCombo
    End If
    
    ' Qualification
    fields.sField = "QualificationType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' UserID
    fields.sField = "UserID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sUserID
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Active From
    fields.sField = "QualificationActiveFrom"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Active From Date must be entered"
    fields.sTitle = "Active From"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.bDateField = True
    colFields.Add fields

    ' Active To
    fields.sField = "QualificationActiveTo"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Active To"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.bDateField = True
    colFields.Add fields

    frmEditUser.dgQualifications.SetColumns colFields, "EditUserQualification", "User Qualifications"
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Loads all details from the tables into the screens controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetScreenFields() As Boolean
    
    'This routine is stubbed in this class as it uses a bound datagrid.
    SetScreenFields = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all details from the relevant screen controls to table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()

    Dim iRecordIdx As Integer
    
    'If we're adding, then the UserID won't be set against each of the grid row/records.
    If m_bIsEdit = False Then
        'If there are any records then move to the first.
        If TableAccess(m_clsQualification).RecordCount > 0 Then
            TableAccess(m_clsQualification).MoveFirst
                
            'Loop through each record in the recordset and set the userid.
            For iRecordIdx = 1 To TableAccess(m_clsQualification).RecordCount
                'If the key is empty then set it now.
                If m_clsQualification.GetUserID = "" Then
                    m_clsQualification.SetUserID m_sUserID
                End If
                
                'Move onto the next record.
                TableAccess(m_clsQualification).MoveNext
            Next iRecordIdx
        End If
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Update all underlying table objects handled by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
        
    Dim clsTable As TableAccess
    
    On Error GoTo Failed
    
    'Update the Qualification table.
    Set clsTable = m_clsQualification
    clsTable.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetUserID
' Description   : Sets the user ID at module level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetUserID(ByVal sUserID As String)
    m_sUserID = sUserID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates any control values relevant to this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = True
    bRet = frmEditUser.dgQualifications.ValidateRows()
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsQualification)
    
        If bRet = False Then
            MsgBox "User Qualification must be unique", vbCritical
        End If
    
    End If

    ValidateScreenData = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Configure the screen controls to work in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Configure the screen controls to work in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    'Stub.
End Sub
