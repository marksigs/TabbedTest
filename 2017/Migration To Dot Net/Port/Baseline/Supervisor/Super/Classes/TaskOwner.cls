VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : TaskOwner
' Description   : Contains all data/functions associated with the ActionOwner table
' Change history
' Prog      Date        Description
'  AA       03/01/00    Added Class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_rs As ADODB.Recordset
Private m_sUnitID As String
Private Const m_sTableName = "ActionOwner"
Private Const m_UNIT_ID = "UNITID"
Private Const USER_ID = "USERID"
Private m_frm As Form

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Initialises ActionOwnertable, and sets the dgtaskOwner on form(m_frm)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    
    Dim clsOmigaUser As OmigaUserTable
    Dim colUpdateFields As Collection
    Dim colUpdateValues As Collection
    
    Set clsOmigaUser = New OmigaUserTable
    Set colUpdateValues = New Collection
    Set colUpdateFields = New Collection
    
    colUpdateFields.Add m_UNIT_ID
    colUpdateValues.Add m_sUnitID
    
    m_clsTableAccess.SetOrderCriteria USER_ID
    m_clsTableAccess.SetKeyMatchValues colUpdateValues
    m_clsTableAccess.SetKeyMatchFields colUpdateFields
    
    clsOmigaUser.GetUsersFromUnit m_sUnitID
    
    If bIsEdit = False Then
        m_clsTableAccess.GetTableData POPULATE_EMPTY
    Else
        m_clsTableAccess.GetTableData
    End If
    
    Set m_rs = m_clsTableAccess.GetRecordSet
    Set m_frm.dgTaskOwner.DataSource = m_rs
    
    m_clsTableAccess.SetCollectionField (m_UNIT_ID)
    
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetForm(frm As Form)
    Set m_frm = frm
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Sets the Datagrid ColumnHeadings
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colMatchFields As Collection
    Dim colComboValues As New Collection
    Dim clsCombo As New ComboUtils
    Dim colComboIDS As New Collection
    Dim sCombo As String
    Dim colValues As Collection
    Dim clsFormProccessing As FormProcessing
    Dim rs As ADODB.Recordset
    Dim clsUserRole As UserRoleTable
    
    On Error GoTo Failed
    Set clsFormProccessing = New FormProcessing
    Set colValues = New Collection
    Set clsUserRole = New UserRoleTable
    Set colMatchFields = New Collection
    
    ' UnitID
    fields.sField = "UNITID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sUnitID
    fields.sError = ""
    fields.sTitle = "UNITID"
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.bDateField = False
    colFields.Add fields
    
    ' Task Owner Type
    fields.sField = "TaskOwnerType"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "TASKOWNERTYPE"
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.bDateField = False
    colFields.Add fields
    
'     User ID
    fields.sField = "UserID"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "User ID must be entered"
    fields.sTitle = "User Id"
    
    Set rs = New ADODB.Recordset
    colMatchFields.Add ("UnitID")
    colValues.Add (m_sUnitID)

    TableAccess(clsUserRole).SetKeyMatchFields colMatchFields
    TableAccess(clsUserRole).SetKeyMatchValues colValues
    TableAccess(clsUserRole).SetOrderCriteria USER_ID
    TableAccess(clsUserRole).SetCollectionField USER_ID
    
    Set rs = TableAccess(clsUserRole).GetTableData
    Set fields.colComboValues = clsFormProccessing.GetCollectionFromRecordset(clsUserRole)
    Set fields.colComboIDS = fields.colComboValues
    
    fields.sOtherField = ""
    fields.bDateField = False
    colFields.Add fields

    ' Task Owner Type Text
    fields.sField = "TaskOwnerTypeText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Task Owner Type must be entered"
    fields.sTitle = "Task Owner Type"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing

    g_clsCombo.FindComboGroup "UserRole", colComboValues, colComboIDS, , , "ID"

    If colComboValues.Count > 0 And colComboIDS.Count > 0 Then
        Set fields.colComboValues = colComboValues
        Set fields.colComboIDS = colComboIDS
        colFields.Add fields
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate combo group" + sCombo
    End If

    fields.sOtherField = "TaskOwnerType"
    fields.bDateField = False
    colFields.Add fields

    m_frm.dgTaskOwner.SetColumns colFields, "TaskOwner", "Task Owner"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetUnitID(sUnitID As String)
    m_sUnitID = sUnitID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveTaskOwner
' Description   : Saves the data entered in the taskowner grid, to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SaveTaskOwner()
    
    On Error GoTo Failed

    Dim clsBandedTable As BandedTable
    Dim colUpdateValues As Collection
    
    Dim clsActionOwner As ActionOwnerTable
    Dim rs As ADODB.Recordset
    Dim colUpdateFields As Collection

    Set colUpdateFields = New Collection
    
    Set clsActionOwner = m_clsTableAccess
    Set colUpdateValues = New Collection
    Set clsBandedTable = clsActionOwner
    
    colUpdateValues.Add m_sUnitID
    colUpdateFields.Add m_UNIT_ID

    'Set the recordset to update
    Set rs = m_frm.dgTaskOwner.DataSource
    
    m_clsTableAccess.SetRecordSet rs
    clsBandedTable.SetTableAccess m_clsTableAccess
    clsBandedTable.SetUpdateValues colUpdateValues
    clsBandedTable.SetUpdateSets
    clsBandedTable.DoUpdateSets
        
    m_clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateTaskOwner
' Description   : Validates the data entered for TaskOwners
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function ValidateTaskOwner() As Boolean
    
    Dim bRet As Boolean
      
    On Error GoTo Failed
    bRet = m_frm.dgTaskOwner.ValidateRows
    
    If bRet = True Then
        Set m_rs = m_frm.dgTaskOwner.DataSource
        If Not m_rs Is Nothing Then
            m_clsTableAccess.SetRecordSet m_rs
            bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
        End If
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Task owner type and User Id must be unique"
        End If
    End If
    
    ValidateTaskOwner = bRet
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub Class_Initialize()
    On Error GoTo Failed
    Set m_clsTableAccess = New ActionOwnerTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetRecordSet(rs As ADODB.Recordset)
    Set m_rs = rs
    m_clsTableAccess.SetRecordSet rs
End Sub
