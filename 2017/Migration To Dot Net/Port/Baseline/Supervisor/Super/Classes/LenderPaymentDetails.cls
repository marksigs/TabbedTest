VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderPaymentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderPaymentDetails
' Description   : Support class to handle the CHAPS, HEXAGON and Barclays
'                 parameter values.
'
' Change history
' Prog  Date        Description
' STB   07-May-2002 MSMS0069 Created class.
' CL    28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Indicates the add/edit mode of the form.
Private m_bIsEdit As Boolean

'Underlying table object.
Private m_clsLenderTable As MortgageLendersTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Updates all tables used by this tab. This method will be
'                 called externally to this class to commit all data that has
'                 been added to any of the table classes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    'Ignore this method. Update on the underlying table object will be done by
    'another tab-handler.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to retrieve all data from the  database, setup fields
'                 etc.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    m_bIsEdit = bIsEdit
    
    'Only do this if we are editing a record.
    If bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the edit state of this tab - edit or add.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTableClass
' Description   : Accosiate the Lender's table with this tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTableClass(ByRef clsTableAccess As TableAccess)
    Set m_clsLenderTable = clsTableAccess
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all screen fields to the database. This method will
'                 be called externally to this class to save LenderDetails tab
'                 of the Lender Details form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
        
    'Save the payment fields to the underlying Lender table.
    m_clsLenderTable.SetBarclaysCHAPSSortCode frmLenderDetails.txtBarclaysCHAPSSortCode.Text
    m_clsLenderTable.SetBarclaysCHAPSAccountNumber frmLenderDetails.txtBarclaysCHAPSAccountNumber.Text
    m_clsLenderTable.SetBarclaysCHAPSOrderingInst frmLenderDetails.txtBarclaysCHAPSOrderingInstitution.Text
    
    m_clsLenderTable.SetHexagonCHAPSAccountNumber frmLenderDetails.txtHexagonCHAPSAccountNumber.Text
    m_clsLenderTable.SetBACSSortCode frmLenderDetails.txtBACSSortCode.Text
    m_clsLenderTable.SetBACSAccountNumber frmLenderDetails.txtBACSAccountNumber.Text
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates that all data entered on the Lender Details tab is
'                 valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
        
    'There is no specific validation required.
    ValidateScreenData = True
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all screen fields from the data retrieved from the
'                 database. This method will be called externally to this class
'                 to populate the LenderDetails tab of the Lender Details form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()

    'Populate the payment screen controls from the underlying table.
    frmLenderDetails.txtBarclaysCHAPSSortCode.Text = m_clsLenderTable.GetBarclaysCHAPSSortCode
    frmLenderDetails.txtBarclaysCHAPSAccountNumber.Text = m_clsLenderTable.GetBarclaysCHAPSAccountNumber
    frmLenderDetails.txtBarclaysCHAPSOrderingInstitution.Text = m_clsLenderTable.GetBarclaysCHAPSOrderingInst
    
    frmLenderDetails.txtHexagonCHAPSAccountNumber.Text = m_clsLenderTable.GetHexagonCHAPSAccountNumber
    frmLenderDetails.txtBACSSortCode.Text = m_clsLenderTable.GetBACSSortCode
    frmLenderDetails.txtBACSAccountNumber.Text = m_clsLenderTable.GetBACSAccountNumber
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup the controls for edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup the controls for add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Old method, now replaced with Initialise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    'Stub.
End Sub
