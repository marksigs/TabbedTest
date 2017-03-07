VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IncomeFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : IncomeFactor
' Description   : Support class for Allowable Income Factors. This classes
'                 responds to combo click events and filters the contents of
'                 dependant combo controls dynamically.
' Change history
' Prog  Date        Description
' STB   13-May-2002 SYS4417 Added AllowableIncomeFactors.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' TK        01/12/2005  MAR575 cboEmploymentStatus_Click
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The parent form's add/edit mode.
Private m_bIsEdit As Boolean

'Indicates to the caller whether any update was successful of not.
Private m_uReturnCode As MSGReturnCode

'A collection containing the primary key values of the record we're editing.
Private m_colKeys As Collection

'Form controls whose events are to be handled within this class.
Private WithEvents cmdOK As CommandButton
Attribute cmdOK.VB_VarHelpID = -1
Private WithEvents cmdCancel As CommandButton
Attribute cmdCancel.VB_VarHelpID = -1
Private WithEvents cmdAnother As CommandButton
Attribute cmdAnother.VB_VarHelpID = -1
Private WithEvents cboIncomeGroup As MSGComboBox
Attribute cboIncomeGroup.VB_VarHelpID = -1
Private WithEvents cboEmploymentStatus As MSGComboBox
Attribute cboEmploymentStatus.VB_VarHelpID = -1

'The parent form which this class is servicing.
Private m_frmParent As frmEditIncomeFactors

'The underlying table object holding the record we're editing.
Private m_clsIncomeFactorTable As AllowableIncomeFactorTable

'Underlying table object used to populate the lender combo.
Private m_clsLenderTable As MortgageLendersTable


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function    : SetForm
' Description : Associates the specified form reference with this class. Event
'               handlers are then setup for the controls on the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog  Date        Description
' SA    04/11/2002  BMIDS00646 Employment Status now part of primary key
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetForm(ByRef frmParent As Form)
    
    'Store a module-level reference to the specified form.
    Set m_frmParent = frmParent
    
    'Setup event handlers for the required controls on the form.
    Set cmdOK = m_frmParent.cmdOK
    Set cmdCancel = m_frmParent.cmdCancel
    Set cmdAnother = m_frmParent.cmdAnother
    Set cboIncomeGroup = m_frmParent.cboIncomeGroup
    Set cboEmploymentStatus = m_frmParent.cboEmploymentStatus
    
End Sub

Public Sub Initialise()
        
    'Populate combos and other static controls.
    PopulateScreenControls
    
    'Load the specified record.
    TableAccess(m_clsIncomeFactorTable).SetKeyMatchValues m_colKeys
    m_clsIncomeFactorTable.GetIncomeFactors
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Update the screen from the underlying record data.
    SetScreenFields
    
End Sub

Public Sub Terminate()
    
    Set cmdOK = Nothing
    Set m_colKeys = Nothing
    Set cmdCancel = Nothing
    Set cmdAnother = Nothing
    Set m_frmParent = Nothing
    Set cboIncomeGroup = Nothing
    
    Set cboEmploymentStatus = Nothing
    Set m_clsLenderTable = Nothing
    Set m_clsIncomeFactorTable = Nothing
    
End Sub

Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub

Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function

Private Sub cboEmploymentStatus_Click()
        
    Dim sEmploymentStatus As String
    
    On Error GoTo Err_Handler
    
    'Get the validation type for the currently selected EmploymentStatus.
    sEmploymentStatus = GetValidationFromCombo("EmploymentStatus", m_frmParent.cboEmploymentStatus)
    
    'Populate the IncomeGroup combo based upon the value in the EmploymentStatus combo.
    Select Case sEmploymentStatus 'TK 01/12/2005 MAR575 - Added hardcoded employment status - This should be fixed by adding new function which returns all combo validation types.
        Case "E", "CON":      g_clsFormProcessing.PopulateComboByValidation m_frmParent.cboIncomeGroup, "IncomeGroupType", "E"
        Case "S", "C": g_clsFormProcessing.PopulateComboByValidation m_frmParent.cboIncomeGroup, "IncomeGroupType", "S"
        Case Else:     g_clsFormProcessing.PopulateComboByValidation m_frmParent.cboIncomeGroup, "IncomeGroupType", "O"
    End Select

    'Clear the child selections.
    m_frmParent.cboIncomeGroup.Text = COMBO_NONE
    
    'Only enable the child combo if we have a value.
    m_frmParent.cboIncomeGroup.Enabled = (sEmploymentStatus <> "")
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub

Private Sub cboIncomeGroup_Click()
            
    Dim sIncomeGroup As String
    
    On Error GoTo Err_Handler
    
    'Get the ValueID for the currently selected IncomeGroup.
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboIncomeGroup, sIncomeGroup, GET_CONTROL_VALUE
    
    'Populate the IncomeGroup combo based upon the value in the EmploymentStatus combo.
    Select Case sIncomeGroup
        Case "1":  g_clsFormProcessing.PopulateCombo "ConfirmedIncomeType", m_frmParent.cboIncomeType, True
        Case "2":  g_clsFormProcessing.PopulateCombo "IncomeType", m_frmParent.cboIncomeType, True
        Case "3":  g_clsFormProcessing.PopulateCombo "SelfEmpConfirmedIncomeType", m_frmParent.cboIncomeType, True
        Case "4":  g_clsFormProcessing.PopulateCombo "SelfEmpDeclaredIncomeType", m_frmParent.cboIncomeType, True
        Case Else: g_clsFormProcessing.PopulateCombo "UnEarnedIncomeDescription", m_frmParent.cboIncomeType, True
    End Select

    'Clear the child selections.
    m_frmParent.cboIncomeType.Text = COMBO_NONE
    
    'Only enable the child combo if we have a value.
    m_frmParent.cboIncomeType.Enabled = (sIncomeGroup <> "")
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub

Private Sub Class_Initialize()
    
    Set m_clsLenderTable = New MortgageLendersTable
    Set m_clsIncomeFactorTable = New AllowableIncomeFactorTable
    
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub

Private Sub cmdAnother_Click()
    
    On Error GoTo Err_Handler
    
    'Perform ok processing. Will raise an error if it fails.
    DoOKProcessing
    
    'If the record was saved, commit the transaction and begin another.
    g_clsDataAccess.CommitTrans
    g_clsDataAccess.BeginTrans
    
    'Clear all fields, set the default focus.
    g_clsFormProcessing.ClearScreenFields m_frmParent
    
    'Re-initialise the screen.
    Initialise
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub

Private Sub cmdCancel_Click()
    
    On Error GoTo Err_Handler
    
    m_frmParent.Hide
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo Err_Handler
    
    'Commit the changes.
    DoOKProcessing
    
    'Set the return status to success.
    m_uReturnCode = MSGSuccess
    
    'Hide the form and return control to the caller.
    m_frmParent.Hide
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.DisplayError Err.DESCRIPTION
    End If
End Sub

Private Sub DoOKProcessing()
    
    Dim bValid As Boolean
    
    On Error GoTo Err_Handler
    
    'Ensure all mandatory fields have been populated.
    bValid = g_clsFormProcessing.DoMandatoryProcessing(m_frmParent)

    'If they have then proceed.
    If bValid = True Then
        'Validate all captured input.
        bValid = ValidateScreenData()
        
        'If the data was valid.
        If bValid = True Then
            'Save the data.
            SaveScreenData
            
            'Ensure the record is flagged for promotion.
            SaveChangeRequest
        End If
    End If
    
Err_Handler:
    If Err Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If
End Sub

Private Sub SetAddState()
        
    'Create a blank income factor record.
    g_clsFormProcessing.CreateNewRecord m_clsIncomeFactorTable
    
End Sub

Private Sub SetEditState()
    m_frmParent.cmdAnother.Enabled = False
End Sub

Private Sub PopulateScreenControls()

    Dim colComboFields As Collection
    
    'We need to load the lenders table with lenders before binding it to the
    'lenders combo.
    TableAccess(m_clsLenderTable).GetTableData POPULATE_ALL
    
    'Create a collection of field names to use when binding the lender combo.
    Set colComboFields = New Collection
    colComboFields.Add "LENDERNAME"
    colComboFields.Add "ORGANISATIONID"
    
    'Bind the lender combo to a table object.
    g_clsFormProcessing.PopulateComboFromTable m_frmParent.cboLenderName, m_clsLenderTable, colComboFields
    
    'EmploymentStatus combo.
    g_clsFormProcessing.PopulateCombo "EmploymentStatus", m_frmParent.cboEmploymentStatus, True

End Sub

Private Sub SetScreenFields()
    
    Dim vVal As Variant
    
    'Select the lender.
    vVal = m_clsIncomeFactorTable.GetOrganisationID
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboLenderName, vVal, SET_CONTROL_VALUE
    
    'Select EmploymentStatus.
    vVal = m_clsIncomeFactorTable.GetEmploymentStatus
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboEmploymentStatus, vVal, SET_CONTROL_VALUE
    
    'Select IncomeGroup.
    vVal = m_clsIncomeFactorTable.GetIncomeGroup
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboIncomeGroup, vVal, SET_CONTROL_VALUE
    
    'Select IncomeType.
    vVal = m_clsIncomeFactorTable.GetIncomeType
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboIncomeType, vVal, SET_CONTROL_VALUE
    
    'Factor.
    m_frmParent.txtFactor.Text = m_clsIncomeFactorTable.GetFactor
    
End Sub

Private Sub SaveScreenData()
    
    Dim vVal As Variant
    
    'Organisation ID.
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboLenderName, vVal, GET_CONTROL_VALUE
    m_clsIncomeFactorTable.SetOrganisationID g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vVal))
    
    'EmploymentStatus.
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboEmploymentStatus, vVal, GET_CONTROL_VALUE
    m_clsIncomeFactorTable.SetEmploymentStatus vVal
    
    'IncomeGroupType.
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboIncomeGroup, vVal, GET_CONTROL_VALUE
    m_clsIncomeFactorTable.SetIncomeGroup vVal
    
    'IncomeType.
    g_clsFormProcessing.HandleComboExtra m_frmParent.cboIncomeType, vVal, GET_CONTROL_VALUE
    m_clsIncomeFactorTable.SetIncomeType vVal
    
    'Factor.
    m_clsIncomeFactorTable.SetFactor m_frmParent.txtFactor.Text
    
    'Update the underlying table.
    TableAccess(m_clsIncomeFactorTable).Update
    
End Sub

Private Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function


Private Sub SaveChangeRequest()
            
    Dim sDescription As String
            
    'If we're adding, update the table object's keys collection.
    If m_bIsEdit = False Then
        Set m_colKeys = New Collection
        m_colKeys.Add m_clsIncomeFactorTable.GetOrganisationID
        m_colKeys.Add m_clsIncomeFactorTable.GetIncomeGroup
        m_colKeys.Add m_clsIncomeFactorTable.GetIncomeType
        '++ BMIDS00646 SA Employment Status needs setting
        m_colKeys.Add m_clsIncomeFactorTable.GetEmploymentStatus
        TableAccess(m_clsIncomeFactorTable).SetKeyMatchValues m_colKeys
    End If
    
    'Generate a description.
    sDescription = m_clsIncomeFactorTable.GetDescription
    
    'Delegate the promotion tag to a global.
    g_clsHandleUpdates.SaveChangeRequest m_clsIncomeFactorTable, sDescription
    
End Sub


Private Function GetValidationFromCombo(ByVal sComboGroup As String, ByRef cboCombo As MSGComboBox) As String

    Dim vVal As Variant
    Dim clsComboTable As ComboValidationTable

    g_clsFormProcessing.HandleComboExtra cboCombo, vVal, GET_CONTROL_VALUE
    
    If vVal <> "" Then
        Set clsComboTable = New ComboValidationTable
        clsComboTable.GetValidationTypeForValueID sComboGroup, CStr(vVal)
        GetValidationFromCombo = clsComboTable.GetValidationType()
    End If
    
End Function

