VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntermediaryBankDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : IntermediaryBankDetails
' Description   : Tab-Handler for the Bank Details Intermediary tab.
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' STB       10/12/01    SYS2550: SQL-Server integration.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit  As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'Local reference to the main Edit form.
Private m_frmInt As frmEditIntermediaries

'Underlying table object for the record.
Private m_clsBankDetails As IntBankAccountTable

'Control Indexes.
Private Const ACCOUNT_NUMBER    As Long = 0
Private Const ACCOUNT_SORT_CODE As Long = 1
Private Const ACCOUNT_NAME      As Long = 2
Private Const ACCOUNT_BANK_NAME As Long = 3


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated
'                 accoding to the underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    On Error GoTo Failed

    'Store the add/edit state at module level.
    m_bIsEdit = bIsEdit

    'Set state-specific configuration.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Populate static data and combo lists.
    PopulateScreenControls
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup data for a new record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    
    On Error GoTo Failed

    'Create a blank details record.
    g_clsFormProcessing.CreateNewRecord m_clsBankDetails
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup data for an edit (i.e. load a record).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    On Error GoTo Failed
    
    'Set the keys collection against the table.
    TableAccess(m_clsBankDetails).SetKeyMatchValues m_colKeys
    
    'Load data into the table.
    TableAccess(m_clsBankDetails).GetTableData
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetForm
' Description   : Associate the Intermediaries form with the tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetForm(ByRef pfrmEditIntermediaries As frmEditIntermediaries)
    Set m_frmInt = pfrmEditIntermediaries
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Associate the collection of primary key values with the tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populate the screen from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    
    On Error GoTo Failed
    
    'Account Number
    m_frmInt.txtBankDetails(ACCOUNT_NUMBER).Text = m_clsBankDetails.GetAccountNumber
    
    'Sort Code
    m_frmInt.txtBankDetails(ACCOUNT_SORT_CODE).Text = m_clsBankDetails.GetSortCode
    
    'Account Name
    m_frmInt.txtBankDetails(ACCOUNT_NAME).Text = m_clsBankDetails.GetAccountName
    
    'Bank Name
    m_frmInt.txtBankDetails(ACCOUNT_BANK_NAME).Text = m_clsBankDetails.GetBankName
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Called from the interemediary tabbed form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    On Error GoTo Failed
    
    'Is this update and add?
    m_clsBankDetails.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
    
    'Account Number
    m_clsBankDetails.SetAccountNumber m_frmInt.txtBankDetails(ACCOUNT_NUMBER).Text
    
    'Sort Code
    m_clsBankDetails.SetSortCode m_frmInt.txtBankDetails(ACCOUNT_SORT_CODE).Text
    
    'Account Name
    m_clsBankDetails.SetAccountName m_frmInt.txtBankDetails(ACCOUNT_NAME).Text
    
    'Bank Name
    m_clsBankDetails.SetBankName m_frmInt.txtBankDetails(ACCOUNT_BANK_NAME).Text
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : saves the changes made to the bankdetails tab
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    
    On Error GoTo Failed

    TableAccess(m_clsBankDetails).Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Terminate
' Description   : Releases all module-level object references. This avoids a cyclic reference
'                 problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Terminate()
    
    Set m_frmInt = Nothing
    Set m_colKeys = Nothing
    Set m_clsBankDetails = Nothing

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates the data on the screen. Duplicates etc
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsBankDetails = New IntBankAccountTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Terminate
' Description   : Release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Terminate
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate static data and combo lists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    'Stub.
End Sub
