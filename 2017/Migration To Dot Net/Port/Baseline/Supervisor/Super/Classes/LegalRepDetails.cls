VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LegalRepDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LegalRepDetails
' Description   : Manages all form processing for the Legal Rep details tab on the Lender Details
'                 form.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port, use Supervisor version of SQLAssist instead
' STB       29/11/2001  SYS2912 - Removed the frmMainThirdParty screen and
'                       ported its functionality to frmEditThirdParty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ MARS Specific History
'~
'~ PROG     DATE        AQR     DESCRIPTION
'~ MF       02/08/2005  MAR20   Added Legal Rep Status to Legal Rep Tab
'~ TK       30/11/2005  MAR81   Added GetLastActivatedDate and GetUserID
'  TK       01/12/2005  MAR775  Amended SaveScreenData()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Text Edit constants.
Private Const PANEL_ID = 0
Private Const ASSOCIATE_PANEL_ID = 1
Private Const NUMBER_OF_PARTNERS = 2
Private Const INDEMNITY_EXPIRY_DATE = 3

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'Local table classes.
Private m_clsPanel As PanelTable
Private m_clsLegalRep As PanelLegalRepTable

'Current records key in the name and address directory table.
Private m_vDirectoryGUID As Variant
' TK 30/11/2005  MAR81
Private m_sLastActivatedDate As String
Private m_sUserID As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if this tab-handler is in add or edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetDirectoryGUID
' Description   : Sets the primary key for the parent record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDirectoryGUID(ByVal vGUID As Variant)
    m_vDirectoryGUID = vGUID
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the add/edit mode of this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Create underlying object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Class_Initialize()
    
    Set m_clsLegalRep = New PanelLegalRepTable
    Set m_clsPanel = New PanelTable
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated accoding to the
'                 underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    Dim sError As String
    Dim rs As ADODB.Recordset
    Dim nRecordCount As Integer
    Dim vDirectoryGUID As Variant
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    sError = "LegalRep:Initialise - "
    
    'Store the add/edit state of the form.
    m_bIsEdit = bIsEdit
    
    If bIsEdit = True Then
        'Load the legal rep record.
        Set clsTableAccess = m_clsLegalRep
        
        ' DJP SQL Server port
        colValues.Add m_vDirectoryGUID
        clsTableAccess.SetKeyMatchValues colValues
        
        'Verify a single record was found.
        Set rs = clsTableAccess.GetTableData()
        clsTableAccess.ValidateData
        nRecordCount = clsTableAccess.RecordCount
        
        If nRecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Incorrect number of LegalRep records returned: " & nRecordCount
        End If
    
        'Load the Panel record.
        Set clsTableAccess = m_clsPanel
        clsTableAccess.SetKeyMatchValues colValues
        
        'Verify a single record was found.
        Set rs = clsTableAccess.GetTableData()
        clsTableAccess.ValidateData
        nRecordCount = clsTableAccess.RecordCount
        
        If nRecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Incorrect number of LegalRep records returned: " & nRecordCount
        End If
    Else
        'Create empty records in the required tables.
        g_clsFormProcessing.CreateNewRecord m_clsPanel
        g_clsFormProcessing.CreateNewRecord m_clsLegalRep
    End If

    'Populate the screen combo's.
    g_clsFormProcessing.PopulateCombo "PaymentMethod", frmEditThirdParty.cboLegalPaymentMethod
    g_clsFormProcessing.PopulateCombo "TypeOfLegalRep", frmEditThirdParty.cboLegalRepType
    
    g_clsFormProcessing.PopulateCombo "CountryOfPractice", frmEditThirdParty.cboCountryOfPractice
    g_clsFormProcessing.PopulateCombo "CountryOfOrigin", frmEditThirdParty.cboCountryOfOrigin
    
    g_clsFormProcessing.PopulateCombo "LegalRepStatus", frmEditThirdParty.cboLegalRepStatus
    
    'AM LegalRepStatus set to inactive and disable it.
    If bIsEdit = False Then
        frmEditThirdParty.cboLegalRepStatus.ListIndex = 1
        frmEditThirdParty.cboLegalRepStatus.Enabled = False
    End If
    
    'Dim nLegalRepRole As Integer
    'nLegalRepRole = g_clsGlobalParameter.FindAmount("SPLegalRepRole")
    'If Not m_bIsEdit Then
        'frmEditThirdParty.cboLegalRepStatus.ListIndex = 1
        'frmEditThirdParty.cmdAllocatePanelId.Enabled = True
   ' Else
    'If m_clsLegalRep.GetLastActivatedDate = Now() Then
    'Dim clsuserrole As New userrole
    'clsuserrole.SetUserID m_sUserID
        
   ' Set clsTableAccess = m_clsuserrole
   ' clsTableAccess.SetKeyMatchValues g_sSupervisorUser
    'Dim rs As New ADODB.Recordset
    'Set rs = clsTableAccess.GetTableData()
    
    
        'if >= nLegalRepRole  then
        'End If
        
    'End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Loads all details from the tables into the screens controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    
    Dim vVal As Variant
    Dim vExpiry As Variant
    
    On Error GoTo Failed
    
    If m_bIsEdit = True Then
        ' Panel ID
        frmEditThirdParty.txtLegalRep(PANEL_ID).Text = m_clsPanel.GetPanelID()
        
        ' Associate Panel ID
        frmEditThirdParty.txtLegalRep(ASSOCIATE_PANEL_ID).Text = m_clsPanel.GetAssociatedPanelID()
        
        ' Payment Method
        vVal = m_clsPanel.GetPaymentMethod()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalPaymentMethod, vVal, SET_CONTROL_VALUE
        
        ' Type of Legal Rep
        vVal = m_clsLegalRep.GetTypeOfLegalRep()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalRepType, vVal, SET_CONTROL_VALUE
        
        '~ Legal Rep Status
        vVal = m_clsLegalRep.GetLegalRepStatus()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalRepStatus, vVal, SET_CONTROL_VALUE
        
        ' Number of Partners
        frmEditThirdParty.txtLegalRep(NUMBER_OF_PARTNERS).Text = m_clsLegalRep.GetNumberOfPartners()
        
        ' Senior Partner Details
        frmEditThirdParty.txtSeniorPartnerDetails.Text = m_clsLegalRep.GetSeniorPartnerDetails()
        
        ' Indemnity Insurance Expiry Date
        vExpiry = m_clsLegalRep.GetIndemnityInsuranceExpiryDate()
        g_clsFormProcessing.HandleDate frmEditThirdParty.txtLegalRep(INDEMNITY_EXPIRY_DATE), vExpiry, SET_CONTROL_VALUE
        
        ' Country of Practice
        vVal = m_clsLegalRep.GetCountryOfPractice()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboCountryOfPractice, vVal, SET_CONTROL_VALUE
        
        ' Country of Origin
        vVal = m_clsLegalRep.GetCountryOfOrigin()
        g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboCountryOfOrigin, vVal, SET_CONTROL_VALUE
        
        ' Temporary appointment indicator
        vVal = m_clsLegalRep.GetTempAppointment
        g_clsFormProcessing.HandleCheckBox frmEditThirdParty.chkTempAppointment, vVal, SET_CONTROL_VALUE
        
        ' TK 30/11/2005 MAR81 GetLastActivatedDate
        vVal = m_clsLegalRep.GetLastActivatedDate()
        frmEditThirdParty.SetLastActivatedDate CStr(vVal)
    
        'GetUserID
        vVal = m_clsLegalRep.GetUserID()
        frmEditThirdParty.SetLegalRepUserID CStr(vVal)
        ' TK 30/11/2005 MAR81 End
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all details from the relevant screen controls to table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim vVal As Variant
    Dim vExpiry As Variant
    
    On Error GoTo Failed
    
    ' Set the Directory GUID's
    m_clsPanel.SetDirectoryGUID m_vDirectoryGUID
    m_clsLegalRep.SetDirectoryGUID m_vDirectoryGUID
        
    ' Panel ID
    m_clsPanel.SetPanelID frmEditThirdParty.txtLegalRep(PANEL_ID).Text
    
    ' Associate Panel ID
    m_clsPanel.SetAssociatedPanelID frmEditThirdParty.txtLegalRep(ASSOCIATE_PANEL_ID).Text
    
    ' Payment Method
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalPaymentMethod, vVal, GET_CONTROL_VALUE
     m_clsPanel.SetPaymentMethod vVal
    
    ' Type of Legal Rep
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalRepType, vVal, GET_CONTROL_VALUE
    m_clsLegalRep.SetTypeOfLegalRep vVal
    
    '~ Legal Rep Status
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboLegalRepStatus, vVal, GET_CONTROL_VALUE
    m_clsLegalRep.SetLegalRepStatus vVal
    
    ' Number of Partners
    m_clsLegalRep.SetNumberOfPartners frmEditThirdParty.txtLegalRep(NUMBER_OF_PARTNERS).Text
    
    ' Senior Partner Details
    m_clsLegalRep.SetSeniorPartnerDetails frmEditThirdParty.txtSeniorPartnerDetails.Text
    
    ' Indemnity Insurance Expiry Date
    g_clsFormProcessing.HandleDate frmEditThirdParty.txtLegalRep(INDEMNITY_EXPIRY_DATE), vExpiry, GET_CONTROL_VALUE
    m_clsLegalRep.SetIndemnityInsuranceExpiryDate vExpiry

    ' Country of Practice
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboCountryOfPractice, vVal, GET_CONTROL_VALUE
    m_clsLegalRep.SetCountryOfPractice vVal
    
    ' Country of Origin
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboCountryOfOrigin, vVal, GET_CONTROL_VALUE
    m_clsLegalRep.SetCountryOfOrigin vVal
    
    ' Temporary appointment indicator
    g_clsFormProcessing.HandleCheckBox frmEditThirdParty.chkTempAppointment, vVal, GET_CONTROL_VALUE
    m_clsLegalRep.SetTempAppointment vVal
        
    'User ID
    m_clsLegalRep.SetUserID g_sSupervisorUser
    'TK 01/12/2005 MAR775 - Last Activated Date
    m_clsLegalRep.SetLastActivatedDate CStr(Now)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Update all underlying table objects handled by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    
    Dim clsTable As TableAccess
    
    On Error GoTo Failed

    'Update the Panel table.
    Set clsTable = m_clsPanel
    clsTable.Update

    'Update the LegalRep table.
    Set clsTable = m_clsLegalRep
    clsTable.Update

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates any control values relevant to this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    
    'Stub.
    'AM 28/09/2005
    ValidateScreenData = True
    
    If frmEditThirdParty.txtThirdParty(0).Text = "LEGAL REP" And frmEditThirdParty.txtLegalRep(0).Text = "Null" Then
       ' MsgBox "enable the button"
        frmEditThirdParty.cmdAllocatePanelId.SetFocus
        ValidateScreenData = False
    End If
    
   'If frmEditThirdParty.cboLegalRepType.List(frmEditThirdParty.cboLegalRepType.ListIndex) = "<Select>" Then
     ' MsgBox "Valid LegalRep is missing "
     ' frmEditThirdParty.cmdAllocatePanelId.SetFocus
     ' ValidateScreenData = False
    'End If
    
   ' If frmEditThirdParty.txtLegalRep(0).Text = "" Then
    '   MsgBox "Invalid Panel ID"
     '  frmEditThirdParty.cmdAllocatePanelId.SetFocus
     '  ValidateScreenData = False
    'End If
    
        
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

