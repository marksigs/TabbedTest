VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThirdPartyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ThirdPartyDetails
' Description   : Class that supports the ThirdParty Main details tab on the third party screen.
'
' Change history
' Prog      Date        Description
' DJP       21/06/01    As part of SQL Server port, found bug where data wasn't being saved correctly
'                       on OK.
' STB       29/11/01    SYS2912 - Removed the frmMainThirdParty screen and
'                       ported its functionality to frmEditThirdParty.
' STB       07/01/02    SYS3581 - Initialise() now sets the keys collection if adding.
' AW        08/01/02    SYS3560 - Set OtherSystemNumber
' STB       13/02/02    SYS4054 - Added Branchname field.
' STB       13/02/02    SYS4055 - Added SetAddressGUID() and SetContactGUID().
' STB       10/02/02    SYS4302 - Updated OTHER_TYPE so it doesn't conflict with 'O'.
' STB       15/05/02    SYS4558 - Export Legal Reps and Valuers to Optimus.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Text Edit constants
Private Const COMPANY_NAME  As Long = 0
Private Const ACTIVE_FROM   As Long = 1
Private Const ACTIVE_TO     As Long = 2
Private Const DX_ID         As Long = 3
Private Const DX_LOCATION   As Long = 4
Private Const BANK_SORT_CODE  As Long = 5
Private Const OTHER_SYSTEM_NO As Long = 6
Private Const BANK_BRANCHNAME As Long = 7
Private Const LEGALREP_TYPE   As String = "T"
Private Const VALUER_TYPE     As String = "P"
Private Const OTHER_TYPE      As String = "TO"

Private Const POST_CODE     As Long = 0
Private Const BUILDING_NAME As Long = 1
Private Const BUILDING_NUMBER As Long = 2
Private Const STREET        As Long = 3
Private Const DISTRICT      As Long = 4
Private Const TOWN          As Long = 5
Private Const COUNTY        As Long = 6
Private Const SURNAME       As Long = 7
Private Const FORENAME      As Long = 8
Private Const COUNTRY_CODE1 As Long = 9
Private Const AREA_CODE1    As Long = 10
Private Const TELEPHONE1    As Long = 11
Private Const EXTENSION1    As Long = 12
Private Const COUNTRY_CODE2 As Long = 13
Private Const AREA_CODE2    As Long = 14
Private Const TELEPHONE2    As Long = 15
Private Const EXTENSION2    As Long = 16
Private Const EMAIL_ADDRESS As Long = 17

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'Local table classes. TODO: Currently, this is created and loaded by the form,
'to be consistant, this should be done in .Initialise()...
Private m_clsNameAndAddressDir As NameAndAddressDirTable

'The third party type (Local, Legal Rep, or Valuer). This is set from the form
'So this tab-handler can enable/disable the type combo(s).
Private m_uThirdPartyType As ThirdPartyType


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Indicates if this tab-handler is in add or edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' BMIDS
' AW    11/07/02    BMIDS00177  Removed Optimus refs
' SA    18/11/2002  BMIDS00987  Save hyphens in sort code

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetIsEdit(Optional ByVal bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Sub Class_Initialize()
    
    'TODO: This is currently redundant as the form passes an existing object into
    'this class.
    Set m_clsNameAndAddressDir = New NameAndAddressDirTable
    
End Sub

Public Sub SetTableClass(ByRef clsTableAccess As TableAccess)
    Set m_clsNameAndAddressDir = clsTableAccess
End Sub

Public Sub Initialise(ByVal bIsEdit As Boolean)
        
    Dim sError As String
    Dim colKeys As Collection
    Dim rs As ADODB.Recordset
    Dim nRecordCount As Integer
    Dim sDirectoryGUID As String
    Dim clsTableAccess As TableAccess
        
    On Error GoTo Failed
    
    sError = "ThirdPartyDetails:Initialise - "
    
    'Store the add/edit state of the form.
    m_bIsEdit = bIsEdit
        
    If bIsEdit = True Then
        Set clsTableAccess = m_clsNameAndAddressDir
        Set rs = clsTableAccess.GetTableData()
        
        clsTableAccess.ValidateData
        
        nRecordCount = clsTableAccess.RecordCount
        
        If nRecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Incorrect number of Lender records returned: " & nRecordCount
        End If
        
        'Enable/Disable Hide/Show controls based upon form state.
        SetEditState
    Else
        'If we're adding, create a new record and setup the GUIDs.
        g_clsFormProcessing.CreateNewRecord m_clsNameAndAddressDir
                
        m_clsNameAndAddressDir.SetDirectoryGUID
        m_clsNameAndAddressDir.SetAddressGUID
        m_clsNameAndAddressDir.SetContactDetailsGUID
        
        'Set the keys collection now otherwise promotions won't work.
        Set colKeys = New Collection
        
        'Add the generated GUID into the keys collection.
        colKeys.Add m_clsNameAndAddressDir.GetDirectoryGUID
        
        'Assign the keys collection to the table.
        TableAccess(m_clsNameAndAddressDir).SetKeyMatchValues colKeys
        
        'Enable/Disable Hide/Show controls based upon form state.
        SetAddState
    End If
    
    'Populate the screen combo's.
    g_clsFormProcessing.PopulateCombo "OrganisationType", frmEditThirdParty.cboOrgType
    
    'This source of the values for the Address Type combo are based upon the third-party-type.
    PopulateAddressCombo
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetScreenFields()
    
    Dim vVal As Variant
    Dim vActiveFrom As Variant
    Dim vActiveTo As Variant
    
    On Error GoTo Failed
    
    frmEditThirdParty.txtThirdParty(COMPANY_NAME).Text = m_clsNameAndAddressDir.GetCompanyName()
    
    vActiveFrom = m_clsNameAndAddressDir.GetActiveFrom()
    g_clsFormProcessing.HandleDate frmEditThirdParty.txtThirdParty(ACTIVE_FROM), vActiveFrom, SET_CONTROL_VALUE
    
     vActiveTo = m_clsNameAndAddressDir.GetActiveTo()
    g_clsFormProcessing.HandleDate frmEditThirdParty.txtThirdParty(ACTIVE_TO), vActiveTo, SET_CONTROL_VALUE
    
    frmEditThirdParty.txtThirdParty(COMPANY_NAME).Text = m_clsNameAndAddressDir.GetCompanyName()
    frmEditThirdParty.txtThirdParty(DX_ID).Text = m_clsNameAndAddressDir.GetDXID()
    frmEditThirdParty.txtThirdParty(DX_LOCATION).Text = m_clsNameAndAddressDir.GetDXLocation()
    frmEditThirdParty.txtThirdParty(BANK_SORT_CODE).Text = m_clsNameAndAddressDir.GetBankSortCode()
    frmEditThirdParty.txtThirdParty(BANK_BRANCHNAME).Text = m_clsNameAndAddressDir.GetBranchName()

    vVal = m_clsNameAndAddressDir.GetOrganisationType()
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboOrgType, vVal, SET_CONTROL_VALUE
    
    vVal = m_clsNameAndAddressDir.GetNameAndAddressType()
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboAddressType, vVal, SET_CONTROL_VALUE
    
    g_clsFormProcessing.HandleCheckBox frmEditThirdParty.chkHeadOffice, m_clsNameAndAddressDir.GetHeadOfficeIndicator, SET_CONTROL_VALUE
    frmEditThirdParty.txtNotes.Text = m_clsNameAndAddressDir.GetNotes()
    
    'BMIDS00177
    'frmEditThirdParty.txtThirdParty(OTHER_SYSTEM_NO).Text = m_clsNameAndAddressDir.GetOtherSystemNumber()
    
    'Enable/Disable the organisation combo based upon the current address/-
    'thirdparty-type combo selection.
    frmEditThirdParty.SetOrganisationComboState
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SaveScreenData()
    
    Dim vVal As Variant
    Dim strValidationID
    Dim nValuer As Integer
    Dim vActiveTo As Variant
    Dim nLegalType As Integer
    Dim vActiveFrom As Variant
    Dim clsValidation As ComboValidationTable
    
    On Error GoTo Failed
    
    Set clsValidation = New ComboValidationTable
      
    nLegalType = clsValidation.GetSingleValueFromValidation("ThirdPartyType", LEGALREP_TYPE)
    nValuer = clsValidation.GetSingleValueFromValidation("ThirdPartyType", VALUER_TYPE)
    
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboAddressType, vVal, GET_CONTROL_VALUE
        
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboOrgType, vVal, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetOrganisationType vVal
    
    g_clsFormProcessing.HandleComboExtra frmEditThirdParty.cboAddressType, vVal, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetNameAndAddressType vVal
    
    m_clsNameAndAddressDir.SetCompanyName frmEditThirdParty.txtThirdParty(COMPANY_NAME).Text
    
    g_clsFormProcessing.HandleDate frmEditThirdParty.txtThirdParty(ACTIVE_FROM), vActiveFrom, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetActiveFrom vActiveFrom
    
    g_clsFormProcessing.HandleDate frmEditThirdParty.txtThirdParty(ACTIVE_TO), vActiveTo, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetActiveTo vActiveTo
    
    m_clsNameAndAddressDir.SetCompanyName frmEditThirdParty.txtThirdParty(COMPANY_NAME).Text
    
    m_clsNameAndAddressDir.SetDXID frmEditThirdParty.txtThirdParty(DX_ID).Text
    m_clsNameAndAddressDir.SetDXLocation frmEditThirdParty.txtThirdParty(DX_LOCATION).Text
    m_clsNameAndAddressDir.SetBankSortCode frmEditThirdParty.txtThirdParty(BANK_SORT_CODE).Text
    m_clsNameAndAddressDir.SetBranchName frmEditThirdParty.txtThirdParty(BANK_BRANCHNAME).Text
    
    g_clsFormProcessing.HandleCheckBox frmEditThirdParty.chkHeadOffice, vVal, GET_CONTROL_VALUE
    m_clsNameAndAddressDir.SetHeadOfficeIndicator vVal
    
    m_clsNameAndAddressDir.SetNotes frmEditThirdParty.txtNotes.Text
    
    ' DJP SQL Server port - this isn't available yet, so don't call it.
    'OutputOmiplusData
    UpdateODIData
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DoUpdates()

    Dim clsTable As TableAccess

    On Error GoTo Failed
    
    'Update the NameAndAddressDirectory table.
    Set clsTable = m_clsNameAndAddressDir
    clsTable.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Function GetContactDetailsGUID() As Variant
    GetContactDetailsGUID = m_clsNameAndAddressDir.GetContactDetailsGUID()
End Function

Friend Function GetAddressGUID() As Variant
    GetAddressGUID = m_clsNameAndAddressDir.GetAddressGUID()
End Function

Friend Function GetThirdPartyType() As Integer
    
    Dim vType As Variant
    
    vType = m_clsNameAndAddressDir.GetNameAndAddressType()

    If Not IsEmpty(vType) And Not IsNull(vType) Then
       GetThirdPartyType = CInt(vType)
    Else
       GetThirdPartyType = -1
    End If

End Function

Public Function ValidateScreenData() As Boolean
    
    Dim bRet As Boolean
    
    ' First, the End Date must not be earlier than the startdate
    bRet = g_clsValidation.ValidateActiveFromTo(frmEditThirdParty.txtThirdParty(ACTIVE_FROM), frmEditThirdParty.txtThirdParty(ACTIVE_TO))
        
    'BMIDS00987 Now we have removed the mask, we need to check bank sort code is the correct format {
    If bRet Then
       bRet = g_clsValidation.ValidateSortCode(frmEditThirdParty.txtThirdParty(BANK_SORT_CODE))
    End If
    'BMIDS00987 }
    ValidateScreenData = bRet
    
End Function

Public Sub SetAddState()
    
    'If the ThirdPartyType is a Valuer or a Legal Rep. then disable the type
    'combos.
    If (m_uThirdPartyType = ThirdPartyLegalRepType) Or (m_uThirdPartyType = ThirdPartyValuersType) Then
        frmEditThirdParty.cboOrgType.Enabled = False
        frmEditThirdParty.cboAddressType.Enabled = False
    Else
        frmEditThirdParty.cboOrgType.Enabled = True
        frmEditThirdParty.cboAddressType.Enabled = True
    End If
    
End Sub

Public Sub SetEditState()
    
    'Disable the address-type combo while editing.
    frmEditThirdParty.cboAddressType.Enabled = False
        
End Sub

Public Sub SetThirdPartyType(ByVal uType As ThirdPartyType)
    m_uThirdPartyType = uType
End Sub

Private Sub PopulateAddressCombo()

    Dim sSQL As String
    Dim rsAddesseTypes As ADODB.Recordset
    
    'Build the SQL to pull all the address types from the database.
    Select Case m_uThirdPartyType
        'If the type is a value or legal rep, we can populate the combo normally
        'as it will always be disabled (regardless of add/edit state).
        Case ThirdPartyValuersType, ThirdPartyLegalRepType
            g_clsFormProcessing.PopulateCombo "ThirdPartyType", frmEditThirdParty.cboAddressType
        
        'If the type is 'Other' the filter the values we populate it with.
        Case ThirdPartyLocalType
            g_clsFormProcessing.PopulateComboByValidation frmEditThirdParty.cboAddressType, "ThirdPartyType", OTHER_TYPE
            
        Case Else
    End Select
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAddressGUID(ByVal vAddressGUID As Variant)
    m_clsNameAndAddressDir.SetAddressGUID vAddressGUID
End Sub

Public Sub SetContactGUID(ByVal vContactGUID As Variant)
    m_clsNameAndAddressDir.SetContactDetailsGUID vContactGUID
End Sub


Private Sub UpdateODIData()

End Sub


