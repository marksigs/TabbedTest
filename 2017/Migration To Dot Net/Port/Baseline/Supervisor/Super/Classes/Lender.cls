VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : Lender
' Description   : Controlling class for the Lender form.
'
' Change history
' Prog      Date        Description
' DJP       01/08/01    SYS2831 Created for Client variants.
' DJP       13/11/01    SYS2831 Add ability to override Postcode validation
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       14/02/02    SYS4094 Allow Lender to use Base Rate Structure. Added SelectItem and
'                       DeSelectItem
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const LENDER_CODE = 0
Private Const POST_CODE As Integer = 0
' Private data
Private m_bIsEdit As Boolean
Private m_colTabs As Collection
Private m_colKeys As Collection
Private m_nActiveTab  As LenderDetailsTabs
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Load
' Description   : Called when the form is first loaded. Needs to initialise all tabs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   BMIDS
'   AW      11/07/02    BMDS00177   Removed Ledger Codes tab
'
'

Public Sub Load()
    On Error GoTo Failed

    SetReturnCode MSGFailure
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm frmLenderDetails
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
    m_clsTableAccess.SetKeyMatchValues m_colKeys
End Sub

Public Sub Show()
    On Error GoTo Failed
    

    frmLenderDetails.Show vbModal, frmMain
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm frmLenderDetails
End Sub

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Public Sub UnloadForm()
    On Error GoTo Failed
    
    Unload frmLenderDetails
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub OK()
On Error GoTo Failed
    
    SetReturnCode
    SaveChangeRequest
    frmLenderDetails.Hide
    
Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub Cancel()
    On Error GoTo Failed
    
    frmLenderDetails.Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub SetButtonsState(nTab As LenderDetailsTabs)
    m_colTabs(nTab).SetButtonsState
End Sub

Public Sub InitialiseTab(nTab)
    On Error GoTo Failed
    
    m_colTabs(nTab + 1).InitTab
    SetTabstops frmLenderDetails
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub ValidateContactDetailsKey(ByVal nIndex As Integer, ByRef nKeyAscii As Integer)
    On Error GoTo Failed
    
    If nIndex = POST_CODE Then
        ValidatePostcodeKey nIndex, nKeyAscii
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub ValidatePostcodeKey(ByVal nIndex As Integer, ByRef nKeyAscii As Integer)
    On Error GoTo Failed
    
    nKeyAscii = Asc(UCase(Chr(nKeyAscii)))
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Add(nTab As LenderDetailsTabs)
    On Error GoTo Failed
    
    m_colTabs(nTab).Add
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Delete(nTab As LenderDetailsTabs)
    On Error GoTo Failed
    
    m_colTabs(nTab).Delete
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Edit(nTab As LenderDetailsTabs)
    On Error GoTo Failed
    
    m_colTabs(nTab).Edit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetActiveTab(Optional nTab As LenderDetailsTabs, Optional bSetNow As Boolean = True)
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = False
    
    If Not IsMissing(nTab) Then
        If nTab > 0 Then
            m_nActiveTab = nTab
            bRet = True
        End If
    ElseIf m_nActiveTab > 0 Then
        bRet = True
    End If
    
    If bSetNow Then
        If bRet = True Then
            frmLenderDetails.SSTabLenderDetails.Tab = m_nActiveTab - 1
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SetActiveTab: Unable to set tab, invalid value of " & m_nActiveTab
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CreateTabs()
On Error GoTo Failed
    Dim clsLegalFees As LegalFees
    Dim clsOtherFees As OtherFees
    Dim clsMIGRateSets As MIGRateSets
    'Dim     As LedgerCodes
    Dim clsLenderDetails As LenderDetails
    Dim clsContactDetails As ContactDetails
    Dim clsLenderParameters As AdditionalLenderParameters
    Dim clsPaymentDetails As LenderPaymentDetails
    
    Set m_colTabs = New Collection
    
    ' Create all Lender tab support classes
    Set clsLenderDetails = New LenderDetailsCS
    Set clsContactDetails = New ContactDetailsCS
    Set clsLenderParameters = New LenderParametersCS
    Set clsLegalFees = New LegalFeesCS
    Set clsOtherFees = New OtherFeesCS
    Set clsMIGRateSets = New MIGRateSetsCS
    'BMIDS00177
    'Set clsLedgerCodes = New LedgerCodesCS
    Set clsPaymentDetails = New LenderPaymentDetailsCS
    
    m_colTabs.Add clsLenderDetails
    m_colTabs.Add clsContactDetails
    m_colTabs.Add clsLenderParameters
    m_colTabs.Add clsLegalFees
    m_colTabs.Add clsOtherFees
    m_colTabs.Add clsMIGRateSets
    'm_colTabs.Add clsLedgerCodes
    
    'MSMS0069 - Add the payment details tab-handler.
    m_colTabs.Add clsPaymentDetails
    
    ' Default initialisation on the Lender details tab (the first to be shown)
    InitialiseTab LenderDetails - 1
    
    ' Contact Details is called from other places too, so set the correct form
    clsContactDetails.SetContactDetailsForm frmLenderDetails
    
    Set clsLenderDetails = m_colTabs(LenderDetails)
    
    ' First populate the Lender Details
    clsLenderDetails.SetTableClass m_clsTableAccess
    ' The LenderDetails class contains the methods to write to the Lender Details tab.
    
    ' Need to populate the Lender Details tab first so we can get the keys from it for the
    ' rest of the tabs - the ones that need keys, that is.
    clsLenderDetails.Initialise m_bIsEdit

    ' If we're adding, create the keys collection
    If Not m_bIsEdit Then
        Dim vOrgID As Variant
        
        Set m_colKeys = New Collection
        vOrgID = clsLenderDetails.GetOrganisationID
        m_colKeys.Add vOrgID
    End If

    'MSMS0069 - Associate the lender table with payment processing tab class.
    clsPaymentDetails.SetTableClass m_clsTableAccess

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitialiseTabs()
    On Error GoTo Failed
    Dim nThisTab As Integer
    Dim nTabCount As Integer
    
    ' Set the first tab the user will see to be the Lender Details tab, but only if a tab
    ' hasn't been specified before.
    If m_nActiveTab = -1 Then
        m_nActiveTab = LenderDetails
    End If
    SetActiveTab m_nActiveTab
    
    nTabCount = m_colTabs.Count
    
    ' Start at the tab after the Lender Details tab as this has been done already
    nThisTab = LenderDetails + 1
    
    ' Loop through all tabs and initialise them
    While nThisTab <= nTabCount
        m_colTabs(nThisTab).Initialise m_bIsEdit
        nThisTab = nThisTab + 1
    Wend
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetTabsCollection() As Collection
    On Error GoTo Failed
    
    If m_colTabs Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Lender tab list is empty"
    End If
    
    Set GetTabsCollection = m_colTabs
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SetContactDetailKeys(clsLenderDetails As LenderDetails)
On Error GoTo Failed
    
    Dim vContactDetailsGUID As Variant
    Dim colValues As New Collection
    Dim clsContactDetails As ContactDetails
    
    Const NUMBER_OF_KEYS = 1
    
    ' Get the Contact Details class for the Contact Details tab.
    Set clsContactDetails = m_colTabs(ContactDetails)
    
    ' To populate the Contant Details tab we need the Contact Details GUID
    vContactDetailsGUID = clsLenderDetails.GetContactDetailsGUID()
    
    If Not IsNull(vContactDetailsGUID) Then
        If Len(vContactDetailsGUID) > 0 Then
            colValues.Add vContactDetailsGUID, CONTACT_DETAILS_KEY
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to set contact detail keys - ContactDetailsGUID is empty"
    End If
    
    If colValues.Count = NUMBER_OF_KEYS Then
        clsContactDetails.SetContactKeyValues colValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetContactDetails: ContactDetailsGUID must be valid"
    End If

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetAddressKeys(clsLenderDetails As LenderDetails)
On Error GoTo Failed
    Dim vAddressGUID As Variant
    Dim colValues As New Collection
    Dim clsContactDetails As ContactDetails
    Const NUMBER_OF_KEYS = 1
    
    Set clsContactDetails = m_colTabs(ContactDetails)
    
    vAddressGUID = clsLenderDetails.GetAddressGUID()
    
    If Not IsNull(vAddressGUID) Then
        If Len(vAddressGUID) > 0 Then
            colValues.Add vAddressGUID, ADDRESS_KEY
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Address GUID is empty"
    End If
    
    If colValues.Count = NUMBER_OF_KEYS Then
        clsContactDetails.SetAddressKeyValues colValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetContactDetails: AddressGUID must be valid"
    End If
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function ValidateScreenData() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean
    Dim nColIndex As Integer
    
    nColIndex = -1
    bRet = g_clsFormProcessing.DoMandatoryProcessing(frmLenderDetails)
    
    If bRet Then
        nColIndex = 1
        
        While nColIndex <= m_colTabs.Count
            m_colTabs(nColIndex).ValidateScreenData
            nColIndex = nColIndex + 1
        Wend
    End If
    
    ValidateScreenData = bRet

Exit Function
Failed:
    
    g_clsErrorHandling.SaveError Err.Number, Err.DESCRIPTION
    If nColIndex <> -1 Then
        frmLenderDetails.SSTabLenderDetails.Tab = nColIndex - 1
    End If
    
    g_clsErrorHandling.RaiseError
End Function

Public Sub PopulateScreenFields()
On Error GoTo Failed
    Dim nTabCount As Integer
    Dim nThisTab As Integer
    
    nTabCount = m_colTabs.Count
    
    For nThisTab = 1 To nTabCount
        m_colTabs(nThisTab).SetScreenFields
    Next
        
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()

On Error GoTo Failed
    Dim sRateSet As String
    Dim nRecordCount As Long
    Dim clsMIGRates As MIGRateSets
    Dim clsLegalFees As LegalFees
    Dim clsLenderDetails As LenderDetails
    
    Set clsMIGRates = m_colTabs(MIGRateSets)
    Set clsLegalFees = m_colTabs(LegalFees)
    Set clsLenderDetails = m_colTabs(LenderDetails)
    
    ' LenderDetails needs the RateSet from the MIGRates class
    sRateSet = clsMIGRates.GetRateSet
    clsLenderDetails.SetMIGRateKey sRateSet

    ' Save data
    clsLenderDetails.SaveScreenData
    m_colTabs(ContactDetails).SaveScreenData
    m_colTabs(AdditionalParameters).SaveScreenData
    clsLegalFees.SaveScreenData
    
    'MSMS0069 - Save payment details.
    m_colTabs(LenderPaymentDetails).SaveScreenData
    
    ' And update it
    m_colTabs(ContactDetails).DoUpdates
    clsLenderDetails.DoUpdates
    clsMIGRates.DoUpdates
    m_colTabs(LegalFees).DoUpdates

    m_colTabs(AdditionalParameters).DoUpdates
    m_colTabs(OtherFees).DoUpdates
    'BMIDS00177
'    m_colTabs(LedgerCodes).DoUpdates
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub Initialise()

    On Error GoTo Failed
    Const FORM_EDIT As Boolean = True
    Dim clsLenderDetails As LenderDetails
    
    ' Now do the Initialisation
    Set clsLenderDetails = m_colTabs(LenderDetails)
    
    ' Get the keys for the Contact Details tab - pass in the LenderDetails tab because
    ' that's where we'll get the keys from
    
    SetContactDetailKeys clsLenderDetails
    SetAddressKeys m_colTabs(LenderDetails)
    
    ' Additional Parameters
    ' This is keyed on OrganisationID
    SetAdditionalParameterKeys m_colTabs(LenderDetails)
        
    ' Need to be able to pass the lender code across to the LegalFee screen so it can be used
    ' to a) display the lender code on the screen in an edit, and also when the user presses
    ' ok we need to use the lender code the user had entered as a key for the LegalFeeSet
    ' table.
    SetOtherFeeKeys
    SetLegalFeeKeys
    SetMIGKeys
    'BMIDS00177
    'SetLedgerCodeKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetAdditionalParameterKeys(clsLenderDetails As LenderDetails)
    On Error GoTo Failed
    Dim colValues As Collection
    Dim vOrganisationID As Variant
    Dim clsAdditionalParameters As AdditionalLenderParameters
    
    On Error GoTo Failed
    Const NUMBER_OF_KEYS = 1
    
    Set colValues = New Collection
    
    ' Get the organisation ID as this is the key for the additional parameters table
    vOrganisationID = clsLenderDetails.GetOrganisationID()

    If Not IsNull(vOrganisationID) Then
        If Len(vOrganisationID) > 0 Then
            colValues.Add vOrganisationID
        
            If g_clsVersion.DoesVersioningExist() Then
                colValues.Add g_sVersionNumber
            End If
        
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot set Additional Parameter keys - organisationID is empty"
    End If
    
    Set clsAdditionalParameters = m_colTabs(AdditionalParameters)
    
    If colValues.Count > 0 Then
        clsAdditionalParameters.SetKeyValues colValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "SetAdditionalParameterKeys: Key count is invalid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetOtherFeeKeys()
    On Error GoTo Failed
    Dim clsOtherFees As OtherFees
    
    ' OtherFees key
    Set clsOtherFees = m_colTabs(OtherFees)
    clsOtherFees.SetKeyValues m_colKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetLegalFeeKeys()
    On Error GoTo Failed
    Dim vOrgID As Variant
    Dim clsLegalFees As LegalFees
    
    Set clsLegalFees = m_colTabs(LegalFees)
    
    ' Legal Fees - OrganisationID, always set
    vOrgID = m_colKeys(ORGANISATION_ID_KEY)
    clsLegalFees.SetOrganisationID vOrgID
     
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetMIGKeys()
    On Error GoTo Failed
    Dim vRateSet As Variant
    Dim clsMIGRates As MIGRateSets
    Dim clsLenderDetails As LenderDetails
    
    Set clsMIGRates = m_colTabs(MIGRateSets)
    Set clsLenderDetails = m_colTabs(LenderDetails)
    
    ' MIG Key is RateSet
    vRateSet = clsLenderDetails.GetRateSet()
    
    ' MIG Rates Key
    clsMIGRates.SetMIGRateSet vRateSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetLedgerCodeKeys()
    On Error GoTo Failed
    
    'm_colTabs(LedgerCodes).SetColKeys m_colKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)

    m_ReturnCode = enumReturn
    
End Sub

Public Function GetReturnCode() As MSGReturnCode

    GetReturnCode = m_ReturnCode
    
End Function

Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim sThisVal As Variant
    Dim colValues As Collection
    
    sDesc = frmLenderDetails.txtLenderDetails(LENDER_CODE).Text
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess, sDesc
    
Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub HandlePostcode(ByVal nIndex As Integer, ByRef nKeyAscii As Integer)
    On Error GoTo Failed
    
    If nIndex = POST_CODE Then
        nKeyAscii = Asc(UCase(Chr(nKeyAscii)))
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetKeyMatchValues() As Variant
    On Error GoTo Failed

    If Not m_colKeys Is Nothing Then
        Set GetKeyMatchValues = m_colKeys
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "Lender.GetKeyMatchValues"
    End If

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub Class_Terminate()
    Set m_clsTableAccess = Nothing
End Sub

Private Sub Class_Initialize()
    m_nActiveTab = -1
    ' The MortgageLender table is the starting point for all Lender data
    Set m_clsTableAccess = New MortgageLendersTable
End Sub

Public Sub SelectItem(nItem As Long)
    On Error GoTo Failed
    
    m_colTabs(nItem).Selected
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeselectItem(nItem As Long)
    On Error GoTo Failed
    
    m_colTabs(nItem).DeSelect
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

