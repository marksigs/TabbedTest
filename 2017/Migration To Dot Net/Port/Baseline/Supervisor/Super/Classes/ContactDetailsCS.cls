VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactDetailsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ContactDetailsCS
' Description   : Client specific version of the ContactDetails class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    SYS2831 Created for client variant support
' DJP       10/01/02    SYS2831 Support client variants. Added SetUKSpecificData and SetUKSpecificData
' DJP       10/01/02    SYS2831 Support client variants. Added GetContactDetailsForm
' STB       13/02/02    SYS4055 ThirdParties may not always have a contact record. Added GetContactTelephoneDetails().
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' The base class for base class support
Implements ContactDetails

' Private data
Private m_clsContactDetails As ContactDetails
Private Sub Class_Initialize()
    Set m_clsContactDetails = New ContactDetails
End Sub
Private Sub ContactDetails_DoUpdates()
    m_clsContactDetails.DoUpdates
End Sub
Private Function ContactDetails_GetAddressDetails() As AddressTable
    Set ContactDetails_GetAddressDetails = m_clsContactDetails.GetAddressDetails
End Function
Private Function ContactDetails_GetAddressGUID() As Variant
    ContactDetails_GetAddressGUID = m_clsContactDetails.GetAddressGUID
End Function
Private Function ContactDetails_GetContactDetails() As ContactDetailsTable
    Set ContactDetails_GetContactDetails = m_clsContactDetails.GetContactDetails
End Function
Private Function ContactDetails_GetContactDetailsForm() As Form
    Set ContactDetails_GetContactDetailsForm = m_clsContactDetails.GetContactDetailsForm
End Function
Private Function ContactDetails_GetContactDetailsGUID() As Variant
    ContactDetails_GetContactDetailsGUID = m_clsContactDetails.GetContactDetailsGUID
End Function
Private Function ContactDetails_GetContactTelephoneDetails() As ContactDetailsTelephoneTable
    Set ContactDetails_GetContactTelephoneDetails = m_clsContactDetails.GetContactTelephoneDetails
End Function

Private Sub ContactDetails_Initialise(bIsEdit As Boolean)
    m_clsContactDetails.Initialise bIsEdit
End Sub
Private Sub ContactDetails_InitTab()
    m_clsContactDetails.InitTab
End Sub
Private Function ContactDetails_IsEdit() As Boolean
    ContactDetails_IsEdit = m_clsContactDetails.IsEdit
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ContactDetails_SaveScreenData
' Description   : Client specific version of SaveScreenData - do any client specific saves, then
'                 delegate to the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ContactDetails_SaveScreenData()
    On Error GoTo Failed
    
    ' Client specific, in this case Core UK specific save
    ContactDetails_SaveUKSpecificData
    
    ' Delegate to core.
    m_clsContactDetails.SaveScreenData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ContactDetails_SaveUKSpecificData()
    On Error GoTo Failed
    
    m_clsContactDetails.SaveUKSpecificData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ContactDetails_SetAddressKeyValues(colValues As Collection)
    m_clsContactDetails.SetAddressKeyValues colValues
End Sub
Private Sub ContactDetails_SetContactDetailsForm(frmContactDetails As Form)
    m_clsContactDetails.SetContactDetailsForm frmContactDetails
End Sub
Private Sub ContactDetails_SetContactKeyValues(colValues As Collection)
    m_clsContactDetails.SetContactKeyValues colValues
End Sub
Private Sub ContactDetails_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsContactDetails.SetIsEdit bEdit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ContactDetails_SetScreenFields
' Description   : Client specific version of SetScreenFields - do any client specific loads, then
'                 delegate to the Core method.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ContactDetails_SetScreenFields()
    On Error GoTo Failed
    
    ' Client specific sets
    ContactDetails_SetUKSpecificData
    
    ' Delegate to the Core method
    m_clsContactDetails.SetScreenFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ContactDetails_SetUKSpecificData()
    On Error GoTo Failed
    
    m_clsContactDetails.SetUKSpecificData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub ContactDetails_ValidatePostCode()
    m_clsContactDetails.ValidatePostCode
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ContactDetails_ValidateScreenData
' Description   : Performs validation of Contact Details and Address details entered onto a screen
'                 that implements this class. If ValidatePostCode is to be overridden, it can be
'                 added to this class in a client variant and overridden.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ContactDetails_ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bSuccess As Boolean
    
    ' Delegate to Core validate method
    bSuccess = m_clsContactDetails.ValidateScreenData

    ' If Successful, do Postcode check
    If bSuccess Then
        m_clsContactDetails.ValidatePostCode
    End If

    ContactDetails_ValidateScreenData = bSuccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
