VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : ContactDetails
' Description   : Generic class which allows all contact details and address
'                 details to be saved from one place. All forms that use this
'                 class must use the same id's and indexes that are defined in
'                 the Constants section of this class.
'
' Change history
' Prog      Date        Description
' AA        12/06/2001  Added functions GetContactDetails & GetAddressDetails
' DJP       28/08/01    DJP SYS2564 MCAP Core functionality
' DJP       22/10/01    DJP SYS2831 Add functionality to suppport client
'                       variants.
' STB       08/11/01    Added telephone table functionality.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server
'                       locking problem.
' DJP       10/01/02    SYS2831 Support client variants. Added
'                       SetUKSpecificData and LoadUKSpecificData
' DJP       10/01/02    SYS2831 Support client variants. Added
'                       GetContactDetailsForm.
' STB       13/02/02    SYS4055 ThirdParties may not always have a contact
'                       record. Added GetContactTelephoneDetails().
' STB       22/02/02    SYS4091 Altered TelephoneUsage combo to use the -
'                       ContactTelephoneUsage combo and a single telephone
'                       record can now be loaded.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM History
' EP1033  HMA  26/07/2006  Add Other Title field.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Control indexes.
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
Private Const OTHER_TITLE   As Long = 18      'EP1033

'The number of telephone records per contact.
Private Const NUM_TELEPHONE_RECORDS As Long = 2

'The add/edit state of the form.
Private m_bIsEdit As Boolean

'A reference to the form containing the contact details controls.
Private m_frmContactDetails As Form

'The primary key of any address record.
Private m_vAddressGUID As Variant

'The primary key of any contact record.
Private m_vContactDetailsGUID As Variant

'Underlying address table.
Private m_clsAddressTable As AddressTable

'Underlying contact table.
Private m_clsContactDetailsTable As ContactDetailsTable

'Underlying telephone table.
Private m_clsContactDetailsTelephoneTable As ContactDetailsTelephoneTable

'Flag indicating if the contact table exists in the database.
Private m_bContactDetailsExists As Boolean


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Instantiate table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

    On Error GoTo Failed
        
    'Create the required underlying table classes for a contact.
    Set m_clsAddressTable = New AddressTable
    Set m_clsContactDetailsTable = New ContactDetailsTable
    Set m_clsContactDetailsTelephoneTable = New ContactDetailsTelephoneTable
    
    'Store a flag to indicate if the base contact details table was found.
    m_bContactDetailsExists = g_clsDataAccess.DoesTableExist(TableAccess(m_clsContactDetailsTable).GetTable())
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called by a form to load all contact and address details from
'                 the database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    
    On Error GoTo Failed
    
    'Only load table data if the base table exists in the database.
    If m_bContactDetailsExists Then
        'Store the objects edit mode.
        m_bIsEdit = bIsEdit
        
        'Only load data if we're editing.
        If bIsEdit Then
            SetEditState
        Else
            SetAddState
        End If
        
        'Load combo details.
        g_clsFormProcessing.PopulateCombo "Country", m_frmContactDetails.cboCountry
        g_clsFormProcessing.PopulateCombo "Title", m_frmContactDetails.cboContactTitle
        g_clsFormProcessing.PopulateCombo "ContactType", m_frmContactDetails.cboContactType
        g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", m_frmContactDetails.cboType1
        g_clsFormProcessing.PopulateCombo "ContactTelephoneUsage", m_frmContactDetails.cboType2
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "ContactDetails:Initialise  - " + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Called to initialise the data for the screen in add mode.
'                 Needs to create Address and ContactDetails GUIDS and
'                 associated records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    
    Dim vAddressGUID As Variant
    Dim vContactDetailsGUID As Variant
    
    On Error GoTo Failed
    
    ' First Address and set the key
    g_clsFormProcessing.CreateNewRecord m_clsAddressTable
    m_clsAddressTable.SetAddressGUID m_vAddressGUID
    
    ' Next, create Contact Details and key
    g_clsFormProcessing.CreateNewRecord m_clsContactDetailsTable
    m_clsContactDetailsTable.SetContactDetailsGUID m_vContactDetailsGUID
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Called to initialise the data for the screen in edit mode.
'                 Needs to go to the database to retrieve the associated
'                 records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    Dim colKeyMatchValues As Collection
    
    On Error GoTo Failed
        
    If IsNull(m_vContactDetailsGUID) Or (IsEmpty(m_vContactDetailsGUID)) Then
        'Create a blank contact record (and set the GUID).
        g_clsFormProcessing.CreateNewRecord m_clsContactDetailsTable
        m_vContactDetailsGUID = g_clsSQLAssistSP.GuidStringToByteArray(m_clsContactDetailsTable.SetContactDetailsGUID)
                
        'Create a keys collection.
        Set colKeyMatchValues = New Collection
        colKeyMatchValues.Add m_vContactDetailsGUID
        
        'Associate the keys collection against the contact table.
        TableAccess(m_clsContactDetailsTable).SetKeyMatchValues colKeyMatchValues
        
        'Create the first blank telephone record.
        g_clsFormProcessing.CreateNewRecord m_clsContactDetailsTelephoneTable
        m_clsContactDetailsTelephoneTable.SetCONTACTDETAILSTELEPHONEGUID m_vContactDetailsGUID
        m_clsContactDetailsTelephoneTable.SetTELEPHONE_SEQ_NUMBER 1
                
        'Create the first second telephone record.
        g_clsFormProcessing.CreateNewRecord m_clsContactDetailsTelephoneTable
        m_clsContactDetailsTelephoneTable.SetCONTACTDETAILSTELEPHONEGUID m_vContactDetailsGUID
        m_clsContactDetailsTelephoneTable.SetTELEPHONE_SEQ_NUMBER 2
        
        'Associate the keys collection against the telephone table.
        TableAccess(m_clsContactDetailsTelephoneTable).SetKeyMatchValues colKeyMatchValues
    
    'Only attempt to load a contact record if we have a GUID.
    Else
        'First, get Contact Details.
        TableAccess(m_clsContactDetailsTable).GetTableData

        'Ensure a single record was found.
        If TableAccess(m_clsContactDetailsTable).RecordCount <> 1 Then
            g_clsErrorHandling.RaiseError errGeneralError, "ContactDetails:Initialise - " & TableAccess(m_clsContactDetailsTable).RecordCount & " records returned"
        End If
        
        'Now for the telephone numbers.
        TableAccess(m_clsContactDetailsTelephoneTable).GetTableData
    End If
    
    'Now for the Address.
    TableAccess(m_clsAddressTable).GetTableData
    
    'Ensure a single record was found.
    If TableAccess(m_clsAddressTable).RecordCount <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "ContactDetails:Initialise - " & TableAccess(m_clsContactDetailsTable).RecordCount & " records returned"
    End If
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Performs validation of Contact Details and Address details
'                 entered onto a screen that implements this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    
    Dim bValid As Boolean
    
    On Error GoTo Failed

    bValid = True
    ValidateScreenData = bValid
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidatePostCode
' Description   : Performs validation of the postcode the user has entered.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ValidatePostCode()
    
    Dim bValid As Boolean
    
    On Error GoTo Failed

    bValid = g_clsValidation.ValidatePostCode(m_frmContactDetails.txtContactDetails(POST_CODE).Text)
    
    If Not bValid Then
        m_frmContactDetails.txtContactDetails(POST_CODE).SetFocus
        g_clsErrorHandling.RaiseError errPostCodeInvalid
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Called by a form to populate the fields associated with
'                 Address and Contact Details. The form MUST have controls of
'                 the same name and same indexes for this generic code to work.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    
    Dim vval As Variant
    
    On Error GoTo Failed
    
    If m_bContactDetailsExists Then
        'Update screen elements from the Address table.
        m_frmContactDetails.txtContactDetails(POST_CODE).Text = m_clsAddressTable.GetPostCode()
        m_frmContactDetails.txtContactDetails(BUILDING_NAME).Text = m_clsAddressTable.GetBuildingOrHouseName()
        m_frmContactDetails.txtContactDetails(BUILDING_NUMBER).Text = m_clsAddressTable.GetBuildingOrHouseNumber()
        m_frmContactDetails.txtContactDetails(STREET).Text = m_clsAddressTable.GetStreet()
        m_frmContactDetails.txtContactDetails(DISTRICT).Text = m_clsAddressTable.GetDistrict()
        m_frmContactDetails.txtContactDetails(TOWN).Text = m_clsAddressTable.GetTown()
    
        'Set the Country combo selection.
        vval = m_clsAddressTable.GetCountry()
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboCountry, vval, SET_CONTROL_VALUE
        
        'Update screen elements from the Contact Details table.
        m_frmContactDetails.txtContactDetails(SURNAME).Text = m_clsContactDetailsTable.GetContactSurname()
        m_frmContactDetails.txtContactDetails(FORENAME).Text = m_clsContactDetailsTable.GetContactForname()
        m_frmContactDetails.txtContactDetails(EMAIL_ADDRESS).Text = m_clsContactDetailsTable.GetEmailAddress()
        'EP1033 Add Other Title
        m_frmContactDetails.txtContactDetails(OTHER_TITLE).Text = m_clsContactDetailsTable.GetOtherTitle()
       
        'Set the Contact Title combo selection.
        vval = m_clsContactDetailsTable.GetContactTitle()
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboContactTitle, vval, SET_CONTROL_VALUE
        
        'Set the Contact Type combo selection.
        vval = m_clsContactDetailsTable.GetContactType()
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboContactType, vval, SET_CONTROL_VALUE
                
        'Only populate Telephone controls if there are enough records present.
        'When a contact is saved, we ensure that two records will be saved.
        If TableAccess(m_clsContactDetailsTelephoneTable).RecordCount > 0 Then
            'Ensure we're at the first telephone number.
            TableAccess(m_clsContactDetailsTelephoneTable).MoveFirst
                                
            'Update screen elements from the Contact Details Telephone table.
            m_frmContactDetails.txtContactDetails(COUNTRY_CODE1).Text = m_clsContactDetailsTelephoneTable.GetCOUNTRYCODE()
            m_frmContactDetails.txtContactDetails(AREA_CODE1).Text = m_clsContactDetailsTelephoneTable.GetAREA_CODE()
            m_frmContactDetails.txtContactDetails(TELEPHONE1).Text = m_clsContactDetailsTelephoneTable.GetTELEPHONE()
            m_frmContactDetails.txtContactDetails(EXTENSION1).Text = m_clsContactDetailsTelephoneTable.GetTELEPHONE_EXT()
            
            'Set the Number type combo selection.
            vval = m_clsContactDetailsTelephoneTable.GetType()
            g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboType1, vval, SET_CONTROL_VALUE
        
            If TableAccess(m_clsContactDetailsTelephoneTable).RecordCount > 1 Then
                'Move onto the next telephone number.
                TableAccess(m_clsContactDetailsTelephoneTable).MoveNext
                
                'Update screen elements from the Contact Details Telephone table.
                m_frmContactDetails.txtContactDetails(COUNTRY_CODE2).Text = m_clsContactDetailsTelephoneTable.GetCOUNTRYCODE()
                m_frmContactDetails.txtContactDetails(AREA_CODE2).Text = m_clsContactDetailsTelephoneTable.GetAREA_CODE()
                m_frmContactDetails.txtContactDetails(TELEPHONE2).Text = m_clsContactDetailsTelephoneTable.GetTELEPHONE()
                m_frmContactDetails.txtContactDetails(EXTENSION2).Text = m_clsContactDetailsTelephoneTable.GetTELEPHONE_EXT()
                
                'Set the Number type combo selection.
                vval = m_clsContactDetailsTelephoneTable.GetType()
                g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboType2, vval, SET_CONTROL_VALUE
            End If
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetUKSpecificData
' Description   : Core method to load any data specific to the UK. Client
'                 projects won't necessarily want to call this, so the call is
'                 made from ContactDetailsCS, where the client can tailor it as
'                 needed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetUKSpecificData()
    On Error GoTo Failed
    
    m_frmContactDetails.txtContactDetails(COUNTY).Text = m_clsAddressTable.GetCounty()
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveUKSpecificData
' Description   : Core method to save any data specific to the UK. Client
'                 projects won't necessarily want to call this, so the call is
'                 made from ContactDetailsCS, where the client can tailor it as
'                 needed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveUKSpecificData()
    On Error GoTo Failed
    
    m_clsAddressTable.SetCounty m_frmContactDetails.txtContactDetails(COUNTY).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Called by a form to save the fields associated with Address
'                 and Contact Details. The form MUST have controls of the same
'                 name and same indexes for this generic code to work.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim vTmp As Variant
    Dim strValidation As String     ' EP1033
    Dim strOtherTitle As String     ' EP1033
    
    On Error GoTo Failed

    If m_bContactDetailsExists Then
        
        'Update the Address table from the screen elements.
        m_clsAddressTable.SetPostCode m_frmContactDetails.txtContactDetails(POST_CODE).Text
        m_clsAddressTable.SetBuildingOrHouseName m_frmContactDetails.txtContactDetails(BUILDING_NAME).Text
        m_clsAddressTable.SetBuildingOrHouseNumber m_frmContactDetails.txtContactDetails(BUILDING_NUMBER).Text
        m_clsAddressTable.SetStreet m_frmContactDetails.txtContactDetails(STREET).Text
        m_clsAddressTable.SetDistrict m_frmContactDetails.txtContactDetails(DISTRICT).Text
        m_clsAddressTable.SetTown m_frmContactDetails.txtContactDetails(TOWN).Text
        
        'Read the country combo value into a variant first.
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboCountry, vTmp, GET_CONTROL_VALUE
        m_clsAddressTable.SetCountry vTmp
                                
        'Update Contact Details table from the screen elements.
        m_clsContactDetailsTable.SetContactSurname m_frmContactDetails.txtContactDetails(SURNAME).Text
        m_clsContactDetailsTable.SetContactForname m_frmContactDetails.txtContactDetails(FORENAME).Text
        m_clsContactDetailsTable.SetEmailAddress m_frmContactDetails.txtContactDetails(EMAIL_ADDRESS).Text
        
        'Read the contact title value into a variant first.
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboContactTitle, vTmp, GET_CONTROL_VALUE
        m_clsContactDetailsTable.SetContactTitle vTmp
        
        'EP1033 If the Contact Title has been set to 'Other' save the Other Title value.
        g_clsFormProcessing.GetComboValidation m_frmContactDetails.cboContactTitle, "Title", strValidation
        strOtherTitle = m_frmContactDetails.txtContactDetails(OTHER_TITLE).Text
    
        If (strValidation = "O") Then
            m_clsContactDetailsTable.SetOtherTitle strOtherTitle
        Else
            If (Len(strOtherTitle) > 0) Then
                'Inform the user that the Other Title data will be lost.
                MsgBox ("Contact Title combo is not set to 'Other' so the Other Title will be set to null.")
            End If
            m_clsContactDetailsTable.SetOtherTitle ""
        End If

        'Read the contact type value into a variant first.
        g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboContactType, vTmp, GET_CONTROL_VALUE
        m_clsContactDetailsTable.SetContactType vTmp
        
        'Ensure that two records exist in the telephone number table before we try
        EnsureTelephoneTableIsPopulated
        
        'Update both records in the telephone numbers table, only if they are both there.
        If TableAccess(m_clsContactDetailsTelephoneTable).RecordCount = NUM_TELEPHONE_RECORDS Then
            'Ensure this is the first telephone number.
            TableAccess(m_clsContactDetailsTelephoneTable).MoveFirst
        
            'Update Telephone table from the screen elements.
            m_clsContactDetailsTelephoneTable.SetCOUNTRYCODE m_frmContactDetails.txtContactDetails(COUNTRY_CODE1).Text
            m_clsContactDetailsTelephoneTable.SetAREA_CODE m_frmContactDetails.txtContactDetails(AREA_CODE1).Text
            m_clsContactDetailsTelephoneTable.SetTELEPHONE m_frmContactDetails.txtContactDetails(TELEPHONE1).Text
            m_clsContactDetailsTelephoneTable.SetTELEPHONE_EXT m_frmContactDetails.txtContactDetails(EXTENSION1).Text
            
            'Read the number type value into a variant first.
            g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboType1, vTmp, GET_CONTROL_VALUE
            m_clsContactDetailsTelephoneTable.SetType vTmp
            
            'Move onto the next telephone number.
            TableAccess(m_clsContactDetailsTelephoneTable).MoveNext
        
            'Update Telephone table from the screen elements.
            m_clsContactDetailsTelephoneTable.SetCOUNTRYCODE m_frmContactDetails.txtContactDetails(COUNTRY_CODE2).Text
            m_clsContactDetailsTelephoneTable.SetAREA_CODE m_frmContactDetails.txtContactDetails(AREA_CODE2).Text
            m_clsContactDetailsTelephoneTable.SetTELEPHONE m_frmContactDetails.txtContactDetails(TELEPHONE2).Text
            m_clsContactDetailsTelephoneTable.SetTELEPHONE_EXT m_frmContactDetails.txtContactDetails(EXTENSION2).Text
            
            'Read the number type value into a variant first.
            g_clsFormProcessing.HandleComboExtra m_frmContactDetails.cboType2, vTmp, GET_CONTROL_VALUE
            m_clsContactDetailsTelephoneTable.SetType vTmp
        End If
    End If
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Called to update the Contact Details, Address and Telephone
'                 number tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    
    On Error GoTo Failed
    
    If m_bContactDetailsExists Then
        'Update the database from the data objects.
        TableAccess(m_clsAddressTable).Update
        
        'If there is no information in the contact record AND telephone
        'records, then delete both.
        TableAccess(m_clsContactDetailsTable).Update
            
        'Ensure the Contact's GUID is set against each telephone record. Addresses
        'seem to be handeled externally to this class.
        EnsureTelephoneRecordsHaveGUIDs
            
        'Now update the telephone table from the data object.
        TableAccess(m_clsContactDetailsTelephoneTable).Update
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddressKeyValues
' Description   : Called to set the Address key for the Address record to be
'                 retrieved by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddressKeyValues(colValues As Collection)

    On Error GoTo Failed

    TableAccess(m_clsAddressTable).SetKeyMatchValues colValues
    m_vAddressGUID = colValues(ADDRESS_KEY)
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetContactKeyValues
' Description   : Called to set the ContactDetails key for the ContactDetails
'                 record to be retrieved by this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetContactKeyValues(colValues As Collection)

    On Error GoTo Failed

    TableAccess(m_clsContactDetailsTable).SetKeyMatchValues colValues
    TableAccess(m_clsContactDetailsTelephoneTable).SetKeyMatchValues colValues
    
    m_vContactDetailsGUID = colValues(CONTACT_DETAILS_KEY)
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetContactDetails
' Description   : Returns the Contact details table Class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetContactDetails() As ContactDetailsTable
    
    On Error GoTo Failed
    
    Set GetContactDetails = m_clsContactDetailsTable
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetContactTelephoneDetails
' Description   : Returns the Contact details table Class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetContactTelephoneDetails() As ContactDetailsTelephoneTable
    
    On Error GoTo Failed
    
    Set GetContactTelephoneDetails = m_clsContactDetailsTelephoneTable
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAddressDetails
' Description   : Returns the address details table Class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAddressDetails() As AddressTable
    
    On Error GoTo Failed
    
    Set GetAddressDetails = m_clsAddressTable
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnsureTelephoneTableIsPopulated
' Description   : There must always be n telephone records for each contact. If
'                 not, this routine adds those records and sets their sequence
'                 numbers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureTelephoneTableIsPopulated()
    
    Dim lSequence As Long
                                
    'Keep adding telephone records until we have a full compliment.
    Do While TableAccess(m_clsContactDetailsTelephoneTable).RecordCount < NUM_TELEPHONE_RECORDS
        'Initialise the sequence number to be the next required value.
        lSequence = (TableAccess(m_clsContactDetailsTelephoneTable).RecordCount + 1)

        'Create a record in the telephone table.
        g_clsFormProcessing.CreateNewRecord m_clsContactDetailsTelephoneTable
                
        'Set the sequence number.
        m_clsContactDetailsTelephoneTable.SetTELEPHONE_SEQ_NUMBER lSequence
    Loop

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnsureTelephoneRecordsHaveGUIDs
' Description   : Populates any telephone records with the Contact's GUID if
'                 they aren't already set. This is primarily done when adding a
'                 new contact, when the telephone records have no GUID.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnsureTelephoneRecordsHaveGUIDs()

    Dim i As Long
    Dim vGUID As Variant
    
    'Copy the contact's GUID. This routine is called AFTER the table is updated,
    'so the GUID should always be present.
    vGUID = m_clsContactDetailsTable.GetContactDetailsGUID

    'The contact record may have been removed at this point (if no details
    'were specified). If this isn't the case, then ensure all telephone records
    'have the contact's GUID assigned to them.
    If Not (IsNull(vGUID) Or IsEmpty(vGUID)) Then
        'Ensure we're at the begining of the table.
        TableAccess(m_clsContactDetailsTelephoneTable).MoveFirst
    
        'Move through each telephone record.
        For i = 1 To TableAccess(m_clsContactDetailsTelephoneTable).RecordCount
            'Set the GUID to join the telephone record to the contact record.
            m_clsContactDetailsTelephoneTable.SetCONTACTDETAILSTELEPHONEGUID vGUID
            
            'Move onto the next telephone record.
            TableAccess(m_clsContactDetailsTelephoneTable).MoveNext
        Next i
    End If
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetContactDetailsForm
' Description   : Sets the form to be associated with this class. The form will
'                 contain all controls the class will read/write.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetContactDetailsForm(frmContactDetails As Form)
    Set m_frmContactDetails = frmContactDetails
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetContactDetailsForm
' Description   : Returns the form to be associated with this class. The form
'                 will contain all controls the class will read/write.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetContactDetailsForm() As Form
    Set GetContactDetailsForm = m_frmContactDetails
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns TRUE if this class is in Edit mode mode, FALSE if in
'                 Add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets whether or not this form is in Edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetContactDetailsGUID
' Description   : Returns the ContactDetailsGUID associated with the
'                 ContactDetails record used to store the ContactDetails
'                 details associated with this class (and hence the form)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetContactDetailsGUID() As Variant
    GetContactDetailsGUID = m_clsContactDetailsTable.GetContactDetailsGUID()
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAddressGUID
' Description   : Returns the AddressGUID associated with the Address record
'                 used to store the Address details associated with this class
'                 (and hence the form)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAddressGUID() As Variant
    GetAddressGUID = m_clsAddressTable.GetAddressGUID()
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally to do any per tab initialisation, if this
'                 class is used to support a tabbed dialog.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
End Sub
