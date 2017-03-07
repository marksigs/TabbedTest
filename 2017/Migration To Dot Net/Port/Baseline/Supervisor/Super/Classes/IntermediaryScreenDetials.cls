VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntermediaryScreenDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : IntermediaryScreenDetails
' Description  : Main intermediary screen functionality. Unlike most other tab-
'                handlers, this one does not control the lifetime and loading
'                of data into/from it's primary table object. This is because
'                the Report/Correspondance popup form also needs access to that
'                table and the main form therefore controls the table object's
'                scope.
'
'                An Intermediary is store in two tables, the base table and
'                then either an Organisation table or an Individual table
'                (based upon their type).
'
' Change history
' Prog      Date        Description
' AA        22/06/2001  Added class
' DJP       27/06/01    SQL Server port
' STB       14/11/01    SYS2972: ReportFrequency combo looking for wrong combo
'                       group.
' STB       10/12/01    SYS2550: SQL-Server integration.
' AW        08/01/02    SYS3560: Added 'Optimus cust No' to SetScreenFields()
' STB       20/02/02    SYS3767: ActiveFrom must always be prior to ActiveTo.
' STB       20/02/02    SYS3768: LTV Limit label also enabled/disabled.
' STB       25/03/02    SYS4312: Added call to export data to Optimus.
' STB       22/04/02    SYS4397: Proc fees can be set against all Intermediary
'                       types.
' STB       14/05/02    SYS4558: ThirdParty type hardcoded to 71 for ODI
'                       updates.
' STB       22/05/02    SYS4558: Send the Country combo ID to ODI.
' STB       27-03-2002  MSMS0011 Added the InteremdiaryStatus combo for individuals.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' BMIDS
' AW        11/07/02    BMIDS00177  Removed UpdateODITata()
'


Option Explicit

'Taken from the contact details class. These should be made public...
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

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'Local reference to the main Edit form.
Private m_frmInt As frmEditIntermediaries

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'The type of intermediary (Individual, Company, etc.)
Private m_uIntermediaryType As IntermediaryTypeEnum

'Underlying table object for the BASE record.
Private m_clsIntermediaryTable As IntermediaryTable

'A generic table interface for the more specific Intermediary data. It will
'be cast onto an Individual table object or an Organisation table object when
'required, based upon the Intermediary Type.
Private m_clsIntermediaryTypeTable As TableAccess

'Store the PanelID to detect if it has changed.
Private m_sPanelID As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Set the edit state Add/Edit.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate all combos and other 'static' lists.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateScreenControls()
                    
    Dim clsIntermediary As Intermediary
    
    On Error GoTo Failed
    
    'Create an Intermediary helper-class.
    Set clsIntermediary = New Intermediary
    
    'Populate the Intermediary Type combo.
    g_clsFormProcessing.PopulateCombo "IntermediaryType", m_frmInt.cboType, False
    
    'Populate the Title combo.
    g_clsFormProcessing.PopulateCombo "Title", m_frmInt.cboTitle, False
    
    'Populate the Key Development Area combo.
    g_clsFormProcessing.PopulateCombo "KeyDevelopmentArea", m_frmInt.cboKeyDevArea, False
    
    'Populate the Procuration Fee Payment Method combo.
    g_clsFormProcessing.PopulateCombo "ProcurationFeePaymentMethod", m_frmInt.cboProcFeeMethod
    
    'Populate the Report Frequency combo.
    g_clsFormProcessing.PopulateCombo "IntermediaryReportFrequency", m_frmInt.cboReportFreq
    
    'Populate the Communication Method combo.
    g_clsFormProcessing.PopulateCombo "CommunicationMethod", m_frmInt.cboCommunicationMethod
    
    'Populate the available parents.
    clsIntermediary.PopulateParentIntCombo m_frmInt.cboParentIntermediary, m_uIntermediaryType
    
    'Populate the Intermediary Status combo.
    g_clsFormProcessing.PopulateCombo "IntermediaryStatus", m_frmInt.cboIntermediaryStatus, True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup the screen controls to work in an Add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup the screen controls to work in an Edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populate the screen controls from the underlying table object(s).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
            
    Dim vVal As Variant
    Dim vParentGUID As Variant
    Dim clsAllIntermediaries As IntermediaryTable
            
    On Error GoTo Failed
    
    'Create a tempory Intermediaries table object to hold all records.
    Set clsAllIntermediaries = New IntermediaryTable
        
    'Active From.
    m_frmInt.txtActiveFrom.Text = m_clsIntermediaryTable.GetActiveFrom
    
    'Active To.
    m_frmInt.txtActiveTo.Text = m_clsIntermediaryTable.GetActiveTo
    
    'Panel ID.
    m_frmInt.txtPanelID.Text = CStr(m_clsIntermediaryTable.GetPanelID)
    
    'Optimus Customer number
    m_frmInt.txtOtherSystemNo.Text = m_clsIntermediaryTable.GetOtherSystemNumber
    
    'Store the panel ID. If we're editing, we need to make sure that (only if
    'it's changed) that it isn't a duplicate.
    m_sPanelID = m_frmInt.txtPanelID.Text
    
    'Communication Method.
    vVal = m_clsIntermediaryTable.GetCommunicationMethod
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboCommunicationMethod, vVal, SET_CONTROL_VALUE
    
    'Procuration fee payment method.
    vVal = m_clsIntermediaryTable.GetProcFeePaymentMethod
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboProcFeeMethod, vVal, SET_CONTROL_VALUE
    
    'Procuration fee ltv limit.
    m_frmInt.txtLTVLimit.Text = m_clsIntermediaryTable.GetProcurationFeeLTVLimit
        
    'Intermediary Type.
    vVal = m_clsIntermediaryTable.GetIntermediaryType
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboType, vVal, SET_CONTROL_VALUE
    
    'Populate the controls specific to the intermediary type.
    If m_uIntermediaryType = IndividualType Then
        LoadIndividual
    Else
        LoadOrganisation
    End If
        
    'This is an edit, so the panel id of the parent must be retrieved
    TableAccess(clsAllIntermediaries).SetRecordSet m_frmInt.cboParentIntermediary.RowSource
        
    'Get the parent intermediaries GUID.
    vParentGUID = m_clsIntermediaryTable.GetParentIntermediaryGUID
    
    If Not IsNull(vParentGUID) Then
        If Len(CStr(vParentGUID)) > 0 Then
            vParentGUID = g_clsSQLAssistSP.GuidToString(CStr(vParentGUID))
            
            'Select the parent's PanelID.
            m_frmInt.cboParentIntermediary.SelText = clsAllIntermediaries.GetPanelIDForGUID(CStr(vParentGUID))
        End If
    End If
    
    'The caption of the interemdiary screen should reflect the type.
    m_frmInt.Caption = "Add/Edit " & m_frmInt.cboType.Text
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Create underlying objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableFrame
' Description   : Enables/disables ctrls within a frame.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableFrame(ByRef fra As Frame, Optional ByVal bEnabled As Boolean = False)
        
    Dim ctrl As Control
    
    On Error GoTo Failed
    
    'Set the state for each control within the frame.
    For Each ctrl In m_frmInt
        'The state change affects any controls within the frame as well as the frame itself.
        If ctrl.Container.Name = fra.Name Or ctrl.Name = fra.Name Then
            ctrl.Enabled = bEnabled
        End If
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetForm
' Description   : Store a reference to the form itself.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetForm(ByRef pfrmEditIntermediaries As Form)
    Set m_frmInt = pfrmEditIntermediaries
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Routine to ensure table objects are created and populated
'                 accoding to the underlying add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal bIsEdit As Boolean)
    
    On Error GoTo Failed
    
    'Store the add/edit state at module level.
    m_bIsEdit = bIsEdit
    
    'Create the type-specific table appropriate to the intermediary-type.
    CreateAdditionalTable
    
    'Setup the screen controls properties which are not governed by the add/edit mode.
    SetCommonState

    'Set state-specific stuff.
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
' Function      : SetIntermediaryType
' Description   : sets the type index for the type combo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIntermediaryType(ByVal uIntermediaryType As IntermediaryTypeEnum)
    m_uIntermediaryType = uIntermediaryType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIntermediaryTable
' Description   : Associate the base, intermediary table with the tab-handler.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIntermediaryTable(ByRef clsIntermediaryTable As IntermediaryTable)
    Set m_clsIntermediaryTable = clsIntermediaryTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Copies the screen control data (handled by this class) into the underlying
'                 table object(s).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    
    Dim vTmp As Variant
    
    On Error GoTo Failed
    
    'Active from
    g_clsFormProcessing.HandleDate m_frmInt.txtActiveFrom, vTmp, GET_CONTROL_VALUE
    m_clsIntermediaryTable.SetActiveFrom CDate(vTmp)
    
    'Active To
    g_clsFormProcessing.HandleDate m_frmInt.txtActiveTo, vTmp, GET_CONTROL_VALUE
    m_clsIntermediaryTable.SetActiveTo vTmp
    
    'Interemediary Type
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboType, vTmp, GET_CONTROL_VALUE
    m_clsIntermediaryTable.SetIntermediaryType CStr(vTmp)
    
    'Panel ID
    m_clsIntermediaryTable.SetPanelID m_frmInt.txtPanelID.Text
    
    'Store the PanelID at module level. Used to detect change.
    m_sPanelID = m_frmInt.txtPanelID.Text
        
    'Parent GUID.
    vTmp = m_frmInt.cboParentIntermediary.BoundText
    
    If Len(vTmp) > 0 Then
        'Store the parent GUID in the table.
        m_clsIntermediaryTable.SetParentGUID g_clsSQLAssistSP.GuidToString(CStr(vTmp))
    End If
    
    'Communication Method
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboCommunicationMethod, vTmp, GET_CONTROL_VALUE
    m_clsIntermediaryTable.SetCommunicationMethod vTmp
    
    'Procuration Fee Limit
    m_clsIntermediaryTable.SetProcurationFee m_frmInt.txtLTVLimit.Text
    
    'Procuration Fee Payment Method
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboProcFeeMethod, vTmp, GET_CONTROL_VALUE
    m_clsIntermediaryTable.SetProcurationFeeMethod vTmp
        
    'Need to save the intermediary type details
    If m_uIntermediaryType = IndividualType Then
        SaveIndividual
    Else
        SaveOrganisation
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Saves the field data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()

    On Error GoTo Failed

    'Update the underlying table.
    TableAccess(m_clsIntermediaryTable).Update
    
    'If that worked, update the specific table.
    m_clsIntermediaryTypeTable.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveIndividual
' Description   : Saves all screen data relating to the individual intermediary. The generic
'                 information is saved in SaveScreenData.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveIndividual()
      
   Dim vVal As Variant
   Dim clsIndividualTable As IntermediaryIndividual
   
   On Error GoTo Failed
    
   'Cast a specific interface onto the generic object.
   Set clsIndividualTable = m_clsIntermediaryTypeTable
    
   'Forename.
   clsIndividualTable.SetForeName m_frmInt.txtForename.Text
       
    'Surname.
   clsIndividualTable.SetSurName m_frmInt.txtSurname.Text
   
   'Title.
   g_clsFormProcessing.HandleComboExtra m_frmInt.cboTitle, vVal, GET_CONTROL_VALUE
   clsIndividualTable.SetTitle vVal
   
   'Dev Area.
   g_clsFormProcessing.HandleComboExtra m_frmInt.cboKeyDevArea, vVal, GET_CONTROL_VALUE
   clsIndividualTable.SetKeyDevArea vVal
   
   'IntermediaryStatus.
   g_clsFormProcessing.HandleComboExtra m_frmInt.cboIntermediaryStatus, vVal, GET_CONTROL_VALUE
   clsIndividualTable.SetIntermediaryStatus vVal
   
   Exit Sub
   
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LoadIndividual
' Description   : Populates screen controls with data from the underlying table which is
'                 specific to the individual intermediary type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadIndividual()

    Dim vVal As Variant
    Dim clsIndividual As IntermediaryIndividual

    'Cast an interface of the appropriate type on the type-specific table object.
    Set clsIndividual = m_clsIntermediaryTypeTable
    
    'Forename and Surname.
    m_frmInt.txtForename.Text = clsIndividual.GetIndividualForeName
    m_frmInt.txtSurname.Text = clsIndividual.GetIndividualSurName
    
    'Title.
    vVal = clsIndividual.GetIndividualTitle
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboTitle, vVal, SET_CONTROL_VALUE
    
    'Key development area combo.
    vVal = clsIndividual.GetIndividualDevArea
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboKeyDevArea, vVal, SET_CONTROL_VALUE
        
    'IntermediaryStatus combo.
    vVal = clsIndividual.GetIntermediaryStatus
    g_clsFormProcessing.HandleComboExtra m_frmInt.cboIntermediaryStatus, vVal, SET_CONTROL_VALUE
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveOrganisation
' Description   : Saves all screen date relating to the organisation intermediary. The generic
'                 information is saved in SaveScreenData.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveOrganisation()
            
    Dim clsOrganisationTable As IntermediaryOrganisation
   
    On Error GoTo Failed
    
    'Cast a specific interface onto the generic object.
    Set clsOrganisationTable = m_clsIntermediaryTypeTable
   
    'Name
    clsOrganisationTable.SetCompanyName m_frmInt.txtOrgName.Text
    
    'Commission number
    clsOrganisationTable.SetCommisionNumber m_frmInt.txtCommisionNo.Text
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LoadOrganisation
' Description   : Populates screen controls with data from the underlying table which is
'                 specific to the individual intermediary type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOrganisation()

    Dim clsOrganisation As IntermediaryOrganisation

    'Cast an interface of the appropriate type on the type-specific table object.
    Set clsOrganisation = m_clsIntermediaryTypeTable
    
    'Organisation Name.
    m_frmInt.txtOrgName.Text = clsOrganisation.GetCompanyName
    
    'Commission Number.
    m_frmInt.txtCommisionNo.Text = clsOrganisation.GetCommisionNumber
        
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates the data on the screen. Duplicates etc
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    
    Dim sPanelID As String
    Dim bSuccess As Boolean
    Dim colValues As Collection
    Dim clsIntermediary As Intermediary
    
    On Error GoTo Failed

    'Create a duplicate-key value collection.
    Set colValues = New Collection
    
    'Store the panel ID, the code will be more readable.
    sPanelID = m_frmInt.txtPanelID.Text
    
    'Add the PanelID into the collection.
    colValues.Add sPanelID
    
    'Ensure that if we're adding, or if we're editing (and the PanelID is
    'altered) that it is not duplicated.
    If (Not m_bIsEdit) Or (m_sPanelID <> sPanelID) Then
        bSuccess = Not TableAccess(m_clsIntermediaryTable).DoesRecordExist(colValues, TableAccess(m_clsIntermediaryTable).GetDuplicateKeys)
    Else
        bSuccess = True
    End If
    
    'Report the problem if the PanelID is a duplicate.
    If Not bSuccess Then
        g_clsErrorHandling.DisplayError "Panel ID must be unique"
        m_frmInt.txtPanelID.SetFocus
    End If
    
    'Create a helper class to use the data validation routine.
    Set clsIntermediary = New Intermediary
    
    'Verify the from and to dates.
    clsIntermediary.ValidateDate m_frmInt.txtActiveFrom, m_frmInt.txtActiveTo
    
    'Indicate to the caller if this tab is valid or not.
    ValidateScreenData = bSuccess
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Terminate
' Description   : Releases all module-level object references. This can be used to avoid a
'                 cyclic-reference problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Terminate()

    Set m_frmInt = Nothing
    Set m_colKeys = Nothing
    Set m_clsIntermediaryTable = Nothing
    Set m_clsIntermediaryTypeTable = Nothing

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Terminate
' Description   : Call into the shared routine used to release object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Terminate
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetCommonState
' Description   : Setup screen controls which don't depend upon add/edit state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCommonState()

    'Adjust mandatory fields and specific input controls state.
    If m_uIntermediaryType = IndividualType Then
        'Enable / disable the frames accordingly.
        EnableFrame m_frmInt.fraOrganisation, False
        EnableFrame m_frmInt.fraIndividual, True
        
        'Set the mandatory fields accoringly.
        m_frmInt.txtForename.Mandatory = True
        m_frmInt.txtSurname.Mandatory = True
        m_frmInt.cboTitle.Mandatory = True
        m_frmInt.txtOrgName.Mandatory = False
        m_frmInt.cboIntermediaryStatus.Mandatory = True
        
    'If this is an organisation type.
    Else
        'Enable / disable the frames accordingly.
        EnableFrame m_frmInt.fraOrganisation, True
        EnableFrame m_frmInt.fraIndividual, False
        
        'Set the mandatory fields accoringly.
        m_frmInt.txtForename.Mandatory = False
        m_frmInt.txtSurname.Mandatory = False
        m_frmInt.cboTitle.Mandatory = False
        m_frmInt.txtOrgName.Mandatory = True
        m_frmInt.cboIntermediaryStatus.Mandatory = False
    End If
    
    'Enable or disable command buttons and other controls.
    
    'LTV Fee Limit - if Lead Agent.
    m_frmInt.txtLTVLimit.Enabled = (m_uIntermediaryType = LeadAgentType)
    m_frmInt.lblLTVLimit.Enabled = (m_uIntermediaryType = LeadAgentType)
    
    'SYS4397 - Procuration Fees can be set on all types.
    'Procuration Fee button - if Lead Agent.
    'm_frmInt.cmdProcFee.Enabled = (m_uIntermediaryType = LeadAgentType)
    'SYS4397 - End.
    
    'Cross Selling button - if Individual.
    m_frmInt.cmdCrossSelling.Enabled = (m_uIntermediaryType = IndividualType)
    
    'Target button - if Individual.
    m_frmInt.cmdTarget.Enabled = (m_uIntermediaryType = IndividualType)
    
    'Correspondence button - when NOT Lead Agent.
    m_frmInt.cmdCorrespondence.Enabled = (m_uIntermediaryType <> LeadAgentType)
    
    'Always disable the intermediary type combo.
    m_frmInt.cboType.Enabled = False

    'Always disable the parent intermediary combo.
    m_frmInt.cboParentIntermediary.Enabled = False
    
    'Only Individuals have an IntermediaryStatus.
    m_frmInt.cboIntermediaryStatus.Enabled = (m_uIntermediaryType = IndividualType)
    
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateAdditionalTable
' Description   : Creates either an Individual or an Organisation table object and
'                 loads or creates a record relevant to the current Intermediary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateAdditionalTable()

    Dim clsIndividualTable As IntermediaryIndividual
    Dim clsOrganisationTable As IntermediaryOrganisation
    
    If m_uIntermediaryType = IndividualType Then
        'Create the table object.
        Set clsIndividualTable = New IntermediaryIndividual
        
        'Cast a more generic interface onto it.
        Set m_clsIntermediaryTypeTable = clsIndividualTable
    Else
        'Create the table object.
        Set clsOrganisationTable = New IntermediaryOrganisation
        
        'Cast a more generic interface onto it.
        Set m_clsIntermediaryTypeTable = clsOrganisationTable
    End If
    
    'If we're editing just get some data.
    If m_bIsEdit Then
        'Set the keys collection.
        m_clsIntermediaryTypeTable.SetKeyMatchValues m_colKeys
    
        'Load the additional data into the type-specific table.
        m_clsIntermediaryTypeTable.GetTableData
    Else
        'If we're adding, we must create a record AND set the GUID.
        g_clsFormProcessing.CreateNewRecord m_clsIntermediaryTypeTable
                
        'Set the IntermediaryGUID.
        If m_uIntermediaryType = IndividualType Then
            clsIndividualTable.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
        Else
            clsOrganisationTable.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
        End If
    End If
    
End Sub

