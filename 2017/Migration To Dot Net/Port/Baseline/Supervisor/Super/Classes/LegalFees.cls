VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LegalFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LegalFees
' Description   : Support class for the Lender form. Contains code to save/load and validate all
'                 data for this tab.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_vSetKey  As Variant
Private m_clsLegalFees As LegalFeesTable
Private m_clsLegalFeeSet As LegalFeeSetTable
Private m_vOrganisationID As Variant
Private Sub Class_Initialize()
    ' Create table classes
    Set m_clsLegalFeeSet = New LegalFeeSetTable
    Set m_clsLegalFees = New LegalFeesTable
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to initialise the Legal Fees tab. Will load all data required by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    m_bIsEdit = bIsEdit
    SetupListViewHeaders
    
    SetRateSet
    
    PopulateListView
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the listview with the data returned from the database via the legal
'                 fees table class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsLegalFees
    
    ' It is a banded get
    clsTableAccess.GetTableData POPULATE_FIRST_BAND
        
    ' Populate the listview
    g_clsFormProcessing.PopulateFromRecordset frmLenderDetails.lvLegalFees, clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Sets the listview headers to be displayed to the user.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders()
    On Error GoTo Failed
    
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    ' Width, name
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Set No."
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Fee Type"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Type of Application"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Max Value "
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Amount "
    headers.Add lvHeaders
    
    frmLenderDetails.lvLegalFees.AddHeadings headers

    frmLenderDetails.lvLegalFees.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on this tab, but there aren't any as it is a listview.
'                 This method is called externally so still has to be supported even though it
'                 doesn't do anything.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all data on this tab to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    
    m_clsLegalFeeSet.SetOrganisationID m_vOrganisationID
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Add
' Description   : Called when the user presses the Add button. Needs to load the Legal Fee Edit
'                 screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess

    m_bIsEdit = False

    frmLegalFeeSets.SetIsEdit False
    
    If Len(m_vSetKey) > 0 Then
        ' Set the fee set number on the fee sets GUI, too.
        frmLegalFeeSets.SetSetNumber CStr(m_vSetKey)
        frmLegalFeeSets.Show vbModal, frmLenderDetails
    
        HandleReturn
        
        InitTab
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "LegalFees:Add - Set key is empty"
    End If
    
    SetButtonsState
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub HandleReturn()
    On Error GoTo Failed
    
    If frmLegalFeeSets.GetReturnCode = MSGSuccess Then
        PopulateListView
    End If
    
    Unload frmLegalFeeSets
    frmLenderDetails.lvLegalFees.SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetLegalFeeKeys
' Description   : Obtains the keys from the current row in the listview to identifiy that row. The
'                 PopulateDetails class will contain the keys required, as stored in the listview
'                 itself.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetLegalFeeKeys(colKeyValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmLenderDetails.lvLegalFees.SelectedItem.Index
    Set clsPopulateDetails = frmLenderDetails.lvLegalFees.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colKeyValues = clsPopulateDetails.GetKeyMatchValues
    End If
    
    If colKeyValues Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "LegalFees, unable to obtain Keys"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Edit
' Description   : Edits the legal fee at the current selected row.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Edit()
    On Error GoTo Failed
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess

    m_bIsEdit = True
    Set clsTableAccess = m_clsLegalFees
    
    GetLegalFeeKeys colMatchValues
    
    If Len(m_vSetKey) > 0 Then
        ' Set the fee set number on the fee sets GUI, too.
        frmLegalFeeSets.SetKeyMatchValues colMatchValues
        frmLegalFeeSets.SetIsEdit
        frmLegalFeeSets.SetSetNumber CStr(m_vSetKey)
        frmLegalFeeSets.Show vbModal, frmLenderDetails
        
        HandleReturn
        
        InitTab
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Delete
' Description   : Deletes the legal fee at the current selected row.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
    On Error GoTo Failed
    Dim colValues As Collection
    
    If Not frmLenderDetails.lvLegalFees.SelectedItem Is Nothing Then
        If MsgBox("Delete selected Legal Fee?", vbYesNo + vbQuestion) = vbYes Then
            GetLegalFeeKeys colValues
            
            TableAccess(m_clsLegalFees).SetKeyMatchValues colValues
            TableAccess(m_clsLegalFees).DeleteRecords
                
            PopulateListView
            SetButtonsState
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetRateSet
' Description   : Sets the FeeSet key for the current FeeSet. This will either be obtained from the
'                 LegalFeeSet table, or if it doesn't exist, will be created.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetRateSet()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    Set clsTableAccess = m_clsLegalFeeSet
    
    colMatchValues.Add m_vOrganisationID
    
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.GetTableData
    clsTableAccess.ValidateData

    If clsTableAccess.RecordCount > 0 Then
        clsTableAccess.MoveFirst

        m_vSetKey = m_clsLegalFeeSet.GetFeeSet()
    Else
        ' No Legal fee records
        Create
    End If

    m_clsLegalFees.SetFeeSetKey m_vSetKey
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Create
' Description   : Creates a LegalFeeSet record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Create()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    
    ' Now the MIG Rate Set
    m_vSetKey = m_clsLegalFeeSet.NextFeeSet
    
    Set colMatchValues = New Collection
    colMatchValues.Add m_vSetKey
    TableAccess(m_clsLegalFeeSet).SetKeyMatchValues colMatchValues

    ' Create the entry on the MIGRateSet table
    g_clsFormProcessing.CreateNewRecord m_clsLegalFeeSet
    m_clsLegalFeeSet.SetFeeSet m_vSetKey
    m_clsLegalFees.SetFeeSetKey m_vSetKey

    TableAccess(m_clsLegalFeeSet).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Commits all data on this tab to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed

    TableAccess(m_clsLegalFeeSet).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally to this class and performs any tab initialisation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Set frmLenderDetails.lvLegalFees.SelectedItem = Nothing
    SetButtonsState
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetButtonsState
' Description   : Sets the state of the Add/Edit/Delete buttons based on the current state of
'                 the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetButtonsState()
    On Error GoTo Failed
    Dim objSel As Object
    Dim bEnable As Boolean
    
    bEnable = False
    Set objSel = frmLenderDetails.lvLegalFees.SelectedItem
    
    If Not objSel Is Nothing Then
        bEnable = True
    End If

    frmLenderDetails.cmdEditLegalFees.Enabled = bEnable
    frmLenderDetails.cmdDeleteLegalFees.Enabled = bEnable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called externally to this class and performs any tab validation, if required
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets the edit state of this tab - edit or add.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsEdit
' Description   : Returns the edit state of this tab - edit or add.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsEdit()
    IsEdit = m_bIsEdit
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetOrganisationID
' Description   : Given the Organisation ID passed in, save it for future use.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetOrganisationID(vOrgID As Variant)
    On Error GoTo Failed
    
    If Not IsEmpty(vOrgID) Then
        If Len(vOrgID) > 0 Then
            m_vOrganisationID = vOrgID
        Else
            g_clsErrorHandling.RaiseError errOrganisationIDMissing
        End If
    Else
        g_clsErrorHandling.RaiseError errOrganisationIDMissing
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "LegalFees.SetOrganisationID: " & Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsCreated
' Description   : Returs TRUE if a FeeSet key has been created, FALSE if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsCreated() As Boolean
    Dim bIsCreated As Boolean
    On Error GoTo Failed
    
    bIsCreated = False
    
    If Not IsEmpty(m_vSetKey) Then
        If Len(m_vSetKey) > 0 Then
            bIsCreated = True
        End If
    End If
    
    IsCreated = bIsCreated
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
