VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIGRateSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MIGRateSets
' Description   : Manages all form processing for the MIG Rate Sets tab on the Lender Details
'                 form.
' Change history
' Prog      Date        Description
' DJP       16/06/01    SQL Server port - use variants for guid keys
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' JD        30/03/2005  BMIDS982 Changed screen text from MIG to HLC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const ORGANISATION_ID_FIELD = 1

' Private data
Private m_vSetKey  As Variant
Private m_clsMIGRates As MIGRateTable
Private m_clsMIGRateSet As MIGRateSetsTable
Private Sub Class_Initialize()
    ' Create table classes
    Set m_clsMIGRates = New MIGRateTable
    Set m_clsMIGRateSet = New MIGRateSetsTable
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to initialise the Legal Fees tab. Will load all data required by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    
    ' Setup the listview
    SetupListViewHeaders
    
    ' The RateSet is the key for MIGRates - make sure we have one
    SetRateSet
    
    ' And finally populate the listview any records we may have.
    PopulateListView
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetRateSet
' Description   : The RateSet is the key for MIGRates. This method ensures there is a rateset
'                 whether we are adding or editing. If a rateset does not exist, create one.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetRateSet()
    On Error GoTo Failed
    Dim bCreated As Boolean
    
    bCreated = IsCreated
    
    If Not bCreated Then
        CreateRateSet
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CreateRateSet
' Description   : Creates the rateset in the MIGRateSet table. This has to exist in order to create
'                 and update records in the frmEditMIGRates form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateRateSet()
    On Error GoTo Failed
    Dim clsMIGRateSet As MIGRateSetsTable
    
    Set clsMIGRateSet = New MIGRateSetsTable
    
    ' Now the MIG Rate Set
    m_vSetKey = clsMIGRateSet.NextRateSet

    ' Create the entry on the MIGRateSet table
    g_clsFormProcessing.CreateNewRecord clsMIGRateSet
    clsMIGRateSet.SetRateSet m_vSetKey
    
    TableAccess(clsMIGRateSet).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the listview with the data returned from the database via the MIG
'                 Rates table class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed
    
    m_clsMIGRates.GetMIGRateSets m_vSetKey
    g_clsFormProcessing.PopulateFromRecordset frmLenderDetails.lvMIGRateSets, m_clsMIGRates
    
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
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "HLC Rate"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Maximum LTV"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Max Loan Component"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Amount"
    headers.Add lvHeaders
    
    frmLenderDetails.lvMIGRateSets.AddHeadings headers
    frmLenderDetails.lvMIGRateSets.LoadColumnDetails TypeName(Me)
    
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
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetMIGRateKeys
' Description   : Obtains the keys from the current row in the listview to identifiy that row. The
'                 PopulateDetails class will contain the keys required, as stored in the listview
'                 itself.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetMIGRateKeys(colKeyValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmLenderDetails.lvMIGRateSets.SelectedItem.Index
    Set clsPopulateDetails = frmLenderDetails.lvMIGRateSets.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colKeyValues = clsPopulateDetails.GetKeyMatchValues
    End If
    
    If colKeyValues Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "HLC Rates, unable to obtain Keys"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Edit
' Description   : Edits the MIG Rate set at the current selected row.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Edit()
    Dim colMatchValues As New Collection

    On Error GoTo Failed
    
    ' Get the keys for the highlighted row
    GetMIGRateKeys colMatchValues
    
    frmEditMIGRates.SetKeyMatchValues colMatchValues
    frmEditMIGRates.Show vbModal, frmLenderDetails
    
    HandleReturn
    
    InitTab
    
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

    If Len(m_vSetKey) > 0 Then
        frmEditMIGRates.SetRateSet m_vSetKey
        frmEditMIGRates.Show vbModal, frmLenderDetails
        
        HandleReturn
                
        InitTab
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "MIGRateSets:Add - Set key is empty"
    End If
    
    SetButtonsState
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : HandleReturn
' Description   : Handles the return code from the MIG Rates form. If the return code was success,
'                 the ListView has to be repopulated (as something was changed).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HandleReturn()
    On Error GoTo Failed
    
    If frmEditMIGRates.GetReturnCode() = MSGSuccess Then
        PopulateListView
    End If
    
    Unload frmEditMIGRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Delete
' Description   : Deletes the MIG Rate set at the current selected row.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
    On Error GoTo Failed
    Dim colValues As Collection
    
    If Not frmLenderDetails.lvMIGRateSets.SelectedItem Is Nothing Then
        If MsgBox("Delete selected HLC Rate Set?", vbYesNo + vbQuestion) = vbYes Then
            GetMIGRateKeys colValues
            TableAccess(m_clsMIGRates).SetKeyMatchValues colValues
            TableAccess(m_clsMIGRates).DeleteRecords
                
            PopulateListView
            SetButtonsState
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Called externally - does nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed

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
    
    Set frmLenderDetails.lvMIGRateSets.SelectedItem = Nothing
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
    Dim objSel As Object
    Dim bEnable As Boolean
    
    bEnable = False
    Set objSel = frmLenderDetails.lvMIGRateSets.SelectedItem
    
    If Not objSel Is Nothing Then
        bEnable = True
    End If

    frmLenderDetails.cmdMIGEdit.Enabled = bEnable
    frmLenderDetails.cmdMIGDelete.Enabled = bEnable

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called externally to this class and performs any tab validation, if required
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetMIGRateSet
' Description   : Sets the key values to identify the list of MIG Rate Sets to be displayed on this
'                 tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetMIGRateSet(vRateSet As Variant)
    On Error GoTo Failed
    
    m_vSetKey = vRateSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetRateSet
' Description   : Returns the RateSet used for the curret MIGRate set.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRateSet() As String
    On Error GoTo Failed
    
    If Not IsEmpty(m_vSetKey) Then
        If Len(m_vSetKey) > 0 Then
            GetRateSet = m_vSetKey
        Else
            g_clsErrorHandling.RaiseError errRateSetEmpty, "MIGRateSets.GetRateSet"
        End If
    Else
        g_clsErrorHandling.RaiseError errRateSetEmpty, "MIGRateSets.GetRateSet"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsCreated
' Description   : Returns true if a MIGRate set exists, false if not.
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
