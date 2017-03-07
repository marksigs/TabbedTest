VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdAdminFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdAdminFees
' Description   : Provides functionality for the Admin Fees tab on Mortgage Product. Needs to
'                 update the admin fee set on the product, so the Mortgage Product table class
'                 is passed in for this to happen.
' Change history
' Prog      Date        Description
' DJP       22/06/01    SYS2381 SQL Server port, and tidy up
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
' DJP       14/02/02    SYS4095 Reset AdminFeeSet when DeSelected.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const SET_KEY As Integer = 1
Private Const DATE_KEY As Integer = 2

' Private data
Private m_clsAdminFees As AdminFeeTable
Private m_bIsEdit As Boolean
Private m_sAdminFeeSet As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsAdminFees = New AdminFeeTable
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Admin
'                 Fees, including populating the listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    
    m_bIsEdit = bIsEdit
    
    InitialiseFields
    PopulateListView
    
    frmProductDetails.lvAdminFees.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : RePopulates the Admin Fee listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed
    m_clsAdminFees.SetFindNewestAdminFeeSet
    
    TableAccess(m_clsAdminFees).GetTableData POPULATE_FIRST_BAND
    TableAccess(m_clsAdminFees).ValidateData "AdminFees"
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvAdminFees, TableAccess(m_clsAdminFees)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the Mortgage Product Admin Fees tab - called externaly.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
    frmProductDetails.txtAdminFeeSelected.Text = m_sAdminFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all fields on the Admin Fees tab - doesn't actually do anything for this
'                 tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    Dim sFeeSet As String
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKeyMatchValues
' Description   : Returns the keys for the currently selected admin fee.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmProductDetails.lvAdminFees.SelectedItem.Index
    
    Set clsPopulateDetails = frmProductDetails.lvAdminFees.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Admin Fee Keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Admin Fee Keys"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Interest Rates tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    
    ValidateScreenData = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAdminFeeSet
' Description   : Sets the Admin Fee set to be used by this tab when in Edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAdminFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sAdminFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAdminFeeSet
' Description   : Returns the Admin Fee used by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAdminFeeSet()
    GetAdminFeeSet = m_sAdminFeeSet
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DeSelect
' Description   : De-selects the currently selected Admin Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeSelect()
    On Error GoTo Failed
    
    frmProductDetails.txtAdminFeeSelected.Text = ""
    frmProductDetails.cmdAdminFeeDeSelect.Enabled = False
    
    ' We've removed the selection, so set the NULL selection
    m_sAdminFeeSet = ""
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Selected
' Description   : Selects the currently highlighted Admin Fee.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Selected()
    On Error GoTo Failed
    Dim colValues As Collection
    
    Set colValues = New Collection
    
    GetKeyMatchValues colValues

    If colValues.Count >= SET_KEY Then
        m_sAdminFeeSet = colValues(SET_KEY)
        frmProductDetails.txtAdminFeeSelected.Text = m_sAdminFeeSet
        frmProductDetails.cmdAdminFeeDeSelect.Enabled = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProductKeys
' Description   : Sets the product keys to be used by this tab when retrieving data, if in
'                 edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetButtonState
' Description   : Sets the state of the Select/Deselect buttons on the Admin Fees tab depending
'                 on the state of the Admin Fees listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetButtonState(lvListView As MSGListView, cmdSelect As CommandButton, cmdDeselect As CommandButton)
    On Error GoTo Failed
    Dim nRecordCount As Long
    Dim bEnable As Boolean
    Dim lstItem As ListItem
    
    nRecordCount = lvListView.ListItems.Count
            
    bEnable = False
    
    Set lstItem = lvListView.SelectedItem
    
    If Not lstItem Is Nothing Then
        bEnable = True
    End If
    
    cmdSelect.Enabled = bEnable
    cmdDeselect.Enabled = bEnable

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the state of fields and buttons on the Admin Fees tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields()
On Error GoTo Failed

    frmProductDetails.cmdAdminFeeSelect.Enabled = False
    frmProductDetails.cmdValuationFeeSelect.Enabled = False

Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "InitialiseFields:" + Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called whenever this tab is switched to, so perform any initialisation required
'                 by the tab when it gets focus.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

