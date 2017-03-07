VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdTranEquFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdTranEquFee
' Description   : Support class for the Transfer Of Equity Fee tab on the Mortgage Product screen
' Change history
' Prog      Date        Description
' TW        11/12/2006  EP2_20 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const SET_KEY As Integer = 1
Private Const DATE_KEY As Integer = 2

' Private data
Private m_bIsEdit As Boolean
Private m_colMatchValues As Collection
Private m_sTransferOfEquityFeeSet As String
Private m_clsTransferOfEquityFees As TransferOfEquityFeeTable

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsTransferOfEquityFees = New TransferOfEquityFeeTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Additional Borrowing
'                 Fees, including populating the listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    m_bIsEdit = bIsEdit

    InitialiseFields
    PopulateListView
    
    If m_bIsEdit Then
        SetScreenFields
    End If
    
    frmProductDetails.lvTransferOfEquityFees.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : RePopulates the Transfer of Equity Fee listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed
    
    m_clsTransferOfEquityFees.SetFindNewestTransferOfEquityFeeSet
    
    TableAccess(m_clsTransferOfEquityFees).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvTransferOfEquityFees, m_clsTransferOfEquityFees
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on the Mortgage Product ValuationFee tab - called externaly.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
    
    frmProductDetails.txtTransferOfEquityFeeSelected.Text = m_sTransferOfEquityFeeSet

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all fields on the Transfer of Equity Fees tab - doesn't actually do anything for this
'                 tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
    On Error GoTo Failed
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DeSelect
' Description   : De-selects the currently selected Transfer of Equity Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeSelect()
    On Error GoTo Failed
    
    frmProductDetails.txtTransferOfEquityFeeSelected.Text = ""
    frmProductDetails.cmdTransferOfEquityFeeDeSelect.Enabled = False

    ' Reset the Transfer of Equity Fee Set
    m_sTransferOfEquityFeeSet = ""
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Selected
' Description   : Selects the currently highlighted Transfer of Equity Fee.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Selected()
    On Error GoTo Failed
    Dim sSet As String
    Dim colValues As New Collection
    
    GetKeyMatchValues colValues
    
    If colValues.Count >= SET_KEY Then
        m_sTransferOfEquityFeeSet = colValues(SET_KEY)
        frmProductDetails.txtTransferOfEquityFeeSelected.Text = colValues(SET_KEY)
        frmProductDetails.cmdTransferOfEquityFeeDeSelect.Enabled = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Transfer of Equity Fee tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeSet
' Description   : Sets the Transfer of Equity Fee set to be used by this tab when in Edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sTransferOfEquityFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTransferOfEquityFeeSet
' Description   : Returns the Transfer of Equity Fee set used by this mortgage product.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTransferOfEquityFeeSet() As String
    On Error GoTo Failed
    
    GetTransferOfEquityFeeSet = m_sTransferOfEquityFeeSet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKeyMatchValues
' Description   : Returns the keys for the currently selected Transfer of Equity Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmProductDetails.lvTransferOfEquityFees.SelectedItem.Index
    
    Set clsPopulateDetails = frmProductDetails.lvTransferOfEquityFees.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Transfer of Equity Fee Keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Transfer of Equity Fee Keys"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the state of fields and buttons on the Transfer of Equity Fees tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields()
On Error GoTo Failed

    frmProductDetails.cmdTransferOfEquityFeeSelect.Enabled = False
    
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


