VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdCredLimIncFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdCredLimIncFee
' Description   : Support class for the Valuation Fee tab on the Mortgage Product screen
' Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const SET_KEY As Integer = 1
Private Const DATE_KEY As Integer = 2

' Private data
Private m_bIsEdit As Boolean
Private m_colMatchValues As Collection
Private m_sCreditLimitIncreaseFeeSet As String
Private m_clsCreditLimitIncreaseFees As CreditLimitIncreaseFeeTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Set m_clsCreditLimitIncreaseFees = New CreditLimitIncreaseFeeTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called from the Mortgage Product form. Performs specific initialisation for Credit Limit Increase
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
    
    frmProductDetails.lvCreditLimitIncreaseFees.LoadColumnDetails TypeName(Me)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : RePopulates the Valuation Fee listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    On Error GoTo Failed
    
    m_clsCreditLimitIncreaseFees.SetFindNewestCreditLimitIncreaseFeeSet
    
    TableAccess(m_clsCreditLimitIncreaseFees).GetTableData POPULATE_FIRST_BAND
    g_clsFormProcessing.PopulateFromRecordset frmProductDetails.lvCreditLimitIncreaseFees, m_clsCreditLimitIncreaseFees
    
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
    
    frmProductDetails.txtCreditLimitIncreaseFeeSelected.Text = m_sCreditLimitIncreaseFeeSet

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Saves all fields on the Valuation Fees tab - doesn't actually do anything for this
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
' Description   : De-selects the currently selected Valuation Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeSelect()
    On Error GoTo Failed
    
    frmProductDetails.txtCreditLimitIncreaseFeeSelected.Text = ""
    frmProductDetails.cmdCreditLimitIncreaseFeeDeSelect.Enabled = False

    ' Reset the Valuation Fee Set
    m_sCreditLimitIncreaseFeeSet = ""
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Selected
' Description   : Selects the currently highlighted Valuation Fee.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Selected()
    On Error GoTo Failed
    Dim sSet As String
    Dim colValues As New Collection
    
    GetKeyMatchValues colValues
    
    If colValues.Count >= SET_KEY Then
        m_sCreditLimitIncreaseFeeSet = colValues(SET_KEY)
        frmProductDetails.txtCreditLimitIncreaseFeeSelected.Text = colValues(SET_KEY)
        frmProductDetails.cmdCreditLimitIncreaseFeeDeSelect.Enabled = True
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Validates all fields on the Mortgage Product Valuation Fee tab - called externaly,
'                 and doesn't actually validate any fields on this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ValidateScreenData() As Boolean
    Dim bRet As Boolean
    
    bRet = True
    
    ValidateScreenData = bRet
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetFeeSet
' Description   : Sets the Valuation Fee set to be used by this tab when in Edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetFeeSet(sFeeSet As String)
    On Error GoTo Failed
    
    m_sCreditLimitIncreaseFeeSet = sFeeSet
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValuationFeeSet
' Description   : Returns the Valuation Fee set used by this mortgage product.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFeeSet() As String
    On Error GoTo Failed
    
    GetFeeSet = m_sCreditLimitIncreaseFeeSet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKeyMatchValues
' Description   : Returns the keys for the currently selected Valuation Fee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetKeyMatchValues(colValues As Collection)
    On Error GoTo Failed
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    nListIndex = frmProductDetails.lvCreditLimitIncreaseFees.SelectedItem.Index
    
    Set clsPopulateDetails = frmProductDetails.lvCreditLimitIncreaseFees.GetExtra(nListIndex)
    
    If Not clsPopulateDetails Is Nothing Then
        Set colValues = clsPopulateDetails.GetKeyMatchValues()
        If Not colValues Is Nothing Then
            If colValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Valuation Fee Keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain Valuation Fee Keys"
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
' Description   : Sets the state of fields and buttons on the Valuation Fees tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields()
On Error GoTo Failed

    frmProductDetails.cmdCreditLimitIncreaseFeeSelect.Enabled = False
    
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
