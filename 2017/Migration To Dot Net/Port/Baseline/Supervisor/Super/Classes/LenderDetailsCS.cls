VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderDetailsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderDetailsCS
' Description   : Client specific version of the LenderDetails class. This class will be called by
'                 a form or controlling class and any client specific code can be put in here for
'                 each client. Each method can a) override the base code, b) delegate to the base
'                 code or c) overried and delegate.
' Change history
' Prog      Date        Description
' DJP       22/10/01    Created.
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' DJP       20/11/01    SYS2831 Support client variants for Lender.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Base class
Implements LenderDetails

' Private data
Private m_clsLenderDetails As LenderDetails

'BMids client

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   AW      13/05/02    BM022, BM024 - Added SaveClientScreenData, SetClientScreenFields,
'                       ValidateClientScreenData
'   AW      28/05/02    BM022 - Updated ValidateClientScreenData
'   MO      10/06/02    BMIDS00040 - Updated ValidateClientScreenData
'

' Text Edit boxes
Private Const ARRANGEMENT_FEE_PERCENT = 11
Private Const ARRANGEMENT_FEE_PERCENT_MAX = 12
Private Const ARRANGEMENT_FEE = 13
Private Const ARRANGEMENT_FEE_PERCENT_MIN = 15

' CheckBoxes
Private Const ADD_RE_VALUATION_FEE = 9

Private Sub Class_Initialize()
    Set m_clsLenderDetails = New LenderDetails
End Sub


Private Sub LenderDetails_DoUpdates()
    m_clsLenderDetails.DoUpdates
End Sub


Private Function LenderDetails_GetAddressGUID() As Variant
    LenderDetails_GetAddressGUID = m_clsLenderDetails.GetAddressGUID
End Function


Private Function LenderDetails_GetContactDetailsGUID() As Variant
    LenderDetails_GetContactDetailsGUID = m_clsLenderDetails.GetContactDetailsGUID
End Function


Private Function LenderDetails_GetFeeSetKey() As Variant
    LenderDetails_GetFeeSetKey = m_clsLenderDetails.GetFeeSetKey
End Function


Private Function LenderDetails_GetLenderTable() As MortgageLendersTable
    LenderDetails_GetLenderTable = m_clsLenderDetails.GetLenderTable
End Function


Private Function LenderDetails_GetOrganisationID() As Variant
    LenderDetails_GetOrganisationID = m_clsLenderDetails.GetOrganisationID
End Function


Private Function LenderDetails_GetRateSet() As String
    LenderDetails_GetRateSet = m_clsLenderDetails.GetRateSet
End Function


Private Sub LenderDetails_Initialise(bIsEdit As Boolean)
    m_clsLenderDetails.Initialise bIsEdit
End Sub


Private Sub LenderDetails_InitTab()
    m_clsLenderDetails.InitTab
End Sub


Private Function LenderDetails_IsEdit() As Boolean
    LenderDetails_IsEdit = m_clsLenderDetails.IsEdit
End Function

Private Sub LenderDetails_SaveScreenData()
    m_clsLenderDetails.SaveScreenData
    SaveClientScreenData
End Sub

Private Sub LenderDetails_SetIsEdit(Optional bEdit As Boolean = True)
    m_clsLenderDetails.SetIsEdit bEdit
End Sub


Private Sub LenderDetails_SetMIGRateKey(sRateSet As String)
    m_clsLenderDetails.SetMIGRateKey sRateSet
End Sub

Private Sub LenderDetails_SetScreenFields()

    m_clsLenderDetails.SetScreenFields
    SetClientScreenFields
End Sub

Private Sub LenderDetails_SetTableClass(clsTableAccess As TableAccess)
    m_clsLenderDetails.SetTableClass clsTableAccess
End Sub

Private Function LenderDetails_ValidateScreenData() As Boolean
    LenderDetails_ValidateScreenData = m_clsLenderDetails.ValidateScreenData
    LenderDetails_ValidateScreenData = ValidateClientScreenData
End Function

Private Sub SetClientScreenFields()
    On Error GoTo Failed
    Dim clsMortProdLender As MortgageLendersTable
    Set clsMortProdLender = m_clsLenderDetails.GetLenderTable
    
    frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN).Text = clsMortProdLender.GetArrangementFeePercentMin()
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_RE_VALUATION_FEE), clsMortProdLender.GetAllowReValuationFeeAdded(), SET_CONTROL_VALUE
    'Set clsMortProdLender = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SaveClientScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsMortProdLender As MortgageLendersTable
    
    Set clsMortProdLender = m_clsLenderDetails.GetLenderTable
    
    clsMortProdLender.SetArrangementFeePercentMin frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN).Text
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkLenderDetails(ADD_RE_VALUATION_FEE), vTmp, GET_CONTROL_VALUE
    clsMortProdLender.SetAllowReValuationFeeAdded vTmp
    'Set clsMortProdLender = Nothing
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidateClientScreenData() As Boolean
    On Error GoTo Failed
    Dim sFee As String
    Dim sFeePercentMax As String
    Dim sFeePercentMin As String
    Dim sFeePercent As String
    
    sFee = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE).Text
    sFeePercentMax = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MAX).Text
    sFeePercentMin = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN).Text
    sFeePercent = frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT).Text
    
    If Val(sFee) > 0 And Val(sFeePercent) > 0 Then
        g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT)
        g_clsErrorHandling.RaiseError errGeneralError, "Only an arrangement fee amount OR an arrangement fee percentage can be entered"
    End If
    
    If Len(sFee) > 0 And (Val(sFeePercentMin) > 0 Or Val(sFeePercentMax) > 0) Then
        g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN)
        g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum arrangement fee values can only be applied against the arrangement fee percentage"
    End If
    
    If Val(sFeePercent) = 0 And (Val(sFeePercentMin) > 0 Or Val(sFeePercentMax) > 0) Then
        g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN)
        g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum arrangement fee values cannot be specified when there is no arrangement fee percentage"
    End If
    
    If Len(sFeePercentMax) > 0 Then
        If Val(sFeePercentMin) > Val(sFeePercentMax) Then
            g_clsFormProcessing.SetControlFocus frmLenderDetails.txtLenderDetails(ARRANGEMENT_FEE_PERCENT_MIN)
            g_clsErrorHandling.RaiseError errGeneralError, "Maximum arrangement fee must be greater than or equal to the minimum fee"
        End If
    End If
    
    ValidateClientScreenData = True
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

