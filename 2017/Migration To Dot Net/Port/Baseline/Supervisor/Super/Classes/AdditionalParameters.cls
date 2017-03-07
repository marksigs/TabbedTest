VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdditionalParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Text Edits
Private Const ACCRUED_INTEREST_ROUNDING_FACTOR = 0
Private Const PAYMENT_ROUNDING_FACTOR = 1
Private Const INTEREST_CALC_ROUNDING_FACTOR = 2
Private Const SHIFT_IN_MONTHS = 3
Private Const POSITION_IN_MONTH = 4
Private Const SHIFT_IN_DAYS = 5

' CheckBoxes
Private Const INCLUDE_ACCRUED_INTEREST = 0
Private Const INCLUDE_IN_ACCRUED_INTEREST = 1

' Combo's
Private Const ACCRUED_INTEREST_INDICATOR = 0
Private Const ACCRUED_INTEREST_PAYABLE = 1
Private Const ACCRUED_INTEREST_ROUNDING_DIRECTION = 2
Private Const FIRST_PAYMENT_DUE_DATE = 3
Private Const ACCOUNTING_START_MONTH = 4
Private Const INTEREST_CHARGED = 5
Private Const DEVON_RULING_INDICATOR = 6
Private Const INTEREST_CHANGE_INDICATOR = 7
Private Const PART_AND_PART_ROUNDING = 8
Private Const MONTHLY_PAYMENT_INDICATOR = 9
Private Const NET_GROSS_PROFILE_INDICATOR = 10
Private Const PAYMENT_ROUNDING_DIRECTION = 11
Private Const INTEREST_CALC_ROUNDING_DIRECTION = 12

Private m_rs As ADODB.Recordset
Private m_clsTableAccess As TableAccess
Private m_clsAdditionalParametersTable As MortgageLenderParametersTable
Sub Class_Initialize()

End Sub
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsAdditionalParametersTable = clsTableAccess
End Sub
Public Function Initialise() As Boolean
    If m_clsAdditionalParametersTable Is Nothing Then
        Set m_clsAdditionalParametersTable = New MortgageLenderParametersTable
    End If
    
    Set m_clsTableAccess = m_clsAdditionalParametersTable
    Set m_rs = m_clsTableAccess.GetTableData()
    Initialise = True
    Exit Function
Failed:
    MsgBox "AdditionalParameters:Initialise - Error is " + Err.Description
    Initialise = False

End Function
Public Function SetScreenFields()
    Dim bRet As Boolean
    On Error GoTo Failed
    Dim a As String
    bRet = True
    
    If Not m_rs Is Nothing Then
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_INDICATOR), m_clsAdditionalParametersTable.GetAccInterestInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_PAYABLE), m_clsAdditionalParametersTable.GetAccInterestPayableInd(), SET_CONTROL_VALUE
        frmLenderDetails.txtAdditionalParams(ACCRUED_INTEREST_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetAccInterestRoundingFactor()
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_ROUNDING_DIRECTION), m_clsAdditionalParametersTable.GetAccInterestRoundingDirection(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(FIRST_PAYMENT_DUE_DATE), m_clsAdditionalParametersTable.GetFirstPaymentDueDateInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCOUNTING_START_MONTH), m_clsAdditionalParametersTable.GetAccountingStartMonth(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CHARGED), m_clsAdditionalParametersTable.GetInterestChargedInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(DEVON_RULING_INDICATOR), m_clsAdditionalParametersTable.GetDevonRulingInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CHANGE_INDICATOR), m_clsAdditionalParametersTable.GetInterestChangeInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(PART_AND_PART_ROUNDING), m_clsAdditionalParametersTable.GetPartAndPartRounding(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(MONTHLY_PAYMENT_INDICATOR), m_clsAdditionalParametersTable.GetMonthlyPaymentInd(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(NET_GROSS_PROFILE_INDICATOR), m_clsAdditionalParametersTable.GetNetGrossProfileInd(), SET_CONTROL_VALUE
        frmLenderDetails.txtAdditionalParams(PAYMENT_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetPaymentRoundingFactor()
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(PAYMENT_ROUNDING_DIRECTION), m_clsAdditionalParametersTable.GetPaymentRoundingDirection(), SET_CONTROL_VALUE
        frmLenderDetails.txtAdditionalParams(INTEREST_CALC_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetInterestCalcRoundingFactor()
        g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CALC_ROUNDING_DIRECTION), m_clsAdditionalParametersTable.GetInterestCalcRoundingDirection(), SET_CONTROL_VALUE
        frmLenderDetails.txtAdditionalParams(SHIFT_IN_MONTHS).Text = m_clsAdditionalParametersTable.GetShiftInMonths()
        frmLenderDetails.txtAdditionalParams(POSITION_IN_MONTH).Text = m_clsAdditionalParametersTable.GetPositionInMonth()
        frmLenderDetails.txtAdditionalParams(SHIFT_IN_DAYS).Text = m_clsAdditionalParametersTable.GetShiftInDays()
        g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_ACCRUED_INTEREST), m_clsAdditionalParametersTable.GetIncludeACCInterest(), SET_CONTROL_VALUE
        g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_IN_ACCRUED_INTEREST), m_clsAdditionalParametersTable.GetIncludeInACCInterest(), SET_CONTROL_VALUE
    End If
    SetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "SetScreenFields: Error is " + Err.Description
    SetScreenFields = False
End Function
Public Function GetScreenFields()
    Dim bRet As Boolean
    Dim vTmp As Variant
    bRet = True
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccInterestPayableInd vTmp
    
    m_clsAdditionalParametersTable.SetACCInterestRoundingFactor frmLenderDetails.txtAdditionalParams(ACCRUED_INTEREST_ROUNDING_FACTOR).Text
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_ROUNDING_DIRECTION), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccInterestRoundingDirection vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(FIRST_PAYMENT_DUE_DATE), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetFirstPaymentDueDateInd vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(ACCOUNTING_START_MONTH), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccountingStartMonth vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CHARGED), vTmp, SET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetInterestChargedInd vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(DEVON_RULING_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetDevonRulingInd vTmp

    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CHANGE_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetInterestChangeInd vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(PART_AND_PART_ROUNDING), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetPartAndPartRounding vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(MONTHLY_PAYMENT_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetMonthlyPaymentInd vTmp
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(NET_GROSS_PROFILE_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetNetGrossProfileInd vTmp
    
    m_clsAdditionalParametersTable.SetPaymentRoundingFactor frmLenderDetails.txtAdditionalParams(PAYMENT_ROUNDING_FACTOR).Text
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(PAYMENT_ROUNDING_DIRECTION), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetPaymentRoundingDirection vTmp
    
    m_clsAdditionalParametersTable.SetInterestCalcRoundingFactor frmLenderDetails.txtAdditionalParams(INTEREST_CALC_ROUNDING_FACTOR).Text
    
    g_clsFormProcessing.HandleComboText frmLenderDetails.cboAdditionalParams(INTEREST_CALC_ROUNDING_DIRECTION), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetInterestCalcRoundingDirection vTmp
    
    m_clsAdditionalParametersTable.SetShiftInMonths frmLenderDetails.txtAdditionalParams(SHIFT_IN_MONTHS).Text
    m_clsAdditionalParametersTable.SetPositionInMonth frmLenderDetails.txtAdditionalParams(POSITION_IN_MONTH).Text
    m_clsAdditionalParametersTable.SetShiftInDays frmLenderDetails.txtAdditionalParams(SHIFT_IN_DAYS).Text
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_ACCRUED_INTEREST), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetIncludeACCInterest vTmp
    
    g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_IN_ACCRUED_INTEREST), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetIncludeInACCInterest vTmp
    
    GetScreenFields = bRet
    Exit Function
Failed:
    MsgBox "GetScreenFields: Error is " + Err.Description
    GetScreenFields = False

End Function
Friend Sub SetKeyValues(colValues As Collection)
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsAdditionalParametersTable
    
    clsTableAccess.SetKeyMatchValues colValues
    
End Sub
