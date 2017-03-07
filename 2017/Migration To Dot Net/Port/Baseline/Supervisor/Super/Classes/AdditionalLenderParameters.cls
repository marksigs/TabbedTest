VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdditionalLenderParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : AdditionalLenderParameters
' Description   : Support class for the Lender form. Contains code to save/load and validate all
'                 data for this tab.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'BMIDS736 GHun
' Text Edits
Private Const ACCRUED_DAYS_ADDED                As Integer = 0
Private Const SHIFT_IN_DAYS                     As Integer = 1
Private Const SHIFT_IN_MONTHS                   As Integer = 2
Private Const COOLING_OFF_PERIOD                As Integer = 3
Private Const SPECIFIED_PAYMENT_DAY             As Integer = 4
Private Const DAYS_IN_YEAR                      As Integer = 5
Private Const ACCRUED_INTEREST_ROUNDING_FACTOR  As Integer = 6
Private Const PAYMENT_ROUNDING_FACTOR           As Integer = 7
Private Const BALANCE_ROUNDING_FACTOR           As Integer = 8
Private Const CHARGE_ROUNDING_FACTOR            As Integer = 9

'GHun No longer used
' CheckBoxes
'Private Const INCLUDE_ACCRUED_INTEREST = 0
'Private Const INCLUDE_IN_ACCRUED_INTEREST = 1

' Combo's
Private Const ACCRUED_INTEREST_INDICATOR            As Integer = 0
Private Const ACCRUED_INTEREST_PAYABLE              As Integer = 1
Private Const ACCRUED_DAYS_INCLUDED                 As Integer = 2
Private Const FIRST_PAYMENT_DUE_DATE                As Integer = 3
Private Const ACCOUNTING_START_MONTH                As Integer = 4
Private Const INTEREST_CHARGED                      As Integer = 5
Private Const INTEREST_CHARGE_MONTH                 As Integer = 6
Private Const APR_MONTH                             As Integer = 7
Private Const OUTPUT_PAYMENT_SCHEDULE               As Integer = 8
Private Const OUTPUT_BALANCE_SCHEDULE               As Integer = 9
Private Const POSITION_IN_MONTH                     As Integer = 10
Private Const ACCRUED_INTEREST_ROUNDING_DIRECTION   As Integer = 11
Private Const PAYMENT_ROUNDING_DIRECTION            As Integer = 12
Private Const BALANCE_ROUNDING_DIRECTION            As Integer = 13
Private Const CHARGE_ROUNDING_DIRECTION             As Integer = 14
Private Const WHEN_ACC_INT_ROUNDED                  As Integer = 15
Private Const WHEN_PAYMENT_ROUNDED                  As Integer = 16
Private Const WHEN_BALANCE_ROUNDED                  As Integer = 17

'GHun No longer used
'Private Const DEVON_RULING_DEFAULT As Integer = 1
'Private Const NET_GROSS_PROFILE_DEFAULT As String = "N"
'BMIDS736 End

' Private data
Private m_bIsEdit As Boolean
Private m_clsAdditionalParametersTable As MortLenderParamsTable
Private m_colKeyMatchValues As Collection
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to initialise the Lender Parameters tab. Will load all data required
'                 by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'---------------------------------------------------------------------------------------
'BMIDS History
'Prog   Date        Description
'AW     26/06/02    BMIDS00093  Added NET_GROSS_PROFILE_IND combo
'GHun   19/04/2004  BMIDS736 CORE49 Added extra fields for Alpha+
'---------------------------------------------------------------------------------------

'BMIDS736 GHun No longer used
'Private Const NET_GROSS_PROFILE_IND = 6
'BMIDS736 End

Public Sub Initialise(bIsEdit As Boolean)
    On Error GoTo Failed
    
    m_bIsEdit = bIsEdit
    
    If bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    PopulateCombos
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetEditState()
    On Error GoTo Failed
    
    TableAccess(m_clsAdditionalParametersTable).GetTableData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetAddState()
    On Error GoTo Failed
    Dim vOrgID As Variant
    
    g_clsFormProcessing.CreateNewRecord m_clsAdditionalParametersTable
    
    ' Set the key to the table
    vOrgID = m_colKeyMatchValues(ORGANISATION_ID_KEY)
    m_clsAdditionalParametersTable.SetOrganisationId vOrgID
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function GetComboGroup(ByVal nComboID As Integer) As String

    On Error GoTo Failed
    Dim sGroup As String
    Const sFunctionName As String = "GetComboGroup"
    
    Select Case nComboID
        Case ACCRUED_INTEREST_INDICATOR
            sGroup = "AccInterestInd"
        Case ACCRUED_INTEREST_PAYABLE
            sGroup = "AccInterestPayableInd"
        Case FIRST_PAYMENT_DUE_DATE
            sGroup = "FirstPaymentDueDateInd"
        Case ACCOUNTING_START_MONTH
            sGroup = "MonthsInYear"
        Case INTEREST_CHARGED
            sGroup = "InterestChargedInd"
        Case ACCRUED_INTEREST_ROUNDING_DIRECTION
            sGroup = "RoundingDirection"
        'BMIDS736 GHun
        'Case INTEREST_CHANGE_INDICATOR
        '    sGroup = "InterestChangeInd"
        'Case PART_AND_PART_ROUNDING
        '    sGroup = "PartAndPartRounding"
        Case INTEREST_CHARGE_MONTH
            sGroup = "DaysPerMonth"
            'AW 26/06/02    BMids00093
        'Case NET_GROSS_PROFILE_IND
        '    sGroup = "NetGrossProfileInd"
            'AW 26/06/02    BMids00093 -  End
        'BMIDS736 End
        Case PAYMENT_ROUNDING_DIRECTION
            sGroup = "RoundingDirection"
        'BMIDS736 GHun No longer used
        'Case INTEREST_CALC_ROUNDING_DIRECTION
        '    sGroup = "RoundingDirection"
        'BMIDS736 End
        Case POSITION_IN_MONTH
            sGroup = "PositionInMonth"
        'BMIDS736 GHun
        Case ACCRUED_DAYS_INCLUDED
            sGroup = "DaysIncludedInAccrued"
        Case APR_MONTH
            sGroup = "DaysPerMonth"
        Case WHEN_ACC_INT_ROUNDED
            sGroup = "RoundingTiming"
        Case WHEN_PAYMENT_ROUNDED
            sGroup = "RoundingTiming"
        Case WHEN_BALANCE_ROUNDED
            sGroup = "RoundingTiming"
        Case BALANCE_ROUNDING_DIRECTION
            sGroup = "RoundingDirection"
        Case CHARGE_ROUNDING_DIRECTION
            sGroup = "RoundingDirection"
        Case OUTPUT_PAYMENT_SCHEDULE
            sGroup = "OutputPaymentSchedule"
        Case OUTPUT_BALANCE_SCHEDULE
            sGroup = "OutputBalanceSchedule"
        'BMIDS736 End
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName & ": Unknown comboID & " & nComboID
    End Select
    
    GetComboGroup = sGroup
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub PopulateCombo(nComboIndex As Integer)
    On Error GoTo Failed
    Dim sFunctionName As String
    Dim sGroup As String
    
    sFunctionName = "PopulateCombo"
    sGroup = GetComboGroup(nComboIndex)
    
    If Len(sGroup) > 0 Then
        g_clsFormProcessing.PopulateCombo sGroup, frmLenderDetails.cboAdditionalParams(nComboIndex)
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Invalid Group"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetComboValidation(nComboIndex As Integer, sValidation As String)
    On Error GoTo Failed
    Dim sGroup As String
    Dim sFunctionName  As String
    
    sFunctionName = "SetComboValidation"
    
    sGroup = GetComboGroup(nComboIndex)
    
    If Len(sGroup) > 0 Then
        g_clsFormProcessing.SetComboFromValidation frmLenderDetails.cboAdditionalParams(nComboIndex), sGroup, sValidation
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Group is empty"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetComboValidation(nComboIndex As Integer, sValidation As String)
    On Error GoTo Failed
    Dim sGroup As String
    Dim sFunctionName As String
    
    sFunctionName = "GetComboValidation"
    sValidation = ""
    
    sGroup = GetComboGroup(nComboIndex)
    
    If Len(sGroup) > 0 Then
        g_clsFormProcessing.GetComboValidation frmLenderDetails.cboAdditionalParams(nComboIndex), sGroup, sValidation
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Invalid Group"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub PopulateCombos()

    On Error GoTo Failed
    
    PopulateCombo ACCRUED_INTEREST_INDICATOR
    PopulateCombo ACCRUED_INTEREST_INDICATOR
    PopulateCombo ACCRUED_INTEREST_PAYABLE
    PopulateCombo FIRST_PAYMENT_DUE_DATE
    PopulateCombo ACCOUNTING_START_MONTH
    PopulateCombo INTEREST_CHARGED
    PopulateCombo ACCRUED_INTEREST_ROUNDING_DIRECTION
    'BMIDS736 GHun
    'PopulateCombo INTEREST_CHANGE_INDICATOR
    'PopulateCombo PART_AND_PART_ROUNDING
    PopulateCombo INTEREST_CHARGE_MONTH
    'BMIDS736 End
    PopulateCombo PAYMENT_ROUNDING_DIRECTION
    'BMIDS736 GHun No longer used
    'AW 26/06/02    BMids00093
    'PopulateCombo NET_GROSS_PROFILE_IND
    'AW 26/06/02    BMids00093  -  End
    'PopulateCombo INTEREST_CALC_ROUNDING_DIRECTION
    'BMIDS736 End
    PopulateCombo POSITION_IN_MONTH
    'BMIDS736 GHun
    PopulateCombo ACCRUED_DAYS_INCLUDED
    PopulateCombo APR_MONTH
    PopulateCombo WHEN_ACC_INT_ROUNDED
    PopulateCombo WHEN_PAYMENT_ROUNDED
    PopulateCombo WHEN_BALANCE_ROUNDED
    PopulateCombo BALANCE_ROUNDING_DIRECTION
    PopulateCombo CHARGE_ROUNDING_DIRECTION
    PopulateCombo OUTPUT_PAYMENT_SCHEDULE
    PopulateCombo OUTPUT_BALANCE_SCHEDULE
    'BMIDS736 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, "AdditionalLenderParameters:PopulateCombos - " + Err.DESCRIPTION
End Sub


Public Sub SetScreenFields()

    On Error GoTo Failed
    Dim sValidation As String
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_INDICATOR), m_clsAdditionalParametersTable.GetAccInterestInd(), SET_CONTROL_VALUE
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_PAYABLE), m_clsAdditionalParametersTable.GetAccInterestPayableInd(), SET_CONTROL_VALUE
    frmLenderDetails.txtAdditionalParams(ACCRUED_INTEREST_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetAccInterestRoundingFactor()
    
    sValidation = m_clsAdditionalParametersTable.GetAccInterestRoundingDirection()
    SetComboValidation ACCRUED_INTEREST_ROUNDING_DIRECTION, sValidation
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(FIRST_PAYMENT_DUE_DATE), m_clsAdditionalParametersTable.GetFirstPaymentDueDateInd(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCOUNTING_START_MONTH), m_clsAdditionalParametersTable.GetAccountingStartMonth(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHARGED), m_clsAdditionalParametersTable.GetInterestChargedInd(), SET_CONTROL_VALUE
    'BMIDS736 GHun No longer used
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHANGE_INDICATOR), m_clsAdditionalParametersTable.GetInterestChangeInd(), SET_CONTROL_VALUE
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(PART_AND_PART_ROUNDING), m_clsAdditionalParametersTable.GetPartAndPartRounding(), SET_CONTROL_VALUE
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(MONTHLY_PAYMENT_INDICATOR), m_clsAdditionalParametersTable.GetMonthlyPaymentInd(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHARGE_MONTH), m_clsAdditionalParametersTable.GetMonthlyPaymentInd(), SET_CONTROL_VALUE
    'AW 26/06/02    BMids00093
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(NET_GROSS_PROFILE_IND), m_clsAdditionalParametersTable.GetNetGrossProfileInd, SET_CONTROL_VALUE
    'BMIDS736 End
    
    frmLenderDetails.txtAdditionalParams(PAYMENT_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetPaymentRoundingFactor()
       
    sValidation = m_clsAdditionalParametersTable.GetPaymentRoundingDirection()
    SetComboValidation PAYMENT_ROUNDING_DIRECTION, sValidation
    
    'BMIDS736 GHun No longer used
    'frmLenderDetails.txtAdditionalParams(INTEREST_CALC_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetInterestCalcRoundingFactor()
    'sValidation = m_clsAdditionalParametersTable.GetInterestCalcRoundingDirection()
    'SetComboValidation INTEREST_CALC_ROUNDING_DIRECTION, sValidation
    'BMIDS736 End
    
    frmLenderDetails.txtAdditionalParams(SHIFT_IN_MONTHS).Text = m_clsAdditionalParametersTable.GetShiftInMonths()
    
    sValidation = m_clsAdditionalParametersTable.GetPositionInMonth()
    SetComboValidation POSITION_IN_MONTH, sValidation
    
    frmLenderDetails.txtAdditionalParams(SHIFT_IN_DAYS).Text = m_clsAdditionalParametersTable.GetShiftInDays()
        
    'BMIDS736 GHun
    'g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_ACCRUED_INTEREST), m_clsAdditionalParametersTable.GetIncludeACCInterest(), SET_CONTROL_VALUE
    
    frmLenderDetails.txtAdditionalParams(COOLING_OFF_PERIOD).Text = m_clsAdditionalParametersTable.GetCoolingOffPeriod()
    frmLenderDetails.txtAdditionalParams(SPECIFIED_PAYMENT_DAY).Text = m_clsAdditionalParametersTable.GetSpecifiedPaymentDay()
    frmLenderDetails.txtAdditionalParams(DAYS_IN_YEAR).Text = m_clsAdditionalParametersTable.GetDaysInYear()
    frmLenderDetails.txtAdditionalParams(ACCRUED_DAYS_ADDED).Text = m_clsAdditionalParametersTable.GetAccruedDaysAdded()
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_DAYS_INCLUDED), m_clsAdditionalParametersTable.GetAccruedDaysIncluded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(APR_MONTH), m_clsAdditionalParametersTable.GetAPRMonth(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_ACC_INT_ROUNDED), m_clsAdditionalParametersTable.GetWhenAccIntRounded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_PAYMENT_ROUNDED), m_clsAdditionalParametersTable.GetWhenPaymentRounded(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_BALANCE_ROUNDED), m_clsAdditionalParametersTable.GetWhenBalanceRounded(), SET_CONTROL_VALUE
        
    frmLenderDetails.txtAdditionalParams(BALANCE_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetBalanceRoundingFactor()
    
    sValidation = m_clsAdditionalParametersTable.GetBalanceRoundingDirection()
    SetComboValidation BALANCE_ROUNDING_DIRECTION, sValidation
    
    frmLenderDetails.txtAdditionalParams(CHARGE_ROUNDING_FACTOR).Text = m_clsAdditionalParametersTable.GetChargeRoundingFactor()
    
    sValidation = m_clsAdditionalParametersTable.GetChargeRoundingDirection()
    SetComboValidation CHARGE_ROUNDING_DIRECTION, sValidation
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(OUTPUT_PAYMENT_SCHEDULE), m_clsAdditionalParametersTable.GetOutputPaymentSchedule(), SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(OUTPUT_BALANCE_SCHEDULE), m_clsAdditionalParametersTable.GetOutputBalanceSchedule(), SET_CONTROL_VALUE
    'BMIDS736 END
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()

    On Error GoTo Failed
    Dim vTmp As Variant
    Dim sValidation As String
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_INDICATOR), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccInterestInd vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_INTEREST_PAYABLE), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccInterestPayableInd vTmp
    
    m_clsAdditionalParametersTable.SetACCInterestRoundingFactor frmLenderDetails.txtAdditionalParams(ACCRUED_INTEREST_ROUNDING_FACTOR).Text
    
    GetComboValidation ACCRUED_INTEREST_ROUNDING_DIRECTION, sValidation
    m_clsAdditionalParametersTable.SetAccInterestRoundingDirection sValidation
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(FIRST_PAYMENT_DUE_DATE), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetFirstPaymentDueDateInd vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCOUNTING_START_MONTH), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccountingStartMonth vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHARGED), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetInterestChargedInd vTmp
    
    'BMIDS736 GHun No longer used
    'm_clsAdditionalParametersTable.SetDevonRulingInd DEVON_RULING_DEFAULT

    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHANGE_INDICATOR), vTmp, GET_CONTROL_VALUE
    'm_clsAdditionalParametersTable.SetInterestChangeInd vTmp
    
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(PART_AND_PART_ROUNDING), vTmp, GET_CONTROL_VALUE
    'm_clsAdditionalParametersTable.SetPartAndPartRounding vTmp
    'AW 26/06/02    BMids00093
    'g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(NET_GROSS_PROFILE_IND), vTmp, GET_CONTROL_VALUE
    'm_clsAdditionalParametersTable.SetNetGrossProfileInd vTmp
    'BMIDS736 End
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(INTEREST_CHARGE_MONTH), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetMonthlyPaymentInd vTmp
    
    m_clsAdditionalParametersTable.SetPaymentRoundingFactor frmLenderDetails.txtAdditionalParams(PAYMENT_ROUNDING_FACTOR).Text
    
    GetComboValidation PAYMENT_ROUNDING_DIRECTION, sValidation
    m_clsAdditionalParametersTable.SetPaymentRoundingDirection sValidation
    
    'BMIDS736 GHun No longer used
    'm_clsAdditionalParametersTable.SetInterestCalcRoundingFactor frmLenderDetails.txtAdditionalParams(INTEREST_CALC_ROUNDING_FACTOR).Text
    
    'GetComboValidation INTEREST_CALC_ROUNDING_DIRECTION, sValidation
    'm_clsAdditionalParametersTable.SetInterestCalcRoundingDirection sValidation
    'BMIDS736 End
    
    m_clsAdditionalParametersTable.SetShiftInMonths frmLenderDetails.txtAdditionalParams(SHIFT_IN_MONTHS).Text
    
    GetComboValidation POSITION_IN_MONTH, sValidation
    m_clsAdditionalParametersTable.SetPositionInMonth sValidation
    
    m_clsAdditionalParametersTable.SetShiftInDays frmLenderDetails.txtAdditionalParams(SHIFT_IN_DAYS).Text
    
    'BMIDS736 GHun
    'g_clsFormProcessing.HandleCheckBox frmLenderDetails.chkAdditionalParams(INCLUDE_ACCRUED_INTEREST), vTmp, GET_CONTROL_VALUE
    'm_clsAdditionalParametersTable.SetIncludeACCInterest vTmp
    
    m_clsAdditionalParametersTable.SetCoolingOffPeriod frmLenderDetails.txtAdditionalParams(COOLING_OFF_PERIOD).Text
    m_clsAdditionalParametersTable.SetSpecifiedPaymentDay frmLenderDetails.txtAdditionalParams(SPECIFIED_PAYMENT_DAY).Text
    m_clsAdditionalParametersTable.SetDaysInYear frmLenderDetails.txtAdditionalParams(DAYS_IN_YEAR).Text
    m_clsAdditionalParametersTable.SetAccruedDaysAdded frmLenderDetails.txtAdditionalParams(ACCRUED_DAYS_ADDED).Text
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(ACCRUED_DAYS_INCLUDED), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAccruedDaysIncluded vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(APR_MONTH), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetAPRMonth vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_ACC_INT_ROUNDED), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetWhenAccIntRounded vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_PAYMENT_ROUNDED), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetWhenPaymentRounded vTmp
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(WHEN_BALANCE_ROUNDED), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetWhenBalanceRounded vTmp
    
    m_clsAdditionalParametersTable.SetBalanceRoundingFactor frmLenderDetails.txtAdditionalParams(BALANCE_ROUNDING_FACTOR).Text
    GetComboValidation BALANCE_ROUNDING_DIRECTION, sValidation
    m_clsAdditionalParametersTable.SetBalanceRoundingDirection sValidation
    
    m_clsAdditionalParametersTable.SetChargeRoundingFactor frmLenderDetails.txtAdditionalParams(CHARGE_ROUNDING_FACTOR).Text
    GetComboValidation CHARGE_ROUNDING_DIRECTION, sValidation
    m_clsAdditionalParametersTable.SetChargeRoundingDirection sValidation
    
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(OUTPUT_PAYMENT_SCHEDULE), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetOutputPaymentSchedule vTmp
    g_clsFormProcessing.HandleComboExtra frmLenderDetails.cboAdditionalParams(OUTPUT_BALANCE_SCHEDULE), vTmp, GET_CONTROL_VALUE
    m_clsAdditionalParametersTable.SetOutputBalanceSchedule vTmp
    'BMIDS736 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsAdditionalParametersTable).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetKeyValues(colValues As Collection)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsAdditionalParametersTable
    
    clsTableAccess.SetKeyMatchValues colValues
    Set m_colKeyMatchValues = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitTab()

End Sub

Public Function ValidateScreenData() As Boolean
    ValidateScreenData = True
End Function

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Private Sub Class_Initialize()
    If m_clsAdditionalParametersTable Is Nothing Then
        Set m_clsAdditionalParametersTable = New MortLenderParamsTable
    End If
End Sub

