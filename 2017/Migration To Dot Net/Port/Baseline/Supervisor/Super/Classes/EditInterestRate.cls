VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditInterestRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : EditInterestRate
' Description   : ClassTo Add and Edit Interest Rates
'
' Change history
' Prog      Date        Description
' DJP       03/12/01    SYS2912 SQL Server locking problem.
' STB       06/12/01    SYS1942 - 'Another' button / transactions.
' DJP       18/12/01    SYS1942 - Another' button
' DJP       18/12/01    SYS2831 Client variant support
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants - text edits
Private Const SEQUENCE_NUMBER = 0
Private Const RATE_PERIOD = 1
Private Const END_DATE = 2
Private Const INTEREST_RATE = 3
Private Const FLOOR_RATE = 4
Private Const CEILING_RATE = 5
Private Const REDEMPTION_PERIOD = 6
Private Const REPAYMENT_PERIOD = 7

Private Const TYPE_COUNT As Integer = 4

' Application constants
Private Const TYPE_BASE_SAVE = "B"
Private Const TYPE_FIXED_SAVE = "F"
Private Const TYPE_CAPPED_SAVE = "C"
Private Const TYPE_DISCOUNT_SAVE = "D"

Private Const TYPE_BASE = "Base"
Private Const TYPE_FIXED = "Fixed"
Private Const TYPE_CAPPED = "Capped/Floored"
Private Const TYPE_DISCOUNT = "Discount"

'Used to indicate if cmdAnother has been used, this will store the number of
'records added in one 'session'.
Private m_iRecordsAdded As Integer

' Private data
Private m_bIsEdit As Boolean
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_Rates(TYPE_COUNT) As InterestRateTypes
Private m_ReturnCode As MSGReturnCode
Private m_clsInterestRates As MortProdIntRatesTable
Private m_colMatchValues As Collection

Private m_Frm As frmEditInterestRates
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Need to set this class as being the support
'                 class it will use.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009. Amended : PopulateCombos
'                                   Added :PopulateBaseRateSetCombo,GetInterestRateTypeTable
'AW         11/07/02    BMIDS00177  Removed processing of redundant controls

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    ' Initialise the form
    Set m_Frm = frmEditInterestRates
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Show()
    On Error GoTo Failed

    m_Frm.Show vbModal, frmMain
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm frmLenderDetails
End Sub

Public Sub Load()
    On Error GoTo Failed
    
    'Set the number of records modified to one.
    m_iRecordsAdded = 1

    SetReturnCode MSGFailure

    ' Read combo table for rates.
    SetupRateCombo
    PopulateCombos

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetAddState()
    On Error GoTo Failed

    m_frm.cmdAnother.Enabled = True
    m_frm.txtInterestRates(SEQUENCE_NUMBER).Enabled = True
    'm_frm.txtInterestRates(REDEMPTION_PERIOD).Enabled = True
    'm_frm.txtInterestRates(REPAYMENT_PERIOD).Enabled = True

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetEditState()
    On Error GoTo Failed
    TableAccess(m_clsInterestRates).ApplyFilter

    PopulateScreenFields

    ' Set state of buttons - if editing, the Another button and the SEQUENCE  number are disabled.
    m_Frm.cmdAnother.Enabled = False
    m_Frm.txtInterestRates(SEQUENCE_NUMBER).Enabled = False

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub PopulateCombos()
'GD BMIDS
    On Error GoTo Failed

    'g_clsFormProcessing.PopulateCombo "InterestCalculationType", m_frm.cboCalculationType
    'g_clsFormProcessing.PopulateCombo "EarlyRedemptionTerms", m_frm.cboRedemptionTerms
    'g_clsFormProcessing.PopulateCombo "PartRepaymentTerms", m_frm.cboRepaymentTerms
    'GD BMIDS
    PopulateBaseRateSetCombo
    'Dim colFields As Collection
    'Set colFields = New Collection
    'colFields.Add "RATEID"
    'colFields.Add "BASERATESETDESCRIPTION"
    'Set m_clsBaseRateSet = New BaseRateSetTable
    'g_clsFormProcessing.PopulateComboFromTable m_frm.cboBaseRateSet, m_clsBaseRateSet, colFields
    'GD BMIDS END
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim vTmp As Variant

    m_frm.txtInterestRates(SEQUENCE_NUMBER).Text = m_clsInterestRates.GetIntRateSeqNumber()
    m_frm.txtInterestRates(RATE_PERIOD).Text = m_clsInterestRates.GetInterestRatePeriod()
    vTmp = m_clsInterestRates.GetInterestRateEndDate()
    g_clsFormProcessing.HandleDate m_frm.txtInterestRates(END_DATE), vTmp, SET_CONTROL_VALUE

    'Populate text fields
    m_frm.txtInterestRates(INTEREST_RATE).Text = m_clsInterestRates.GetRate()
    m_frm.txtInterestRates(FLOOR_RATE).Text = m_clsInterestRates.GetFlooredRate()
    m_frm.txtInterestRates(CEILING_RATE).Text = m_clsInterestRates.GetCeilingRate()
    'BMIDS00177
'    m_frm.txtInterestRates(REDEMPTION_PERIOD).Text = m_clsInterestRates.GetRedemptionPeriod()
'    m_frm.txtInterestRates(REDEMPTION_PERIOD).Enabled = True
'    m_frm.txtInterestRates(REPAYMENT_PERIOD).Text = m_clsInterestRates.GetPaymentPeriod()
'    m_frm.txtInterestRates(REPAYMENT_PERIOD).Enabled = True

    'Populate combo's
    GetRateType

'    vTmp = m_clsInterestRates.GetCalculationType()
'    g_clsFormProcessing.HandleComboExtra m_frm.cboCalculationType, vTmp, SET_CONTROL_VALUE
'
'    vTmp = m_clsInterestRates.GetRedemptionTerms()
'    g_clsFormProcessing.HandleComboExtra m_frm.cboRedemptionTerms, vTmp, SET_CONTROL_VALUE
'
'    vTmp = m_clsInterestRates.GetRepaymentTerms()
'    g_clsFormProcessing.HandleComboExtra m_frm.cboRepaymentTerms, vTmp, SET_CONTROL_VALUE


    PopulateScreenFields = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function



Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant

    ' If adding, create a record.
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsInterestRates
    End If

    m_clsInterestRates.SetIntRateSeqNumber m_frm.txtInterestRates(SEQUENCE_NUMBER).Text
    m_clsInterestRates.SetInterestRatePeriod m_frm.txtInterestRates(RATE_PERIOD).Text

    g_clsFormProcessing.HandleDate m_frm.txtInterestRates(END_DATE), vTmp, GET_CONTROL_VALUE
    m_clsInterestRates.SetInterestRateEndDate vTmp

    ' Save the text fields
    m_clsInterestRates.SetRate m_frm.txtInterestRates(INTEREST_RATE).Text
    m_clsInterestRates.SetFlooredRate m_frm.txtInterestRates(FLOOR_RATE).Text
    m_clsInterestRates.SetCeilingRate m_frm.txtInterestRates(CEILING_RATE).Text
    'BMIDS00177
'    m_clsInterestRates.SetRedemptionPeriod m_frm.txtInterestRates(REDEMPTION_PERIOD).Text
'    m_clsInterestRates.SetRepaymentPeriod m_frm.txtInterestRates(REPAYMENT_PERIOD).Text

    ' Save the combo values
    SetRateType
'BMIDS00177
'    g_clsFormProcessing.HandleComboExtra m_frm.cboCalculationType, vTmp, GET_CONTROL_VALUE
'    m_clsInterestRates.SetCalculationType CStr(vTmp)
'
'    g_clsFormProcessing.HandleComboExtra m_frm.cboRedemptionTerms, vTmp, GET_CONTROL_VALUE
'    m_clsInterestRates.SetRedemptionTerms CStr(vTmp)
'
'    g_clsFormProcessing.HandleComboExtra m_frm.cboRepaymentTerms, vTmp, GET_CONTROL_VALUE
'    m_clsInterestRates.SetRepaymentTerms CStr(vTmp)

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sVal As String
    Dim nRateLen As Integer
    Dim nEndDateLen As Integer
    Dim sRatePeriod As String
    Dim sSequenceNumber As String

    sRatePeriod = m_Frm.txtInterestRates(RATE_PERIOD).Text

    bRet = True

    sVal = m_Frm.cboType.SelText

    If sVal = TYPE_CAPPED Then
        If Len(m_Frm.txtInterestRates(CEILING_RATE).Text) = 0 And Len(m_Frm.txtInterestRates(FLOOR_RATE).Text) = 0 Then
            MsgBox "Must enter one or both of Ceiling or Floored rates", vbCritical
            m_Frm.txtInterestRates(FLOOR_RATE).SetFocus
            bRet = False
        End If
    End If

    nRateLen = Len(sRatePeriod)
    nEndDateLen = Len(StripUnderscoreAndSlash(m_Frm.txtInterestRates(END_DATE).Text))

    If nRateLen > 0 And nEndDateLen > 0 Then
        MsgBox "Only one of Rate Period and End Date can be entered", vbCritical
        m_Frm.txtInterestRates(RATE_PERIOD).SetFocus
        bRet = False
    ElseIf nRateLen = 0 And nEndDateLen = 0 Then
        MsgBox "At least one of Rate Period and End Date must be entered", vbCritical
        m_Frm.txtInterestRates(RATE_PERIOD).SetFocus
        bRet = False
    End If

    If bRet = True Then
        ' The Rate Period, if negative, can only be -1

        If nRateLen > 0 Then
            If CLng(sRatePeriod) < -1 Then
                MsgBox "Rate period must be -1 or above"
                m_Frm.txtInterestRates(RATE_PERIOD).SetFocus
                bRet = False
            End If
        End If
    End If

    ' Check if the record being added (if in Add mode) exists or not.
    If bRet And m_bIsEdit = False Then
        ' Does this record exist?
        Dim colValues As Collection

        If TableAccess(m_clsInterestRates).RecordCount > 0 Then
            Set colValues = New Collection

            sSequenceNumber = m_Frm.txtInterestRates(SEQUENCE_NUMBER).Text
            colValues.Add sSequenceNumber
            TableAccess(m_clsInterestRates).SetKeyMatchValues colValues
            TableAccess(m_clsInterestRates).ApplyFilter

            If TableAccess(m_clsInterestRates).RecordCount > 0 Then
                m_Frm.txtInterestRates(SEQUENCE_NUMBER).SetFocus
                MsgBox "Interest Rate Sequence Number must be unique", vbCritical
                bRet = False
            End If

            TableAccess(m_clsInterestRates).CancelFilter
        End If

    End If

    ValidateScreenData = bRet

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SetupRateCombo()
    On Error GoTo Failed
    Dim rate As InterestRateTypes

    rate.sValue = TYPE_BASE
    m_Frm.cboType.AddItem rate.sValue
    rate.sSaveValue = TYPE_BASE_SAVE
    m_Rates(1) = rate

    rate.sValue = TYPE_FIXED
    m_Frm.cboType.AddItem rate.sValue
    rate.sSaveValue = TYPE_FIXED_SAVE
    m_Rates(2) = rate

    rate.sValue = TYPE_CAPPED
    m_Frm.cboType.AddItem rate.sValue
    rate.sSaveValue = TYPE_CAPPED_SAVE
    m_Rates(3) = rate

    rate.sValue = TYPE_DISCOUNT
    m_Frm.cboType.AddItem rate.sValue
    rate.sSaveValue = TYPE_DISCOUNT_SAVE
    m_Rates(4) = rate

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Another()
    On Error GoTo Failed

    SetAddState
    'Trip this flag to true.
    
    g_clsFormProcessing.ClearScreenFields m_Frm
    m_iRecordsAdded = m_iRecordsAdded + 1
    m_Frm.txtInterestRates(SEQUENCE_NUMBER).SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub OK()
    On Error GoTo Failed

    ' When returning, cancel the filter so the caller can see all records.
    TableAccess(m_clsInterestRates).CancelFilter
    SetReturnCode MSGSuccess
    m_Frm.Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(m_Frm)

    DoOKProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub HandleType()
    On Error GoTo Failed
    Dim sVal As String

    sVal = m_Frm.cboType.SelText

    If sVal = TYPE_BASE Or sVal = TYPE_FIXED Or sVal = TYPE_DISCOUNT Then
        m_Frm.txtInterestRates(FLOOR_RATE).Text = ""
        m_Frm.txtInterestRates(CEILING_RATE).Text = ""

        m_Frm.txtInterestRates(FLOOR_RATE).Enabled = False
        m_Frm.txtInterestRates(CEILING_RATE).Enabled = False

        m_Frm.lblInterestRates(FLOOR_RATE).Enabled = False
        m_Frm.lblInterestRates(CEILING_RATE).Enabled = False
    Else
        m_Frm.lblInterestRates(FLOOR_RATE).Enabled = True
        m_Frm.lblInterestRates(CEILING_RATE).Enabled = True

        m_Frm.txtInterestRates(FLOOR_RATE).Enabled = True
        m_Frm.txtInterestRates(CEILING_RATE).Enabled = True
    End If

    ' If Base, then interest rate isn't mandatory
    If sVal = TYPE_BASE Then
        m_Frm.txtInterestRates(INTEREST_RATE).Mandatory = False
    Else
        m_Frm.txtInterestRates(INTEREST_RATE).Mandatory = True
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Cancel()

    Dim uResponse As VbMsgBoxResult

    On Error GoTo Failed

    ' DJP SYS1942, if the user has added a record and pressed Another, ensure it is saved.
    If m_iRecordsAdded > 1 Then
        SetReturnCode
    End If

    m_Frm.Hide

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub SetTableClass(clsTable As TableAccess)
    Set m_clsInterestRates = clsTable
End Sub

Private Sub SetRateType()
    On Error GoTo Failed
    Dim sRate As String
    Dim vVal As Variant

    g_clsFormProcessing.HandleComboIndex m_Frm.cboType, vVal, GET_CONTROL_VALUE

    If Not IsEmpty(vVal) Then
        sRate = GetRateFromIndex(CInt(vVal) + 1)

        If Len(sRate) > 0 Then
            m_clsInterestRates.SetRateType (sRate)
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetRateType()
    On Error GoTo Failed
    Dim nIndex As Integer
    Dim vVal As Variant

    vVal = m_clsInterestRates.GetRateType()

    If Not IsEmpty(vVal) Then
        nIndex = GetIndexFromRate(CStr(vVal))

        If nIndex >= 0 Then
            nIndex = nIndex - 1
            g_clsFormProcessing.HandleComboIndex m_Frm.cboType, CVar(nIndex), SET_CONTROL_VALUE
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function GetRateFromIndex(nIndex As String) As String
    On Error GoTo Failed
    Dim rate As InterestRateTypes

    If nIndex <= TYPE_COUNT Then
        rate = m_Rates(nIndex)
        GetRateFromIndex = rate.sSaveValue
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetRateFromIndex - Index invalid"

    End If

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function GetIndexFromRate(sRate) As Integer
    On Error GoTo Failed
    Dim bFound As Boolean
    Dim nThisIndex As Integer
    Dim rate As InterestRateTypes

    nThisIndex = 1
    bFound = False

    While nThisIndex <= TYPE_COUNT And bFound = False
        rate = m_Rates(nThisIndex)

        If rate.sSaveValue = sRate Then
            bFound = True
        Else
            nThisIndex = nThisIndex + 1
        End If

    Wend

    If bFound = False Then
        nThisIndex = -1
    End If

    GetIndexFromRate = nThisIndex

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub UnloadForm()
    On Error GoTo Failed
    Unload m_Frm
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Function GetInterestRateTable() As MortProdIntRatesTable
    On Error GoTo Failed
    
    If m_clsInterestRates Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, "Interest Rate table class is empty"
    End If
    
    Set GetInterestRateTable = m_clsInterestRates
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub PopulateBaseRateSetCombo()
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed
    Dim clsBaseRateSet As New BaseRateSetTable
    Dim clsTableAccess As TableAccess
    
    clsBaseRateSet.GetBaseRateSet
    
    g_clsFormProcessing.PopulateDataCombo m_frm.cboBaseRateSet, clsBaseRateSet
    Set clsTableAccess = clsBaseRateSet
    
    If clsTableAccess.RecordCount = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "No Base Rate Sets exist"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetInterestRateTypeTable() As MortProdIntRatesTable
'GD        16/05/2002   BMIDS00009
    On Error GoTo Failed
    Set GetInterestRateTypeTable = m_clsInterestRates
    
    Exit Function
Failed:    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
