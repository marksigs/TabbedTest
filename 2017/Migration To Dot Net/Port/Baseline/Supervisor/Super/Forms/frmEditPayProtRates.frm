VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPayProtRates 
   Caption         =   "Add/Edit Payment Protection Rates"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   Icon            =   "frmEditPayProtRates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8595
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Payment Protection Rates"
      Height          =   4215
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgPayProt 
         Height          =   3855
         Left            =   300
         TabIndex        =   7
         Top             =   240
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   6800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7230
      TabIndex        =   2
      Top             =   5175
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5175
      Width           =   1230
   End
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGEditBox txtPayProt 
      Height          =   315
      HelpContextID   =   2
      Index           =   1
      Left            =   4320
      TabIndex        =   0
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   10
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   3300
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Rate Number"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditPayProtRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditPayProtRates
' Description   :   Form which allows the user to edit and add Payment Protection rates.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Changed call
'                       to NextNumber as it has been moved from FormProcessing to DataAccess.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const BANDED_FEE_SET = 0
Private Const BANDED_DATE = 1

' Private data
Private m_sOriginalSet As String
Private m_sOriginalStartDate As String
Private m_vNextNumber As Variant
Private m_bIsEdit As Boolean

Private m_clsPayProtRates As PayProtRatesTable
Private m_clsPayProtRatesSet As PayProtRatesSetTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bFailed As Boolean
    Dim bShowError As Boolean
    Dim clsTableAccess As TableAccess

    
    bFailed = False
    bShowError = True

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)

    If bRet = True Then
        bRet = dgPayProt.ValidateRows()
    End If

    If bRet = True Then
        ValidateScreenData
        SaveScreenData
        DoUpdates
        DoSetUpdates
        
        ' Update the payment protection rates table
        Set clsTableAccess = m_clsPayProtRates
        clsTableAccess.Update
        SaveChangeRequest
        SetReturnCode
        
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub ValidateScreenData()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsPayProtRates)
    
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Duplicate records found - please enter a unique combination"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    Dim sSetNumber As String
    Dim sStartDate As String
    
    sSetNumber = Me.txtPayProt(BANDED_FEE_SET).Text
    sStartDate = Me.txtPayProt(BANDED_DATE).Text
    
    Set clsTableAccess = m_clsPayProtRates
    Set colMatchValues = New Collection
    colMatchValues.Add sSetNumber
    colMatchValues.Add sStartDate
    clsTableAccess.SetKeyMatchValues colMatchValues

    g_clsHandleUpdates.SaveChangeRequest m_clsPayProtRates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub dgPayProt_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgPayProt.ValidateRow(nCurrentRow)
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    ' Initialise Form

    Set m_clsPayProtRates = New PayProtRatesTable
    Set m_clsPayProtRatesSet = New PayProtRatesSetTable
    
    PopulateRates
    
    If m_bIsEdit = True Then
        PopulateScreen
        SetEditState
    Else
        SetAddState
    End If
    
    m_sOriginalSet = txtPayProt(BANDED_FEE_SET).Text
    m_sOriginalStartDate = txtPayProt(BANDED_DATE).Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Public Sub PopulateScreen()
    On Error GoTo Failed

    PopulateScreenFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetAddState()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsPayProtRates
    m_vNextNumber = m_clsPayProtRates.NextRate
    
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        txtPayProt(BANDED_FEE_SET).Text = m_vNextNumber
    Else
        g_clsErrorHandling.RaiseError errGeneralError, " Cannot load form - Rate Number is empty"
    End If
    SetGridFields
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetEditState()
    On Error GoTo Failed
    
    txtPayProt(BANDED_FEE_SET).Enabled = False
    SetDataGridState
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateRates()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    Dim rs As ADODB.Recordset

    Set clsTableAccess = m_clsPayProtRates

    If m_bIsEdit = True Then
        clsTableAccess.SetKeyMatchValues m_colKeys
        Set rs = clsTableAccess.GetTableData()
    Else
        Set rs = clsTableAccess.GetTableData(POPULATE_EMPTY)
    End If
    If Not rs Is Nothing Then
        Set dgPayProt.DataSource = rs
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, "Payment Protection Rates"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetDistributionChannels(colText As Collection, colValues As Collection)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim colFields As New Collection
    Dim clsTableAccess As TableAccess
    Dim clsDistChannel As New DistributionChannelTable
    
    Set colFields = clsDistChannel.GetComboFields()
        
    If colFields.Count > 0 Then
        Set clsTableAccess = clsDistChannel
        Set rs = clsTableAccess.GetTableData(POPULATE_ALL)
        
        ValidateRecordset rs, "Distribution Channels "
        g_clsFormProcessing.PopulateCollectionFromTable clsTableAccess, colFields, colText, colValues
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetGridFields()
    On Error GoTo Failed

    Dim bRet As Boolean
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    Dim clsCombo As New ComboUtils

    bRet = True
    
    ' First, Fee Set. Not visible, but needs the Fee Set copied in
    fields.sField = "PaymentProtectionRatesNumber"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtPayProt(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = ""
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    colFields.Add fields

    ' StartDate not visible, but has to be copied in.
    fields.sField = "PPStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtPayProt(BANDED_DATE).Text
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Next, Channel Text has to be entered
    fields.sField = "ChannelIDText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Channel must be entered"
    fields.sTitle = "Channel Name"
    fields.sOtherField = "ChannelID"

    GetDistributionChannels colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS

    colFields.Add fields

    ' ChannelID
    fields.sField = "ChannelID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Gender Text
    fields.sField = "ApplicantsGenderText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Applicants' Gender must be entered"
    fields.sTitle = "Applicants' Gender"
    fields.sOtherField = "ApplicantsGender"

    clsCombo.FindComboGroup "LifeCoverGender", colComboValues, colComboIDS

    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS

    colFields.Add fields

    ' Gender
    fields.sField = "ApplicantsGender"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' Next, High Applicants age
    fields.sField = "HighApplicantsAge"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Applicants Age must be entered"
    fields.sTitle = "Maximum Age"
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    
    colFields.Add fields

    ' ASU Rate
    fields.sField = "ASURate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "ASU Rate must be entered"
    fields.sTitle = "ASU Rate"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' AS Rate
    fields.sField = "ASRate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "AS Rate must be entered"
    fields.sTitle = "AS Rate"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    ' U Rate
    fields.sField = "URate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "U Rate must be entered"
    fields.sTitle = "U Rate"
    Set fields.colComboValues = Nothing
    Set fields.colComboIDS = Nothing
    fields.sOtherField = ""
    colFields.Add fields

    dgPayProt.SetColumns colFields, "EditPayProtRates", "Payment Protection"
    'dgPayProt.SetWidths
    Exit Sub
Failed:
    Err.Raise Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    m_vNextNumber = m_clsPayProtRates.GetRateNumber()
    
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        txtPayProt(BANDED_FEE_SET).Text = m_vNextNumber
        txtPayProt(BANDED_DATE).Text = m_clsPayProtRates.GetStartDate()
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub txtPayProt_LostFocus(Index As Integer)
    If m_bIsEdit = False Then
        SetDataGridState
    End If
End Sub
Private Sub txtPayProt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtPayProt(Index).ValidateData()
End Sub
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)

    If bRet = True Then
        Dim bEnabled As Boolean

        bEnabled = dgPayProt.Enabled
        dgPayProt.Enabled = True

        SetGridFields

        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtPayProt(BANDED_FEE_SET).Text
                m_sOriginalStartDate = txtPayProt(BANDED_DATE).Text

                dgPayProt.AddRow
                dgPayProt.SetFocus
            End If
        End If
    End If
End Sub
Private Function DoesRecordExist() As Boolean
    Dim bRet As Boolean
    Dim sSet As String
    Dim sStartDate As String
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    
    sSet = txtPayProt(BANDED_FEE_SET).Text
    sStartDate = txtPayProt(BANDED_DATE).Text

    col.Add sSet
    col.Add sStartDate
    
    Set clsTableAccess = m_clsPayProtRates
    bRet = clsTableAccess.DoesRecordExist(col)

    If bRet = True Then
        MsgBox "Name and Start Date already exist - please enter a unique combination"
        txtPayProt(BANDED_FEE_SET).SetFocus
    End If
    DoesRecordExist = bRet
End Function
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    Dim colValues As New Collection
    Dim sSetID As String
    Dim sStartDate As String

    Set clsBandedTable = m_clsPayProtRates

    sStartDate = txtPayProt(BANDED_DATE).Text
    sSetID = txtPayProt(BANDED_FEE_SET).Text

    If Len(sStartDate) > 0 And Len(sSetID) > 0 Then
        If sStartDate <> m_sOriginalStartDate Or sSetID <> m_sOriginalSet Then
            colValues.Add sSetID
            colValues.Add sStartDate

            clsBandedTable.SetUpdateValues colValues

            ' We only want to check that the record exists if the name or start date has changed
            clsBandedTable.SetUpdateSets
            clsBandedTable.DoUpdateSets
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "PaymentProtectionRates - Start Date and End Date must be valid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' Need to see if the SetID the user has typed in exists already in the ValuationFeeSet table.
' If it does, do nothing, because there will already be entries for it in the ValuationFee table.
' If it doesn't exist, we need to put an entry into the ValuationFeeSet table first.
Private Sub DoSetUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess

    Dim col As New Collection

    sSet = txtPayProt(BANDED_FEE_SET).Text

    If Len(sSet) > 0 Then
        Set clsTableAccess = m_clsPayProtRatesSet

        col.Add sSet
        clsTableAccess.SetKeyMatchValues col

        bRet = clsTableAccess.DoesRecordExist(col)

        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess

            m_clsPayProtRatesSet.SetRateNumber sSet
            clsTableAccess.Update
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveScreenData()
    On Error GoTo Failed
    
    If Not IsNull(m_vNextNumber) And Not IsEmpty(m_vNextNumber) Then
        m_clsPayProtRates.SetRateNumber m_vNextNumber
    Else
        g_clsErrorHandling.RaiseError errGeneralError, " Cannot save data - Rate Number is empty"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
