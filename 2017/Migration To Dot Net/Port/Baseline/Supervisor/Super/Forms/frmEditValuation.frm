VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditValuation 
   Caption         =   "Valuation Fee Set Add/Edit"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   Icon            =   "frmEditValuation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "A&nother"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Fee Set Bands"
      Height          =   4215
      Left            =   300
      TabIndex        =   5
      Top             =   780
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgValuationFees 
         Height          =   3855
         Left            =   660
         TabIndex        =   2
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
   Begin MSGOCX.MSGEditBox txtValuationFees 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   0
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      TextType        =   6
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
   End
   Begin MSGOCX.MSGEditBox txtValuationFees 
      Height          =   315
      Index           =   1
      Left            =   5100
      TabIndex        =   1
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
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
   Begin VB.Label lblValuationFees 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   7
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label lblValuationFees 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditValuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditValuation
' Description   :

' Change history
' Prog      Date        Description
' DJP       09/11/00    Created for Phase 2 Task Management
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   BMIDS
'
'   AW      16/05/02    BM023 - Added 3 extra fields and ValidateFeeAmount()
'   MO      10/06/02    BMIDS00040 - Changed validation for maximum and minimum valuation amounts
'   GD      02/07/02    BMIDS00083 - Changed constants so that ValidateFeeAmount picks out correct columns
'
'   INGDUK
'   JD      03/09/2005  MAR40 added type of application


Option Explicit

Private Const BANDED_FEE_SET = 0
Private Const BANDED_DATE = 1

Private Const AMOUNT_COL As Integer = 5
'GD 02/07/02 Original - Start
'Private Const MINIMUM_COL As Integer = 8
'Private Const MAXIMUM_COL As Integer = 9
'Private Const PERCENT_COL As Integer = 10
'GD 02/07/02 Original - End
'GD 02/07/02 New for BMIDS00083 - Start
Private Const MINIMUM_COL As Integer = 9
Private Const MAXIMUM_COL As Integer = 10
Private Const PERCENT_COL As Integer = 8
'GD 02/07/02 New for BMIDS00083 - End

Private m_sOriginalSet As String
Private m_sOriginalStartDate As String

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
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

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        g_clsFormProcessing.ClearScreenFields Me
        Form_Load
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
'Public Sub SetTableClass(clsTable As TableAccess)
'    Set m_clsTableAccess = clsTable
'End Sub
Private Function DoesRecordExist() As Boolean
    Dim bRet As Boolean
    Dim sSet As String
    Dim sStartDate As String
    
    Dim col As New Collection
    
    sSet = txtValuationFees(BANDED_FEE_SET).Text
    sStartDate = txtValuationFees(BANDED_DATE).Text
    
    col.Add sSet
    col.Add sStartDate
    
    If g_clsVersion.DoesVersioningExist() Then
        col.Add g_sVersionNumber
    End If
    
    bRet = m_clsTableAccess.DoesRecordExist(col)
    
    If bRet = True Then
        MsgBox "Name and Start Date already exist - please enter a unique combination", vbCritical
        txtValuationFees(BANDED_FEE_SET).SetFocus
    End If
    DoesRecordExist = bRet
End Function
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()
    
    If bRet = True Then
        DoUpdates
        DoSetUpdates
        m_clsTableAccess.Update
        SaveChangeRequest
        SetReturnCode
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sSet As String
    Dim sDate As String
    
    Set colMatchValues = New Collection
    sSet = txtValuationFees(BANDED_FEE_SET).Text
    sDate = txtValuationFees(BANDED_DATE).Text
    
    colMatchValues.Add sSet
    colMatchValues.Add sDate
    
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = DoOKProcessing()
    
    If bRet = True Then
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    Dim colValues As New Collection
    Dim sSetID As String
    Dim sStartDate As String
    Dim bRet As Boolean
    
    Set clsBandedTable = m_clsTableAccess
    bRet = False
    
    sStartDate = txtValuationFees(BANDED_DATE).Text
    sSetID = txtValuationFees(BANDED_FEE_SET).Text
    
    If Len(sStartDate) > 0 And Len(sSetID) > 0 Then
        If sStartDate <> m_sOriginalStartDate Or sSetID <> m_sOriginalSet Then
            colValues.Add sSetID
            colValues.Add sStartDate
        
            clsBandedTable.SetUpdateValues colValues
        
            ' We only want to check that the record exists if the name or start date has changed
            clsBandedTable.SetUpdateSets
                
            clsBandedTable.DoUpdateSets
        End If
    
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
    Dim clsValuationFeeSet As ValuationFeeSetTable
    
    Dim col As New Collection
    
    sSet = txtValuationFees(BANDED_FEE_SET).Text

    If Len(sSet) > 0 Then
        Set clsValuationFeeSet = New ValuationFeeSetTable
        Set clsTableAccess = clsValuationFeeSet
        
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        
        bRet = clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess
        
            clsValuationFeeSet.SetFeeSet sSet
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub dgValuationFees_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgValuationFees.ValidateRow(nCurrentRow)
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    ' Initialise Form
    SetReturnCode MSGFailure
    Set m_clsTableAccess = New ValuationFeeTable
    dgValuationFees.Enabled = False
    
    If m_bIsEdit = True Then
        SetEditState
        PopulateScreen
    Else
        SetAddState
    End If
    
    m_sOriginalSet = txtValuationFees(BANDED_FEE_SET).Text
    m_sOriginalStartDate = txtValuationFees(BANDED_DATE).Text
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub SetupDBControl()
    On Error GoTo Failed
    Dim colTables As New Collection
    Dim colDataControls As New Collection
    
    colTables.Add m_clsTableAccess
    colDataControls.Add dgValuationFees
    
    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If
    
    If m_bIsEdit = True Then
        PopulateScreenFields
    End If

    SetGridFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetAddState()
    SetupDBControl
    
    If txtValuationFees(BANDED_FEE_SET).Visible = True Then
        txtValuationFees(BANDED_FEE_SET).SetFocus
    End If
End Sub
Public Sub SetEditState()
    On Error GoTo Failed
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    dgValuationFees.Enabled = True
    txtValuationFees(BANDED_FEE_SET).Enabled = False
    txtValuationFees(BANDED_DATE).Enabled = False
    cmdAnother.Enabled = False
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Public Sub PopulateScreen()
    On Error GoTo Failed
    
    SetupDBControl
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    Dim clsCombo As New ComboUtils
    Dim clsValuationFeeTable As ValuationFeeTable
    
    Set clsValuationFeeTable = New ValuationFeeTable
    ' First, Fee Set. Not visible, but needs the Fee Set copied in
    fields.sField = "ValuationFeeSet"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtValuationFees(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = ""
    fields.sOtherField = ""
    Set fields.colComboValues = Nothing
    colFields.Add fields

    ' StartDate not visible, but has to be copied in.
    fields.sField = "ValuationFeeStartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtValuationFees(BANDED_DATE).Text
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Next, TypeOfValuation Text has to be entered
    fields.sField = "TypeOfValuationText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Type of Valuation must be entered"
    fields.sTitle = "Valuation Type"
    fields.sOtherField = "TypeOfValuation"
    clsCombo.FindComboGroup "ValuationType", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    colFields.Add fields
    
    ' Location Text
    fields.sField = "LocationText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Location must be entered"
    fields.sTitle = "Location"
    fields.sOtherField = "Location"
    
    clsCombo.FindComboGroup "PropertyLocation", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    colFields.Add fields
    
    ' Next, TypeOfValuation
    fields.sField = "TypeOfValuation"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = "Type of Valuation must be entered"
    fields.sTitle = ""
    fields.sOtherField = ""
    colFields.Add fields


    ' Location
    fields.sField = "MaximumValue"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Maximum Value must be entered"
    fields.sTitle = "Maximum Value"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""

    colFields.Add fields

    ' Location
    fields.sField = "Location"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = ""
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""

    colFields.Add fields

    ' Amount
    fields.sField = "Amount"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Amount must be entered"
    fields.sTitle = "Amount"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""

    colFields.Add fields
    
    ' Fee Percentage
    fields.sField = "FeePercentage"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Fee Percentage"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""

    colFields.Add fields
    
    ' Min Fee Value
    fields.sField = "MinimumFeeValue"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Min Fee Value"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = 0
    colFields.Add fields
    
    ' Max Fee Value
    fields.sField = "MaximumFeeValue"
    fields.bRequired = False
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Max Fee Value"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    fields.sMinValue = 0
    colFields.Add fields
    
    'JD MAR40 added typeofapplication and TypeOfApplicationtext
    ' TypeOfApplicationText
    fields.sField = "TypeOfApplicationText"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Type of Application must be entered"
    fields.sTitle = "Type Of Application"
    fields.sOtherField = "TypeOfApplication"
    clsCombo.FindComboGroup "TypeOfMortgage", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    colFields.Add fields
    
    ' TypeOfApplication
    fields.sField = "TypeOfApplication"
    fields.bRequired = False
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "Type of application text"
    fields.sOtherField = ""
    fields.sMinValue = 0
    Set fields.colComboValues = Nothing
    colFields.Add fields
    
    ' Versioning...
    If g_clsVersion.DoesVersioningExist() Then
        fields.sField = clsValuationFeeTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        
        colFields.Add fields
    End If
    
    dgValuationFees.SetColumns colFields, "EditValuationFees", "Valuation Fees"
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim clsValuationFees As ValuationFeeTable
    
    Set clsValuationFees = m_clsTableAccess
    txtValuationFees(BANDED_FEE_SET).Text = clsValuationFees.GetFeeSet()
    txtValuationFees(BANDED_DATE).Text = clsValuationFees.GetStartDate()
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    ' If we're adding, check that this record doesn't exist already
    If bRet = True And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet = True Then
        bRet = dgValuationFees.ValidateRows()
    End If
    
    If bRet = True Then
        bRet = ValidateFeeAmount()
    End If
    
    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Type of Valuation, Maximum Value and Location must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub txtValuationFees_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtValuationFees(Index).ValidateData()

    If Cancel = False And m_bIsEdit = False Then
        SetDataGridState
    End If

End Sub
Public Sub SetDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgValuationFees.Enabled
        dgValuationFees.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtValuationFees(BANDED_FEE_SET).Text
                m_sOriginalStartDate = txtValuationFees(BANDED_DATE).Text
                
                dgValuationFees.AddRow
                dgValuationFees.SetFocus
            End If
        End If
    End If
End Sub


Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   AW      16/05/02    BM023
'   MO      10/06/02    BMIDS00040
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateFeeAmount() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    Dim nRowIndex As Integer
    Dim strFee As String, strMin As String, strMax As String, strFeePercent As String

    ValidateFeeAmount = False
    
    'Iterate through all the rows of the grid control.
    For nRowIndex = 0 To (dgValuationFees.Rows - 1)

        strFee = dgValuationFees.GetAtRowCol(nRowIndex, AMOUNT_COL)
        strMin = dgValuationFees.GetAtRowCol(nRowIndex, MINIMUM_COL)
        strMax = dgValuationFees.GetAtRowCol(nRowIndex, MAXIMUM_COL)
        strFeePercent = dgValuationFees.GetAtRowCol(nRowIndex, PERCENT_COL)

        If (Len(strFee) > 0 And Len(strFeePercent) > 0) Or _
            (Len(strFee) = 0 And Len(strFeePercent) = 0) Then

            'Place the cursor in the invalid field.
            dgValuationFees.col = AMOUNT_COL
            dgValuationFees.Row = nRowIndex
            dgValuationFees.SetGridFocus

            g_clsErrorHandling.RaiseError errGeneralError, "Either an amount OR percentage must be entered"

        End If

        If Len(strFeePercent) = 0 And ((Val(strMin) > 0) Or (Val(strMax) > 0)) Then

            'Place the cursor in the invalid field.
            dgValuationFees.col = MINIMUM_COL
            dgValuationFees.Row = nRowIndex
            dgValuationFees.SetGridFocus

            g_clsErrorHandling.RaiseError errGeneralError, "Minimum and maximum fee values can only be applied against a fee percentage."
        End If
        
        If (Val(strMax) < Val(strMin)) And (Len(strMax) > 0) Then

            'Place the cursor in the invalid field.
            dgValuationFees.col = MAXIMUM_COL
            dgValuationFees.Row = nRowIndex
            dgValuationFees.SetGridFocus

            g_clsErrorHandling.RaiseError errGeneralError, "Maximum fee must be greater than or equal to the minimum fee."
        End If
    
    Next nRowIndex
       
    ValidateFeeAmount = True
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
