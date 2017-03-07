VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditTransferOfEquityFees 
   Caption         =   "Transfer of Equity Fee Set Add/Edit"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   Icon            =   "frmEditTransferOfEquityFees.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8475
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   5145
      Width           =   1215
   End
   Begin VB.Frame frameFeeSetBands 
      Caption         =   "Fee Set Bands"
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   8235
      Begin MSGOCX.MSGDataGrid dgTransferOfEquityFees 
         Height          =   3855
         Left            =   600
         TabIndex        =   2
         Top             =   300
         Width           =   6915
         _ExtentX        =   12197
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
   Begin VB.CommandButton cmdAnother 
      Caption         =   "A&nother"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtTransferOfEquityFees 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
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
   Begin MSGOCX.MSGEditBox txtTransferOfEquityFees 
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   120
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
   Begin VB.Label lblAdminFees 
      Caption         =   "Start Date"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label lblAdminFees 
      Caption         =   "Set Number"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditTransferOfEquityFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditTransferOfEquityFees
' Description   :   Form which allows the user to edit Tranfer Of Equity Fees. This is a banded
'                   form, so has a datagrid to enter the details.
' Change history
' Prog      Date        Description
' TW        11/12/06    EP2_20 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Text indexes
Private Const BANDED_FEE_SET = 0
Private Const BANDED_DATE = 1

' Private data
Private m_sOriginalSet As String
Private m_sOriginalStartDate As String
Private m_bIsEdit As Boolean

Private m_clsTableAccess As TableAccess
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

Private Sub Form_Initialize()
    Set m_colKeys = New Collection
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    Set m_clsTableAccess = New TransferOfEquityFeeTable
    
    dgTransferOfEquityFees.Enabled = False
      
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    m_sOriginalSet = txtTransferOfEquityFees(BANDED_FEE_SET).Text
    m_sOriginalStartDate = txtTransferOfEquityFees(BANDED_DATE).Text
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Function DoesRecordExist() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim sStartDate As String
    
    Dim col As New Collection
    
    sSet = txtTransferOfEquityFees(BANDED_FEE_SET).Text
    sStartDate = txtTransferOfEquityFees(BANDED_DATE).Text
    
    If Len(sSet) > 0 And Len(sStartDate) > 0 Then
        col.Add sSet
        col.Add sStartDate
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        bRet = m_clsTableAccess.DoesRecordExist(col)
        
        If bRet = True Then
            MsgBox "Set number and Start Date already exist - please enter a unique combination", vbCritical
            txtTransferOfEquityFees(BANDED_FEE_SET).SetFocus
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Set and Start date must be valid", vbCritical
    End If
    
    DoesRecordExist = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function

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

Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sSet As String
    Dim sDate As String
    
    Set colMatchValues = New Collection
    sSet = Me.txtTransferOfEquityFees(BANDED_FEE_SET).Text
    sDate = Me.txtTransferOfEquityFees(BANDED_DATE).Text
    colMatchValues.Add sSet
    colMatchValues.Add sDate
    
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = ValidateScreenData()
    
    If bRet = True Then
        'g_clsDataAccess.BeginTrans
        ' Do the banded updates, if there are any
        DoUpdates
        ' Update the set table if needed
        DoSetUpdates
        ' Update the TransferOfEquityFeeset table.
        m_clsTableAccess.Update
        
        SaveChangeRequest
        SetReturnCode
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    'g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    Dim colValues As New Collection
    Dim sSetID As String
    Dim sStartDate As String
    Dim bRet As Boolean
    
    ' Use the banded interface
    Set clsBandedTable = m_clsTableAccess
    bRet = False
    
    sStartDate = txtTransferOfEquityFees(BANDED_DATE).Text
    sSetID = txtTransferOfEquityFees(BANDED_FEE_SET).Text
    
    If Len(sStartDate) > 0 And Len(sSetID) > 0 Then
        If sStartDate <> m_sOriginalStartDate Or sSetID <> m_sOriginalSet Then
            colValues.Add sSetID
            colValues.Add sStartDate
        
            clsBandedTable.SetUpdateValues colValues
            clsBandedTable.SetUpdateSets
            ' Do the set updates with the new values
            clsBandedTable.DoUpdateSets
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "TransferOfEquityFees - StartDate and SetID must be valid"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub DoSetUpdates()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sSet As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim clsTransferOfEquityFeeset As TransferOfEquityFeeSetTable
    
    Dim col As New Collection
    
    sSet = txtTransferOfEquityFees(BANDED_FEE_SET).Text

    If Len(sSet) > 0 Then
        Set clsTransferOfEquityFeeset = New TransferOfEquityFeeSetTable
        Set clsTableAccess = clsTransferOfEquityFeeset
        
        col.Add sSet
        
        If g_clsVersion.DoesVersioningExist() Then
            col.Add g_sVersionNumber
        End If
        
        clsTableAccess.SetKeyMatchValues col
        
        bRet = clsTableAccess.DoesRecordExist(col)
    
        If bRet = False Then
            ' Doesn't exist, so add a new one
            g_clsFormProcessing.CreateNewRecord clsTableAccess
        
            clsTransferOfEquityFeeset.SetFeeSet sSet
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub dgTransferOfEquityFees_BeforeAdd(bCancel As Boolean, nCurrentRow As Integer)
    If nCurrentRow >= 0 Then
        bCancel = Not dgTransferOfEquityFees.ValidateRow(nCurrentRow)
    End If
End Sub

Private Sub SetupDataGrid()
    On Error GoTo Failed
    Dim colTables As New Collection
    Dim colDataControls As New Collection
    
    colTables.Add m_clsTableAccess
    colDataControls.Add dgTransferOfEquityFees
    
    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If

    If m_bIsEdit = True Then
        If m_clsTableAccess.RecordCount > 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate current record"
        End If
    End If

    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAddState()
    SetupDataGrid
    
    If txtTransferOfEquityFees(BANDED_FEE_SET).Visible = True Then
        txtTransferOfEquityFees(BANDED_FEE_SET).SetFocus
    End If
End Sub

Public Sub SetEditState()
    On Error GoTo Failed
    
    m_clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDataGrid
    
    txtTransferOfEquityFees(BANDED_FEE_SET).Enabled = False
    txtTransferOfEquityFees(BANDED_DATE).Enabled = False
    
    dgTransferOfEquityFees.Enabled = True
    cmdAnother.Enabled = False
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
    Dim clsTransferOfEquityFeeTable As TransferOfEquityFeeTable
    Dim sVersionField As String
    Set clsTransferOfEquityFeeTable = m_clsTableAccess
    
    ' First, Fee Set. Not visible, but needs the Fee Set copied in
    fields.sField = "TransferOfEquityFeeset"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtTransferOfEquityFees(BANDED_FEE_SET).Text
    fields.sError = ""
    fields.sTitle = "Fee Set"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = clsTransferOfEquityFeeTable.GetStartDateField() '"StartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtTransferOfEquityFees(BANDED_DATE).Text
    fields.sError = ""
    fields.sTitle = "Start Date"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Nature Of Loan
    fields.sField = "NatureOfLoan"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = "Nature Of Loan must be entered"
    fields.sTitle = "Nature Of Loan Value"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    ' Nature Of Loan Description
    fields.sField = "NatureOfLoanDescription"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Nature Of Loan Description must be entered"
    fields.sTitle = "Nature Of Loan"
    fields.sOtherField = "NatureOfLoan"
    clsCombo.FindComboGroup "NatureOfLoan", colComboValues, colComboIDS
    
    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS
    
    colFields.Add fields
    
    Set fields.colComboValues = Nothing
    
    ' Amount
    fields.sField = "Amount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Fee Amount must be entered"
    fields.sTitle = "Fee Amount"
    Set fields.colComboValues = Nothing
    fields.sOtherField = ""
    colFields.Add fields
    
    If g_clsVersion.DoesVersioningExist() Then
        ' Versioning...
        fields.sField = clsTransferOfEquityFeeTable.GetVersionField()
        fields.bRequired = True
        fields.bVisible = False
        fields.sDefault = g_sVersionNumber
        fields.sError = ""
        fields.sTitle = ""
        
        colFields.Add fields
    End If
    
    dgTransferOfEquityFees.SetColumns colFields, "EditTransferOfEquityFees", "Transfer of Equity Fees"
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim clsTransferOfEquityFees As TransferOfEquityFeeTable
    
    Set clsTransferOfEquityFees = m_clsTableAccess
    txtTransferOfEquityFees(BANDED_FEE_SET).Text = clsTransferOfEquityFees.GetFeeSet()
    txtTransferOfEquityFees(BANDED_DATE).Text = clsTransferOfEquityFees.GetStartDate()
    
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
    If bRet And m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    If bRet Then
        bRet = dgTransferOfEquityFees.ValidateRows()
    End If

    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsTableAccess)
    
        If bRet = False Then
            g_clsErrorHandling.RaiseError errGeneralError, "Set number and Start Date must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub txtTransferOfEquityFees_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtTransferOfEquityFees(Index).ValidateData()

    If Cancel = False Then
        SetDataGridState
    End If
End Sub

Public Sub SetDataGridState()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    ' Make sure all data that needs to be entered, has been
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgTransferOfEquityFees.Enabled
        dgTransferOfEquityFees.Enabled = True
        
        SetGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                m_sOriginalSet = txtTransferOfEquityFees(BANDED_FEE_SET).Text
                m_sOriginalStartDate = txtTransferOfEquityFees(BANDED_DATE).Text

                dgTransferOfEquityFees.AddRow
                dgTransferOfEquityFees.SetFocus
            End If
        End If
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


