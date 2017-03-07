VERSION 5.00
Object = "{8FBFAD4D-5ED6-467A-98A5-FAA33BE4B270}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBatch 
   Caption         =   "Batch Scheduler"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   Icon            =   "frmEditBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetParameters 
      Caption         =   "&Set Parameters"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   7260
      TabIndex        =   8
      Top             =   3660
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3660
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Frame fraBatchScheduler 
      Height          =   3375
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   8715
      Begin MSGOCX.MSGDataCombo cboProgramType 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mandatory       =   -1  'True
      End
      Begin MSGOCX.MSGEditBox txtBatchNumber 
         Height          =   315
         Left            =   6480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Enabled         =   0   'False
         BackColor       =   12632256
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
      Begin MSGOCX.MSGEditBox txtDescription 
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   1200
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BackColor       =   16777215
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   50
      End
      Begin MSGOCX.MSGEditBox txtRunDate 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   1860
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSGOCX.MSGEditBox txtRunTime 
         Height          =   315
         Left            =   3900
         TabIndex        =   4
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Mask            =   "##:##"
         TextType        =   8
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
         MaxLength       =   5
      End
      Begin MSGOCX.MSGComboBox cboFrequency 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   2520
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Mandatory       =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblRunTime 
         Caption         =   "Run Time"
         Height          =   255
         Left            =   3060
         TabIndex        =   15
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lblBatchNumber 
         Caption         =   "Batch Number"
         Height          =   435
         Left            =   5340
         TabIndex        =   14
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label lblFrequency 
         Caption         =   "Frequency"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblRunDate 
         Caption         =   "Run Date"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label lblProgramType 
         Caption         =   "Program Type"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmEditBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditBatch
' Description   : Form which allows the adding/editing of batchs
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added Form
' SA        18/01/02    SYS3327
' SA        15/02/02    SYS4072 Descope to only allow Ad hoc runs
' GHun      26/04/02    SYS1621 Return from SetParameter incorrectly enables BatchNumber
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
' SA        10/06/2002  SYS4499 Add functionality to run scheduled batch jobs back in
' JD        13/12/2004  BM0092 don't get the new batchnumber until Save
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private m_ReturnCode As MSGReturnCode
Private m_clsPaymentProc As PaymentProcessingTable
Private m_clsBatch As BatchTable
Private m_sBatchNumber As String
Private m_bIsEdit As Boolean
Private m_clsComboValidation As ComboValidationTable
Private m_colKeys As Collection

Private m_bSavePaymentProcParams As Boolean
Private m_bIsView As Boolean

'ComboGroup Constants

Private Const COMBO_BATCH_STATUS As String = "BatchStatus"
Private Const COMBO_PROGRAM_TYPE As String = "BatchProgramType"
Private Const COMBO_BATCH_FREQUENCY As String = "BatchFrequency"
'Private Const BATCH_FREQ_ADHOC As String = "0"  'SYS4072 'SYS4499 Add functionality to run scheduled batches back in

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
'SYS3327 Can enter in View mode
Public Sub SetIsView(Optional bIsView As Boolean = True)
    m_bIsEdit = True
    m_bIsView = bIsView
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Private Sub cboProgramType_Change()
    
    Dim sValidation As String
    On Error GoTo Failed
    
    sValidation = m_clsComboValidation.GetValidationTypeFromCombo(CLng(cboProgramType.SelectedItem))
    
    Select Case sValidation
        Case PAYMENT_PROCESSING_TYPE
            cmdSetParameters.Enabled = True
        Case Else
            cmdSetParameters.Enabled = False
    End Select
    'SYS3327 Generate Batch number
    'Pass in Valueid of program type selected
    'BM0092 don't get the new batchnumber until Save
    'If txtBatchNumber.Text = "" Or m_bIsEdit = False Then
    '    txtBatchNumber.Text = m_clsBatch.GetNextBatchNumber(sValidation)
    'End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError "Batch Scheduler - no records to edit"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdSetParameters_Click
' Description   : Checks if a batch Number has been entered, and calls the Setparameters screen
'                 BatchNumber is disabled if OK is pressed on the frmEditPaymentProcessing form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSetParameters_Click()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim sBatchNumber As String
    Dim bRet As Boolean
    Set colValues = New Collection
    
    'Show the payment processing form, where parameters are set.
    
    sBatchNumber = txtBatchNumber.Text
    'JD BM0092 Batchnumber now created on save.
    'If Len(sBatchNumber) > 0 Then
        'SYS3327 Batch Number now automatically generated. Validation not necessary
        'If Not m_bIsEdit Then
        '    bRet = ValidateBatchNumber
        'Else
    '        bRet = True
        'End If
    'Else
    '    g_clsErrorHandling.DisplayError "A unique Batch Number must be entered before you can set the parameters"
    '    txtBatchNumber.SetFocus
    '    bRet = False
    'End If
    
    'If bRet Then
        colValues.Add sBatchNumber
        frmEditPaymentProcessing.SetTableClass m_clsPaymentProc
        frmEditPaymentProcessing.SetKeys colValues
        frmEditPaymentProcessing.SetIsEdit
        frmEditPaymentProcessing.SetIsView m_bIsView    'SYS3327
        frmEditPaymentProcessing.Show vbModal, Me
        m_bSavePaymentProcParams = m_clsPaymentProc.GetIsUpdated
        
        'SYS1621 txtBatchNumber should not be enabled
        'If m_bSavePaymentProcParams Then
        '    txtBatchNumber.Enabled = False
        'Else
        '    txtBatchNumber.Enabled = True
        'End If
        
    'End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub cmdCancel_Click()
    SetIsView False
    Hide
End Sub
Private Sub Form_Load()

    On Error GoTo Failed
    
    Set m_clsBatch = New BatchTable
    Set m_clsComboValidation = New ComboValidationTable
    Set m_clsPaymentProc = New PaymentProcessingTable
    
    PopulateScreenControls

    
    If m_bIsEdit Then
        SetFormEditState
        m_sBatchNumber = m_colKeys(1)
    Else
        SetFormAddState
    End If
    
    
    'm_bScreenUpdated = False
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub

Private Sub SetFormEditState()

    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sDepartmentID As String
    Dim colValues As New Collection

    Set clsTableAccess = m_clsBatch
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.DisplayError "Batch Scheduler - no records to edit"
        End If
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetFormAddState()
    On Error GoTo Failed
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate Combos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim clsBatch As BatchTable
    Dim colValues As New Collection
    Dim colIDS As New Collection
    Dim sField As String
    
    Set clsBatch = New BatchTable
    
    
    'Populate Frequency Combo
    g_clsFormProcessing.PopulateCombo COMBO_BATCH_FREQUENCY, cboFrequency
    'SYS4072 Descope to only allow Adhoc Runs and default combo to Adhoc
    'SYS4499 SA 10/6/02 Add functionality back in.
    'g_clsFormProcessing.PopulateComboByValidation cboFrequency, "BatchFrequency", BATCH_FREQ_ADHOC
        
    'Returns a Recordset with the corresponding validation types
    Set rs = m_clsComboValidation.GetComboGroupWithValidationType(COMBO_PROGRAM_TYPE)
    Set cboProgramType.RowSource = rs
    sField = m_clsBatch.GetComboField
    cboProgramType.ListField = sField
     
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : If in Edit mode populate controls with relevant values
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    
    On Error GoTo Failed
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    Dim dTmp As Date
    Dim sTime As String
    Dim sDate As String
    
    txtBatchNumber.Text = m_clsBatch.GetBatchNumber
    m_sBatchNumber = txtBatchNumber.Text
    
    'Description
    txtDescription.Text = m_clsBatch.GetBatchDescription

    'Get the Program Type, and select it from the combo
    'Return the ProgramType from table class
    vTmp = m_clsBatch.GetProgramType
    'Return the ValueName of the ComboGroup Record
    vTmp = g_clsFormProcessing.SetDataComboTextFromValueID(CStr(vTmp), m_clsComboValidation)
    cboProgramType.Text = vTmp
    
    'Run Date/Time
    dTmp = m_clsBatch.GetExecutionDateTime
    
    sDate = Format(dTmp, "Short Date")
    sTime = Format(dTmp, "Short Time")
    txtRunDate.Text = sDate
    txtRunTime.Text = sTime
    
    'Get the Frequency , and select it
    vTmp = m_clsBatch.GetBatchFrequency
    g_clsFormProcessing.HandleComboExtra cboFrequency, vTmp, SET_CONTROL_VALUE
    
    'SYS3327 disable programtype combo.
    cboProgramType.Enabled = False
    If m_bIsView Then
        txtDescription.Enabled = False
        txtRunDate.Enabled = False
        txtRunTime.Enabled = False
        cboFrequency.Enabled = False
        cmdOK.Enabled = False
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub cmdOK_Click()
On Error GoTo Failed
    
    Dim bRet As Boolean
    bRet = DoOKProcessing
    
    If bRet Then
        SetReturnCode
        Hide
    End If
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bShowError As Boolean
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        bRet = ValidateScreenData()
    End If
    
    If bRet Then
        SaveScreenData
    End If
        
    DoOKProcessing = bRet
        
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    Dim sBatchNo As String
    Dim dRun As Date
    Dim dRunTime As Date
    Dim lStatus As Long
    Set clsTableAccess = m_clsBatch
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsBatch
    End If
    
    
    
    If Not m_bIsEdit Then
        'UserID
        m_clsBatch.SetUserID g_sSupervisorUser
        'Status
        lStatus = m_clsComboValidation.GetValueIDForValidationType(BATCH_STATUS_CREATED, COMBO_BATCH_STATUS)
        m_clsBatch.SetStatus lStatus
    End If
    
    'Batch Description
    m_clsBatch.SetDescription txtDescription.Text
    
    'Run Date/Run Time
    If CLng(Val(txtRunTime.ClipText)) > 0 Then
        dRun = txtRunDate.Text & " " & txtRunTime.Text
    Else
        dRun = txtRunDate.Text & " 00:00"
    End If
    
    m_clsBatch.SetExecutionDate dRun

    'Set Frequency
    g_clsFormProcessing.HandleComboExtra cboFrequency, vTmp, GET_CONTROL_VALUE
    m_clsBatch.SetFrequency CLng(vTmp)
    
    'ProgramType
    m_clsBatch.SetProgramType m_clsComboValidation.GetValueID
       
    'Batch Number
    'JD BM0092  Get the new batchnumber
    Dim sValidation As String
    If txtBatchNumber.Text = "" Or m_bIsEdit = False Then
        sValidation = m_clsComboValidation.GetValidationTypeFromCombo(CLng(cboProgramType.SelectedItem))
        txtBatchNumber.Text = m_clsBatch.GetNextBatchNumber(sValidation)
        m_sBatchNumber = txtBatchNumber.Text
    End If
    m_clsBatch.SetBatchNumber txtBatchNumber.Text
    
    
    Set clsTableAccess = m_clsBatch
    clsTableAccess.Update
        
    SavePaymentProcessingParams
    
    'JD BM0092 Let the user know what the batchnumber is
    Dim sMessage As String
    sMessage = "Batch number : " + txtBatchNumber.Text
    MsgBox sMessage
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function ValidateScreenData() As Boolean
    
    On Error GoTo Failed
    'SYS3327 Now unnecessary due to suto batch numbering
'    If m_sBatchNumber <> txtBatchNumber.Text Then
'        ValidateScreenData = ValidateBatchNumber
'    Else
'        ValidateScreenData = True
'    End If
    'SYS3327 What we do need to do though is check the parameters
    'have been input in Add mode - but only for payment processing batch
    'JD BM0092 check the selection type as the batchnumber won't be there yet.
    Dim sValidation As String
    sValidation = m_clsComboValidation.GetValidationTypeFromCombo(CLng(cboProgramType.SelectedItem))
    'If m_bIsEdit = False And m_bIsView = False And Left$(txtBatchNumber.Text, 1) = PAYMENT_PROCESSING_TYPE Then
    If m_bIsEdit = False And m_bIsView = False And sValidation = PAYMENT_PROCESSING_TYPE Then
        If Not m_clsPaymentProc.GetIsUpdated Then
            ValidateScreenData = False
            g_clsErrorHandling.DisplayError "Parameters for payment processing have not been input."
            cmdSetParameters.SetFocus
        Else
            ValidateScreenData = True
        End If
    Else
        ValidateScreenData = True
    End If
        
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateBatchNumber
' Description   : Checks that a BatchNumber has been entered and it is not a duplicate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateBatchNumber() As Boolean
    
    On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim colMatchKeys As Collection
    Dim bRet As Boolean
    Dim sBatchNo As String
    
    Set clsTableAccess = m_clsBatch
    Set colMatchValues = New Collection
    Set colMatchKeys = New Collection
    sBatchNo = txtBatchNumber.Text
    
    colMatchKeys.Add "BATCHNUMBER"
    colMatchValues.Add sBatchNo
    
    bRet = Not clsTableAccess.DoesRecordExist(colMatchValues, colMatchKeys)
    
    
    If Not bRet Then
        g_clsErrorHandling.DisplayError "Batch Number must be unique"
        txtBatchNumber.SetFocus
    End If
    
    ValidateBatchNumber = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SavePaymentProcessingParams
' Description   : Updates the PayProcBatchParams table with the new/amended values.
'                 This is a sperate function because we dont want to update this table until OK
'                 is pressed on the frmEditbatch Form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub SavePaymentProcessingParams()
    On Error GoTo Failed
    
    Dim bUpdated As Boolean
    Dim clsTableAccess As TableAccess
    
    
    Set clsTableAccess = m_clsPaymentProc
    bUpdated = m_clsPaymentProc.GetIsUpdated
    
    If bUpdated Then
        'JD BM0092 set the batchnumber
        m_clsPaymentProc.SetBatchNumber m_sBatchNumber
        clsTableAccess.Update
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

