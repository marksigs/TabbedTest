VERSION 5.00
Object = "{8FBFAD4D-5ED6-467A-98A5-FAA33BE4B270}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPaymentProcessing 
   Caption         =   "Payment Processing"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmEditPaymentProcessing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboPaymentType 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
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
      Text            =   ""
   End
   Begin MSGOCX.MSGEditBox txtFrom 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   12
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraPaymentMethod 
      Caption         =   "Payment Method"
      Enabled         =   0   'False
      Height          =   2115
      Left            =   120
      TabIndex        =   15
      Top             =   1380
      Width           =   7095
      Begin VB.CommandButton cmdDeselectAll 
         Caption         =   "&Deselect All"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "&Select All"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox chkPaymentMethod 
         Caption         =   "Internal Transfer"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   420
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkPaymentMethod 
         Caption         =   "Cheque"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   420
         TabIndex        =   5
         Top             =   780
         Width           =   1395
      End
      Begin VB.CheckBox chkPaymentMethod 
         Caption         =   "CHAPS/TT"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   420
         TabIndex        =   4
         Top             =   480
         Width           =   1395
      End
      Begin VB.CheckBox chkPaymentMethod 
         Caption         =   "BACS"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   3
         Top             =   300
         Width           =   1395
      End
   End
   Begin MSGOCX.MSGEditBox txtTo 
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   12
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   255
      Left            =   4860
      TabIndex        =   14
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblFrom 
      Caption         =   "From"
      Height          =   255
      Left            =   1860
      TabIndex        =   13
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblAppRange 
      Caption         =   "Application Range"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label lblJobType 
      Caption         =   "Payment Job Type"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmEditPaymentProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPaymentProcessing
' Description   : Form which allows the adding/editing of Payments
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added Form
' SA        18/01/02    SYS3327 Handle View Mode
' SA        15/02/02    SYS4071 PaymentType Datacombo changed to normal combo
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'   BMIDS
'   AW      02/08/02    BM029 Cheque option is now controled by global parameter 'ProcessChequePaymentsFlag'
'   JD      13/12/2004  BM0092 Don't set the batchnumber. Do it on save of EditBatch
'


Option Explicit
Private m_ReturnCode As MSGReturnCode

'ComboGroup Constants
Private Const PAYMENT_JOB_TYPE_COMBO = "PaymentJobType"
Private Const BACS_CHECK_BOX = 0
Private Const TRANSFER_CHECK_BOX = 3
Private Const CHEQUE_CHECK_BOX = 2
Private Const CHAPS_CHECK_BOX = 1
Private Const COMP_APP_VALIDATION_TYPE = "C"

Private m_clsComboValidation As ComboValidationTable
Private m_clsPaymentProcessing As PaymentProcessingTable
Private m_sBatchNumber As String
Private m_colKeys  As Collection
Private m_bIsEdit As Boolean
Private m_bIsView As Boolean 'SYS3327

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Sub SetIsView(Optional bIsView As Boolean = True)
    m_bIsView = bIsView
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Public Sub SetTableClass(clsTable As TableAccess)
    Set m_clsPaymentProcessing = clsTable
End Sub
Private Sub cboPaymentType_Click()
Dim sValidation As String
    On Error GoTo Failed

    'Get the validationType of the selected item
    'SYS4071 Datacombo now ordinary combo
    'sValidation = m_clsComboValidation.GetValidationTypeFromCombo(CLng(cboPaymentType.SelectedItem))
    g_clsFormProcessing.GetComboValidation cboPaymentType, PAYMENT_JOB_TYPE_COMBO, sValidation


    Select Case sValidation
        Case PAY_SANCTIONED_PAYMENTS
            EnablePaymentMethodControls True
        Case Else
            CheckPaymentMethod False
            EnablePaymentMethodControls
    End Select

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION

End Sub




Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdDeselectAll_Click()
    
    On Error GoTo Failed
    CheckPaymentMethod False
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing
    
    If bRet Then
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub cmdSelectAll_Click()
    On Error GoTo Failed
    'Check all Checkboxes
    CheckPaymentMethod
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CheckPaymentMethod
' Description   :   Checks or unchecks (depending on the value of input parameter bCheck) all
'                   Payment Method checkboxes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckPaymentMethod(Optional bCheck As Boolean = True)
    
    On Error GoTo Failed
    Dim nValue As Integer
    
    If bCheck Then
        nValue = 1
    Else
        nValue = 0
    End If
    
    chkPaymentMethod(BACS_CHECK_BOX).Value = nValue
    chkPaymentMethod(TRANSFER_CHECK_BOX).Value = nValue
    
    '   AW  02/08/02  BM029
    Dim bProcessCheques As Boolean
    bProcessCheques = g_clsGlobalParameter.FindBoolean("ProcessChequePaymentsFlag")
    
    If bProcessCheques = True Then
        chkPaymentMethod(CHEQUE_CHECK_BOX).Value = nValue
    Else
        chkPaymentMethod(CHEQUE_CHECK_BOX).Value = False
    End If
    
    chkPaymentMethod(CHAPS_CHECK_BOX).Value = nValue
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenControls
' Description   :   Populate combos
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    On Error GoTo Failed
        
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sField As String

    'Returns a Recordset with the corresponding validation types
'SYS4071 Change to Normal Combo
'    Set rs = m_clsComboValidation.GetComboGroupWithValidationType(PAYMENT_JOB_TYPE_COMBO)
'    Set cboPaymentType.RowSource = rs
'    sField = m_clsPaymentProcessing.GetComboField
'    cboPaymentType.ListField = sField
'    Set clsTableAccess = m_clsComboValidation
'    clsTableAccess.SetRecordSet rs
    
    'SYS4071
    PopulatePaymentJobCombo PAYMENT_JOB_TYPE_COMBO, cboPaymentType, False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub Form_Load()

    On Error GoTo Failed
    
    If m_clsPaymentProcessing Is Nothing Then
        Set m_clsPaymentProcessing = New PaymentProcessingTable
    End If
    
    Set m_clsComboValidation = New ComboValidationTable
    m_sBatchNumber = m_colKeys(1)
    
    m_bIsEdit = DoesRecordExist
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetFormEditState
    Else
        SetFormAddState
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
    
    If bRet Then
        bRet = ValidateScreenData
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
    Dim lStatus As Long
    Set clsTableAccess = m_clsPaymentProcessing
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsPaymentProcessing
    End If
    
    'JD BM0092 BatchNumber saved on save of EditBatch
    'Batch Number
    'm_clsPaymentProcessing.SetBatchNumber m_sBatchNumber
    
    'From Application Number
    m_clsPaymentProcessing.SetApplicationNumberFrom txtFrom.Text

    'To Application Number
    m_clsPaymentProcessing.SetApplicationNumberTo txtTo.Text
    
    'Set Program Job Type
    'SYS4071
    'm_clsPaymentProcessing.SetPaymentJobType m_clsComboValidation.GetValueID
    g_clsFormProcessing.HandleComboExtra cboPaymentType, vTmp, GET_CONTROL_VALUE
    m_clsPaymentProcessing.SetPaymentJobType CLng(vTmp)

    
    'BACS
    m_clsPaymentProcessing.SetBacs CInt(chkPaymentMethod(BACS_CHECK_BOX).Value)
    ' Chaps
    m_clsPaymentProcessing.SetChaps CInt(chkPaymentMethod(CHAPS_CHECK_BOX).Value)
    'Cheque
    m_clsPaymentProcessing.SetCheque CInt(chkPaymentMethod(CHEQUE_CHECK_BOX).Value)
    'Internal Transfer
    m_clsPaymentProcessing.SetInternalTransfer CInt(chkPaymentMethod(TRANSFER_CHECK_BOX).Value)
    
'    Set clsTableAccess = m_clsPaymentProcessing
'    clsTableAccess.Update
    
    m_clsPaymentProcessing.SetUpdated True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Returns a boolean indicating whether the batchnumber exists
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist() As Boolean

    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim bRet As Boolean
    Dim colFields As Collection
    
    Set colFields = New Collection

    Set clsTableAccess = m_clsPaymentProcessing
    
    colFields.Add "BATCHNUMBER"
    
    'JD BM0092 if we have a batchnumber
    If m_sBatchNumber <> "" Then
        bRet = clsTableAccess.DoesRecordExist(m_colKeys, colFields)
    Else
        bRet = False
    End If
    
    DoesRecordExist = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Edit specific code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormEditState()
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    
    Set clsTableAccess = m_clsPaymentProcessing
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.DisplayError "Payment Processing Parameter - record not found"
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets the values of the screen controls on the form
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed
    
    Dim vTmp As Variant
    Dim bTmp As Boolean
    Dim i As Integer
    'Payement Job Type
    
'   SYS4071 Change to Normal Combo
    vTmp = m_clsPaymentProcessing.GetPaymentJobType
    'vTmp = g_clsFormProcessing.SetDataComboTextFromValueID(CStr(vTmp), m_clsComboValidation)
    g_clsFormProcessing.HandleComboExtra cboPaymentType, vTmp, SET_CONTROL_VALUE
    'cboPaymentType.Text = vTmp

    
    'From Date
    txtFrom.Text = m_clsPaymentProcessing.GetApplicationNoFrom
    
    'To Date
    txtTo.Text = m_clsPaymentProcessing.GetApplicationNoTo
    
    'Payment Methods
    
    'Chaps
    bTmp = m_clsPaymentProcessing.GetChaps
    SetCheckBoxValue CHAPS_CHECK_BOX, bTmp
        
    'BACS
    bTmp = m_clsPaymentProcessing.GetBacs
    SetCheckBoxValue BACS_CHECK_BOX, bTmp
    
    'Internal Transfer
    bTmp = m_clsPaymentProcessing.GetInternalTransfer
    SetCheckBoxValue TRANSFER_CHECK_BOX, bTmp
    
    'Cheque
    '   AW  02/08/02  BM029
    Dim bProcessCheques As Boolean
    bProcessCheques = g_clsGlobalParameter.FindBoolean("ProcessChequePaymentsFlag")
    
    If bProcessCheques = True Then
        bTmp = m_clsPaymentProcessing.GetCheque
    Else
        SetCheckBoxValue CHEQUE_CHECK_BOX, False
        chkPaymentMethod(CHEQUE_CHECK_BOX).Enabled = False
    End If
        
    'SYS3327
    If m_bIsView Then
        txtFrom.Enabled = False
        txtTo.Enabled = False
        For i = 0 To chkPaymentMethod.Count - 1
            chkPaymentMethod.Item(i).Enabled = False
        Next i
        cmdOK.Enabled = False
        cmdDeselectAll.Enabled = False
        cmdSelectAll.Enabled = False
        'SYS4071 Change to Normal Combo
        cboPaymentType.Enabled = False
    End If
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetCheckBoxValue
' Description   :   Given an index, the checkbox(index) is set or unset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCheckBoxValue(nIndx As Integer, Optional bValue As Boolean = True)

    On Error GoTo Failed
    
    If bValue Then
        chkPaymentMethod(nIndx).Value = 1
    Else
        chkPaymentMethod(nIndx).Value = 0
    End If
    
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateScreenData
' Description   :   Returns a boolean indicating whether or not the data entered on the form
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    
    On Error GoTo Failed
    Dim bRet As Boolean
    
   bRet = True
    If Len(txtFrom.Text) > 0 Or Len(txtFrom.Text) > 0 Then
        'Is there a from App No
        If Len(txtFrom.Text) > 0 Then
            If Len(txtTo.Text) = 0 Then
                txtTo.SetFocus
                bRet = False
            End If
            
        ElseIf Len(txtTo.Text) > 0 Then
            If Len(txtFrom.Text) = 0 Then
                bRet = False
            End If
        End If
    End If
                
    If Not bRet Then
        g_clsErrorHandling.DisplayError "Where an application range is required, both fields must be input"
    End If
    
    ValidateScreenData = bRet
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   EnablePaymentMethodControls
' Description   :   Enables or disables (depending on the value of input bEnable) all controls
'                   in the Payment method frame
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnablePaymentMethodControls(Optional bCheck As Boolean = False)
    On Error GoTo Failed

    Dim ctrl As Control
    '   AW  02/08/02  BM029
    Dim bProcessCheques As Boolean
    bProcessCheques = g_clsGlobalParameter.FindBoolean("ProcessChequePaymentsFlag")
    
    For Each ctrl In Me
        If ctrl.Container.Name = fraPaymentMethod.Name Or ctrl Is fraPaymentMethod Then
            'The parent of this control is the PaymentMethod Frame (Determined by comparing Ctrl.name with the PaymentMethod Control Name)
            '   AW  02/08/02  BM029
            If (ctrl.Name = "chkPaymentMethod" And ctrl.Caption = "Cheque" And bProcessCheques = False) Then
                ctrl.Enabled = False
            Else
                ctrl.Enabled = bCheck
            End If
        End If
    Next
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub PopulatePaymentJobCombo(sGroup As String, combo As MSGComboBox, Optional bIncludeOptionNone As Boolean = True)
    Dim bRet As Boolean
    Dim colNames As New Collection
    Dim colValues As New Collection
    Dim clsComboTable As New ComboValueTable
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    bRet = True
    
    bRet = clsComboTable.GetPaymentJobComboValues(sGroup, colNames, colValues, COMP_APP_VALIDATION_TYPE)
    
    If bRet = True Then
        If bIncludeOptionNone Then
            colNames.Add COMBO_NONE
        End If
        colValues.Add colNames.Count
        
        combo.SetListTextFromCollection colNames, colValues

    End If
    
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to find combo: " + sGroup
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

