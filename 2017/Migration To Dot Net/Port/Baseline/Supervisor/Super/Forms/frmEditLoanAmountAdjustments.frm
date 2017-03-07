VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditLoanAmountAdjustments 
   Caption         =   "Add/Edit Loan Amount Adjustments"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmEditLoanAmountAdjustments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "A&nother"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin MSGOCX.MSGComboBox cboProductCategory 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSGOCX.MSGEditBox txtMaximumLoan 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MaxValue        =   ""
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
      MaxLength       =   8
   End
   Begin MSGOCX.MSGEditBox txtFeeRate 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      TextType        =   2
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      MinValue        =   "0"
      MaxValue        =   "10.000"
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
      MaxLength       =   6
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fee Rate"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maximum Loan"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Product Category"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "frmEditLoanAmountAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditLoanAmountAdjustments
' Description   :
'
' Change history
' Prog      Date        Description
' TW        14/12/2006  EP2_518 - Created
' TW        21/12/2006  EP2_626 - Added 'Another' button functionality
' TW        02/01/2007  EP2_634 - Fee rate validation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

Private m_clsProcFeeAdjByLoanTable As ProcFeeAdjByLoanTable
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetIsEdit() As Boolean
    GetIsEdit = m_bIsEdit
End Function

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(blnEditStatus As Boolean)
    m_bIsEdit = blnEditStatus
End Sub
Private Sub PopulateScreenFields()
    
Const c_strTHIS_FUNCTION As String = "PopulateScreenFields"

    On Error GoTo Failed

    g_clsFormProcessing.HandleComboExtra cboProductCategory, m_clsProcFeeAdjByLoanTable.GetProductCategory, SET_CONTROL_VALUE
    
    txtMaximumLoan.Text = m_clsProcFeeAdjByLoanTable.GetMaxLoan
    txtFeeRate.Text = m_clsProcFeeAdjByLoanTable.GetFeeRate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub SetEditState()
Dim objRS As ADODB.Recordset
Dim objTableAccess As TableAccess
Dim sGuid As Variant
Dim sDepartmentID As String
Dim colValues As New Collection
    
    On Error GoTo Failed
    
    Set objTableAccess = m_clsProcFeeAdjByLoanTable
    objTableAccess.SetKeyMatchValues m_colKeys
    
    Set objRS = objTableAccess.GetTableData()
    
    If Not objRS Is Nothing Then
        If objRS.RecordCount > 0 Then
            cboProductCategory.Enabled = False
            txtMaximumLoan.Enabled = False
            PopulateScreenFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Edit Loan Amount Adjustyments - no records exist"
        End If
    End If
    
    Set objRS = Nothing
    Set objTableAccess = Nothing
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub EnableOK()
    cmdOK.Enabled = (cboProductCategory.ListIndex > -1 And Val(Format$(txtMaximumLoan.Text)) > 0 And Len(txtFeeRate.Text) > 0)
    cmdAnother.Enabled = cmdOK.Enabled
End Sub





Private Sub cboProductCategory_Click()
    Call EnableOK
End Sub

Private Sub cmdAnother_Click()
' TW 21/12/2006 EP2_626
Dim bRet As Boolean

    On Error GoTo Failed
    
    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
    
        g_clsFormProcessing.ClearScreenFields Me

        Form_Load
        frmMain.PopulateListView
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
' TW 21/12/2006 EP2_626 End
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub


Private Sub cmdOK_Click()
Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = DoOKProcessing()
    If bRet = True Then
        SetReturnCode
        Hide
    
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        SaveScreenData
        SaveChangeRequest
    End If
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function
Private Sub SaveChangeRequest()
Dim colMatchValues As Collection
Dim clsTableAccess As TableAccess
    On Error GoTo Failed
    
    Set colMatchValues = New Collection
    colMatchValues.Add m_colKeys(1)
    colMatchValues.Add m_colKeys(2)
    
    Set clsTableAccess = m_clsProcFeeAdjByLoanTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsProcFeeAdjByLoanTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SaveScreenData()
Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    Set clsTableAccess = m_clsProcFeeAdjByLoanTable
    
    If Not m_bIsEdit Then
        Set m_colKeys = New Collection
        m_colKeys.Add cboProductCategory.GetExtra(cboProductCategory.ListIndex)
        m_colKeys.Add txtMaximumLoan.Text
        
        clsTableAccess.SetKeyMatchValues m_colKeys
        
        'Create a new record.
        g_clsFormProcessing.CreateNewRecord m_clsProcFeeAdjByLoanTable
    End If
    
    If cboProductCategory.ListIndex > -1 Then
        m_clsProcFeeAdjByLoanTable.SetProductCategory cboProductCategory.GetExtra(cboProductCategory.ListIndex)
    End If
    m_clsProcFeeAdjByLoanTable.SetMaxLoan txtMaximumLoan.Text
    m_clsProcFeeAdjByLoanTable.SetFeeRate txtFeeRate.Text
    
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub




Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    Set m_clsProcFeeAdjByLoanTable = New ProcFeeAdjByLoanTable
    
    g_clsFormProcessing.PopulateCombo "NatureOfLoan", Me.cboProductCategory
    If m_bIsEdit = True Then
        SetEditState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub txtFeeRate_Change()
    Call EnableOK
End Sub

Private Sub txtFeeRate_Validate(Cancel As Boolean)
    Cancel = Not txtFeeRate.ValidateData()
End Sub


Private Sub txtMaximumLoan_Change()
    Call EnableOK
End Sub


Private Sub txtMaximumLoan_Validate(Cancel As Boolean)
    Cancel = Not txtMaximumLoan.ValidateData()
End Sub


