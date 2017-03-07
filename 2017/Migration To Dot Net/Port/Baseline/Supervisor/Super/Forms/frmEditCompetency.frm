VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditCompetency 
   Caption         =   "Add/Edit Competency"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "frmEditCompetency.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   7
      Left            =   4620
      TabIndex        =   6
      Top             =   1380
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
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
   End
   Begin MSGOCX.MSGComboBox cboCompetency 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
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
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4395
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3075
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1755
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   1
      Left            =   4620
      TabIndex        =   2
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   3
      Top             =   960
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      TextType        =   7
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
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   3
      Left            =   4620
      TabIndex        =   4
      Top             =   960
      Width           =   990
      _ExtentX        =   1746
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
      MaxLength       =   5
   End
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   4
      Left            =   1740
      TabIndex        =   5
      Top             =   1380
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      TextType        =   2
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
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   5
      Left            =   1740
      TabIndex        =   7
      Top             =   1800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      TextType        =   7
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
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   855
      Index           =   6
      Left            =   1740
      TabIndex        =   8
      Top             =   2220
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1508
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
      MaxLength       =   30
   End
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtCompetency 
      Height          =   315
      Index           =   8
      Left            =   4620
      TabIndex        =   21
      Top             =   1800
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
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
   End
   Begin VB.Label lblLabel 
      Caption         =   "User Mandate Increase Limit %"
      Height          =   375
      Left            =   3180
      TabIndex        =   22
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "User Mandate Decrease Limit %"
      Height          =   375
      Left            =   3180
      TabIndex        =   20
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Risk Assessment Mandate Level"
      Height          =   552
      Index           =   7
      Left            =   3180
      TabIndex        =   19
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Notes"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   18
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Funds Release Limit"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Loan Mandate LTV"
      Height          =   315
      Index           =   5
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Competency Type"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   180
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Active From"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Active To"
      Height          =   255
      Index           =   4
      Left            =   3180
      TabIndex        =   13
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Loan Mandate Level"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   1020
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditCompetency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditCompetency
' Description   :   Form which allows the user to edit and add Competency details
' Prog      Date        Description
' CL        17/05/01    Added 'User Mandate Decrease Limit % field
' STB       06/12/01    SYS1942 - Another button commits current transaction.
' STB       22/04/02    MSMS0035 Added the UserIncreaseLimitPercent field.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History:
'Prog   Date        Description
'MV     11/12/2002  BM0085  Enhanced to support Audit Trail for Competency
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Text edit indexes
Private Const ACTIVE_FROM = 0
Private Const ACTIVE_TO = 1
Private Const LOAN_MANDATE = 2
Private Const RA_MANDATE_LEVEL = 3
Private Const LOAN_MANDATE_LTV = 4
Private Const FUNDS_RELEASE_LIMIT = 5
Private Const COMPETENCY_DESCRIPTION = 6
Private Const USER_MANDATE_DECREASE_LIMIT_PERCENT = 7
Private Const USER_INCREASE_LIMIT_PERCENT = 8

' Private data
Private m_bIsEdit As Boolean
Private m_clsCompetency As CompetencyTable
Private m_sCompetencyType As String
Private m_bUpdateFailed As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_sUserID As String

Private m_clsBMIDSCompetencyAudit As CompetencyAuditTable
Private m_bIschanged As Boolean
Private m_bFormLoaded As Boolean
Public Sub SaveCompetencyAuditData()
    
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
  
    
    g_clsFormProcessing.HandleDate txtCompetency(ACTIVE_FROM), vTmp, GET_CONTROL_VALUE
    m_clsBMIDSCompetencyAudit.SetActiveFrom vTmp
    
    g_clsFormProcessing.HandleDate txtCompetency(ACTIVE_TO), vTmp, GET_CONTROL_VALUE
    m_clsBMIDSCompetencyAudit.SetActiveTo vTmp
    
    m_clsBMIDSCompetencyAudit.SetLoanAmountMandate txtCompetency(LOAN_MANDATE).Text
    m_clsBMIDSCompetencyAudit.SetLTVMandate txtCompetency(LOAN_MANDATE_LTV).Text
    m_clsBMIDSCompetencyAudit.SetFundsReleaseMandate txtCompetency(FUNDS_RELEASE_LIMIT).Text
    m_clsBMIDSCompetencyAudit.SetRiskAssmentMandateLevel txtCompetency(RA_MANDATE_LEVEL).Text

    g_clsFormProcessing.HandleComboExtra Me.cboCompetency, vTmp, GET_CONTROL_VALUE
    m_sCompetencyType = CStr(vTmp)
    m_clsBMIDSCompetencyAudit.SetType m_sCompetencyType
    
    Set clsTableAccess = m_clsBMIDSCompetencyAudit
    clsTableAccess.Update
    
    m_bUpdateFailed = False
    
    Exit Sub
Failed:
    m_bUpdateFailed = True
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetCompetencyAuditData()
    
    On Error GoTo Failed
    
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    g_clsFormProcessing.CreateNewRecord m_clsBMIDSCompetencyAudit
    
    g_clsFormProcessing.HandleComboExtra Me.cboCompetency, vTmp, GET_CONTROL_VALUE
    m_sCompetencyType = CStr(vTmp)
    m_clsBMIDSCompetencyAudit.SetType m_sCompetencyType
    
    m_clsBMIDSCompetencyAudit.SetAuditDate Now()
    
    m_clsBMIDSCompetencyAudit.SetChangeUser ""
    m_clsBMIDSCompetencyAudit.SetChangeUser g_sSupervisorUser
    
    m_clsBMIDSCompetencyAudit.SetPrevActiveFrom m_clsCompetency.GetActiveFrom()
    m_clsBMIDSCompetencyAudit.SetPrevActiveTo m_clsCompetency.GetActiveTo()
    m_clsBMIDSCompetencyAudit.SetPrevFundsReleaseMandate m_clsCompetency.GetFundsReleaseMandate
    m_clsBMIDSCompetencyAudit.SetPrevLoanAmountMandate m_clsCompetency.GetLoanAmountMandate
    m_clsBMIDSCompetencyAudit.SetPrevLTVMandate m_clsCompetency.GetLTVMandate
    m_clsBMIDSCompetencyAudit.SetPrevRiskAssmentMandateLevel m_clsCompetency.GetRiskAssessmentMandateLevel
    
    m_clsBMIDSCompetencyAudit.SetActiveFrom ""
    m_clsBMIDSCompetencyAudit.SetActiveTo ""
    m_clsBMIDSCompetencyAudit.SetFundsReleaseMandate ""
    m_clsBMIDSCompetencyAudit.SetLoanAmountMandate ""
    m_clsBMIDSCompetencyAudit.SetLTVMandate ""
    m_clsBMIDSCompetencyAudit.SetRiskAssmentMandateLevel ""
    
    Exit Sub
    
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub


Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsCompetency = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
Private Sub cboCompetency_Validate(Cancel As Boolean)
    Cancel = Not Me.cboCompetency.ValidateData()
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user presses the Another button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError

End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok or
'                   presses Another
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            If Not m_bUpdateFailed And m_bIschanged Then
                SaveCompetencyAuditData
            End If
            SaveChangeRequest
        End If
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user pressed the OK button. Performs necessary
'                   validation and saves any data that needs to be saved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateScreenData
' Description   :   Validates that all date entered in the screen is valid, and
'                   a record doesn't already exist.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean

    bRet = True
    
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    
        If bRet = False Then
            MsgBox "Competency Type already exists - please enter a unique competency", vbCritical
            cboCompetency.SetFocus
        End If
    End If
    
    If bRet = True Then
        bRet = g_clsValidation.ValidateActiveFromTo(txtCompetency(ACTIVE_FROM), txtCompetency(ACTIVE_TO))
    End If
    
    ValidateScreenData = bRet
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesRecordExist
' Description   :   Returns true if a record exists on the database with the
'                   same keys the user has entered, false if not.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sType As String
    Dim clsTableAccess As TableAccess
    Dim col As New Collection

    g_clsFormProcessing.HandleComboExtra Me.cboCompetency, sType, GET_CONTROL_VALUE
    
    If Len(sType) > 0 Then
        col.Add sType
        
        Set clsTableAccess = m_clsCompetency
        bRet = clsTableAccess.DoesRecordExist(col)
    Else
        MsgBox "Compentency Type is empty"
    End If

    DoesRecordExist = bRet
End Function
Private Sub Form_Initialize()
    m_bUpdateFailed = False
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Initialisation to the form
' CL        28/05/02    SYS4766 Merge MSMS & CORE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    SetReturnCode MSGFailure
    
    Set m_clsCompetency = New CompetencyTable
    Set m_clsBMIDSCompetencyAudit = New CompetencyAuditTable
    g_clsFormProcessing.PopulateCombo "CompetencyType", Me.cboCompetency

    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    m_bFormLoaded = True
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Public Sub SetAddState()
    g_clsFormProcessing.CreateNewRecord m_clsCompetency
    SetCompetencyAuditData
End Sub
Public Sub SetEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsCompetency
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set clsTableAccess = m_clsCompetency
    Set rs = clsTableAccess.GetTableData()
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            PopulateScreenFields
            SetCompetencyAuditData
        End If
    End If
    cmdAnother.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_bFormLoaded = False
    m_bIschanged = False
End Sub

Private Sub txtCompetency_Change(Index As Integer)
    If m_bFormLoaded = True Then
        m_bIschanged = True
    End If
End Sub

Private Sub txtCompetency_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtCompetency(Index).ValidateData()
End Sub

Public Sub PopulateScreen()
    Dim bRet As Boolean
    On Error GoTo Failed

    bRet = PopulateScreenFields()

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
End Sub
Public Function PopulateScreenFields() As Boolean
    
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant

    txtCompetency(COMPETENCY_DESCRIPTION).Text = m_clsCompetency.GetType()
    txtCompetency(COMPETENCY_DESCRIPTION).Text = m_clsCompetency.GetDescription()
    txtCompetency(ACTIVE_FROM).Text = m_clsCompetency.GetActiveFrom()
    txtCompetency(ACTIVE_TO).Text = m_clsCompetency.GetActiveTo()
    txtCompetency(LOAN_MANDATE).Text = m_clsCompetency.GetLoanAmountMandate()
    txtCompetency(LOAN_MANDATE_LTV).Text = m_clsCompetency.GetLTVMandate()
    txtCompetency(FUNDS_RELEASE_LIMIT).Text = m_clsCompetency.GetFundsReleaseMandate()
    txtCompetency(RA_MANDATE_LEVEL).Text = m_clsCompetency.GetRiskAssessmentMandateLevel()
    txtCompetency(USER_MANDATE_DECREASE_LIMIT_PERCENT).Text = m_clsCompetency.GetUserMandateDecreaseLimitPercent()
    
    'MSMS0035 - Added UserIncrease field.
    txtCompetency(USER_INCREASE_LIMIT_PERCENT).Text = m_clsCompetency.GetUserIncreaseLimitPercent()
    'MSMS0035 - End.

    ' Competency combo
    vTmp = m_clsCompetency.GetType()
    g_clsFormProcessing.HandleComboExtra Me.cboCompetency, vTmp, SET_CONTROL_VALUE

    PopulateScreenFields = True
    Exit Function
Failed:
    PopulateScreenFields = False
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()

    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
        
    m_clsCompetency.SetDescription txtCompetency(COMPETENCY_DESCRIPTION).Text
    
    g_clsFormProcessing.HandleDate txtCompetency(ACTIVE_FROM), vTmp, GET_CONTROL_VALUE
    m_clsCompetency.SetActiveFrom vTmp
    
    g_clsFormProcessing.HandleDate txtCompetency(ACTIVE_TO), vTmp, GET_CONTROL_VALUE
    m_clsCompetency.SetActiveTo vTmp
    
    m_clsCompetency.SetLoanAmountMandate txtCompetency(LOAN_MANDATE).Text
    m_clsCompetency.SetLTVMandate txtCompetency(LOAN_MANDATE_LTV).Text
    m_clsCompetency.SetFundsReleaseMandate txtCompetency(FUNDS_RELEASE_LIMIT).Text
    m_clsCompetency.SetRiskAssmentMandateLevel txtCompetency(RA_MANDATE_LEVEL).Text


     m_clsCompetency.SetUserMandateDecreaseLimitPercent txtCompetency(USER_MANDATE_DECREASE_LIMIT_PERCENT).Text
     
     'MSMS0035 - Added UserIncrease field.
     m_clsCompetency.SetUserIncreaseLimitPercent txtCompetency(USER_INCREASE_LIMIT_PERCENT).Text
     'MSMS0035 - End.

    ' Competency combo
    g_clsFormProcessing.HandleComboExtra Me.cboCompetency, vTmp, GET_CONTROL_VALUE
    m_sCompetencyType = CStr(vTmp)
    m_clsCompetency.SetType m_sCompetencyType
    
    Set clsTableAccess = m_clsCompetency
    clsTableAccess.Update
    
    m_bUpdateFailed = False
    
    Exit Sub
Failed:
    m_bUpdateFailed = True
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim sThisVal As Variant
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    colMatchValues.Add m_sCompetencyType
    Set clsTableAccess = m_clsCompetency
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsCompetency
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
