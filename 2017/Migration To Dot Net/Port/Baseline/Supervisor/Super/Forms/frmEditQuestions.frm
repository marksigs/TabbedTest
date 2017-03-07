VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditQuestions 
   Caption         =   "Questions"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmEditQuestions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Frame fraQuestions 
      Height          =   4395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtQuestionShortText 
         Height          =   315
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   16
         Top             =   1380
         Width           =   3795
      End
      Begin MSGOCX.MSGTextMulti txtQuestionText 
         Height          =   1215
         Left            =   2040
         TabIndex        =   2
         Top             =   1920
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         MaxLength       =   255
      End
      Begin VB.TextBox txtQuestionReference 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkDeleteFlag 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   3780
         Width           =   435
      End
      Begin MSGOCX.MSGComboBox cboQuestionType 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   900
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Frame fraDeleteFlag 
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   1980
         TabIndex        =   14
         Top             =   3000
         Width           =   2415
         Begin VB.OptionButton optDetailsRequired 
            Caption         =   "No"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optDetailsRequired 
            Caption         =   "Yes"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Question Short Text"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblDetailsRequired 
         Caption         =   "Are Details Required?"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblDeleteFlag 
         Caption         =   "Delete Flag"
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label lblQuestionText 
         Caption         =   "Question Text"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblQuestionReference 
         Caption         =   "Question Reference"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmEditQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditQuestions
' Description   : Form which allows the adding/editing of Additional Questions
'
' Change history
' Prog      Date        AQR     Description
' AA        22/01/00            Added Form
' BG        23/11/01    SYS3122 Added QuestionShortText field
' STB       06/12/01    SYS1942 - Another button commits current transaction.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const DETAILS_REQUIRED_YES = 0
Private Const DETAILS_REQUIRED_NO = 1
Private Const QUESTION_COMBO_GROUP As String = "AdditionalQuestionType"
Private m_colKeys As Collection
Private m_sQuestionReference As String
Private m_clsQuestions As AdditionalQuestionsTable
Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit  As Boolean
Private m_bScreenUpdated As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub cboQuestionType_Click()
    m_bScreenUpdated = True
End Sub


Private Sub chkDeleteFlag_Click()
    m_bScreenUpdated = True
End Sub

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = DoOKProcessing
    
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
        cboQuestionType.SetComboFocus
        chkDeleteFlag.Value = 0
        optDetailsRequired(DETAILS_REQUIRED_NO).Value = True
        ' BG 23/11/01 SYS3122
        txtQuestionShortText.Text = ""
    End If

    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Fired when the User presses the OK button. Validates and writes the data to the
'                   Database if screen data is valid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = DoOKProcessing
    If bRet Then
        SetReturnCode
        Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Fired when the form is loaded. Displays the form in the appropriate mode (Add/Edit)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    SetReturnCode MSGFailure
    
    PopulateScreenControls
    
    Set m_clsQuestions = New AdditionalQuestionsTable
    
    If m_bIsEdit = True Then
        SetFormEditState
        m_sQuestionReference = txtQuestionReference.Text
    Else
        SetFormAddState
    End If
    

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormAddState
' Description   :   Initialises the form in Add Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormAddState()
    chkDeleteFlag.Enabled = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetFormEditState
' Description   :   Initialises the form in Edit Mode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetFormEditState()
 
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sDepartmentID As String
    Dim colValues As New Collection
    On Error GoTo Failed

    Set clsTableAccess = m_clsQuestions
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()
    cmdAnother.Enabled = False

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            g_clsErrorHandling.DisplayError "Additional Questions - no records to edit"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenControls
' Description   :   Populates cboQuestionType with the values stored in the combogroup 'QuestionType'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    
    g_clsFormProcessing.PopulateCombo QUESTION_COMBO_GROUP, cboQuestionType

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets all screen fields to the values stored on the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    On Error GoTo Failed
    txtQuestionReference.Text = m_clsQuestions.GetQuestionReference
    
    vTmp = m_clsQuestions.GetQuestionType()
    
    g_clsFormProcessing.HandleComboExtra cboQuestionType, vTmp, SET_CONTROL_VALUE
    
    txtQuestionText.Text = m_clsQuestions.GetQuestionText
    
    ' BG 23/11/01 SYS3122
    txtQuestionShortText.Text = m_clsQuestions.GetQuestionShortText
    
    bRet = GetBooleanFromNumberAsBoolean(m_clsQuestions.GetDeleteFlag)
    
    If bRet Then
        'Set Delete Flag
        chkDeleteFlag.Value = 1
    Else
        chkDeleteFlag.Value = 0
    End If
    
    bRet = GetBooleanFromNumberAsBoolean(m_clsQuestions.GetDetailsRequired)
    
    If bRet Then
        'Yes
        optDetailsRequired(DETAILS_REQUIRED_YES).Value = True
    Else
        optDetailsRequired(DETAILS_REQUIRED_NO).Value = True
    End If
    m_bScreenUpdated = False
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsQuestions
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsQuestions
    End If
    
    m_clsQuestions.SetQuestionText txtQuestionText.Text
    
    ' BG 23/11/01 SYS3122
    m_clsQuestions.SetQuestionShortText txtQuestionShortText.Text
    
    If optDetailsRequired(DETAILS_REQUIRED_YES).Value = True Then
        m_clsQuestions.SetDetailsRequired GetNumberFromBoolean(True)
    Else
        m_clsQuestions.SetDetailsRequired GetNumberFromBoolean(False)
    End If
    
    g_clsFormProcessing.HandleComboExtra cboQuestionType, vTmp, GET_CONTROL_VALUE
    m_clsQuestions.SetQuestionType CStr(vTmp)
    
    If chkDeleteFlag.Value = 0 Then
        m_clsQuestions.SetDeleteFlag GetNumberFromBoolean(False)
    Else
        m_clsQuestions.SetDeleteFlag GetNumberFromBoolean(True)
    End If
    
    If m_bIsEdit Then
        m_clsQuestions.SetQuestionReferenceID txtQuestionReference.Text
    Else
        m_sQuestionReference = m_clsQuestions.GetNextQuestionRef
        m_clsQuestions.SetQuestionReferenceID m_sQuestionReference
    End If
    
    Set clsTableAccess = m_clsQuestions
    clsTableAccess.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean

    Dim bRet As Boolean
    Dim bShowError As Boolean
    Dim vSaveChanges As Variant
    
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If m_bIsEdit And m_bScreenUpdated And bRet Then
        'This record is an edit, valid and has been amended
        vSaveChanges = MsgBox("You are about to update the selected question. Do you wish to continue?", vbYesNo + vbQuestion, Me.Caption)
         Select Case vSaveChanges
            Case vbYes
                If bRet Then
                    SaveScreenData
                    SaveChangeRequest
                End If
            Case vbNo
                bRet = False
        End Select
    ElseIf Not m_bIsEdit Then
        SaveScreenData
        SaveChangeRequest
        bRet = True
    End If
        
    DoOKProcessing = bRet
    
End Function

Private Sub optDetailsRequired_Click(Index As Integer)
    m_bScreenUpdated = True
End Sub

Private Sub txtQuestionText_change()
    m_bScreenUpdated = True
End Sub
Private Sub txtQuestionShortText_change()
    m_bScreenUpdated = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Common Function used to setup a promotion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sDesc As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    
    sDesc = txtQuestionReference

    colMatchValues.Add m_sQuestionReference
    Set clsTableAccess = m_clsQuestions
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsQuestions, sDesc
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
