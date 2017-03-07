VERSION 5.00
Begin VB.Form frmEditPaymentForCompletions 
   Caption         =   "Payment for Completions Viewer/Editor"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmEditPaymentForCompletions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Manual Completion?"
      Height          =   855
      Left            =   4560
      TabIndex        =   14
      Top             =   1080
      Width           =   2655
      Begin VB.OptionButton optManualCompletion 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optManualCompletion 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exclude?"
      Height          =   855
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton optExclude 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optExclude 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label ApplicationType 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Application Type"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label CompletionModifiedBy 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label CompletionModifiedDateTime 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label ApplicationNumber 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Modified by"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Modified on"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Application Number"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "frmEditPaymentForCompletions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditPaymentForCompletions
' Description   : Form which allows the editing/viewing of Payments for Completions
'
' Epsom Change history
' Prog      Date        Description
' TW        27/11/2007  Added Form
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_clsDisbursementPaymentTable As DisbursementpaymentTable
Private m_colKeys As Collection

Private Sub cmdCancel_Click()
    Hide
End Sub


Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    SaveScreenData
    
    SetReturnCode
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Private Sub SaveScreenData()
Dim vTmp As Variant
Dim clsTableAccess As TableAccess
Dim sBatchNo As String
Dim dRun As Date
Dim lStatus As Long

    On Error GoTo Failed
    Set clsTableAccess = m_clsDisbursementPaymentTable
    
    m_clsDisbursementPaymentTable.SetCompletionmodifiedby g_sSupervisorUser
    m_clsDisbursementPaymentTable.SetCompletionmodifieddatetime Now
    If optExclude(0).Value = True Then
        m_clsDisbursementPaymentTable.SetExcludefromautocompletionind 1
    Else
        m_clsDisbursementPaymentTable.SetExcludefromautocompletionind 0
    End If
    If optManualCompletion(0) = True Then
        m_clsDisbursementPaymentTable.SetManualinterfaceind 1
    Else
        m_clsDisbursementPaymentTable.SetManualinterfaceind 0
    End If
    clsTableAccess.Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub Form_Load()
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New DisbursementpaymentTable
    Set m_clsDisbursementPaymentTable = New DisbursementpaymentTable
    
    SetEditState
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing Payments for Completions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
Dim colDataSet As New Collection
    
    ApplicationNumber.Caption = m_colKeys(1)
    
    colDataSet.Add m_colKeys.Item(1)
    colDataSet.Add m_colKeys.Item(2)
    TableAccess(m_clsDisbursementPaymentTable).SetKeyMatchValues colDataSet
    TableAccess(m_clsDisbursementPaymentTable).GetTableData
    
    ApplicationNumber.Caption = m_clsDisbursementPaymentTable.GetApplicationNumber()
    ApplicationType.Caption = frmMain.lvListView.getValueFromName("Application Type")
    CompletionModifiedDateTime.Caption = m_clsDisbursementPaymentTable.GetCompletionmodifieddatetime()
    CompletionModifiedBy.Caption = m_clsDisbursementPaymentTable.GetCompletionmodifiedby()
    If Val(m_clsDisbursementPaymentTable.GetExcludefromautocompletionind()) = 1 Then
        optExclude(0).Value = True
    Else
        optExclude(1).Value = True
    End If
    If Val(m_clsDisbursementPaymentTable.GetManualinterfaceind()) = 1 Then
        optManualCompletion(0).Value = True
    Else
        optManualCompletion(1).Value = True
    End If
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub



Private Sub optExclude_Click(Index As Integer)
    If Index = 0 Then
        optManualCompletion(1).Value = True   'If "Exclude" = Yes then set "Manual Completion" = No
    End If
    optManualCompletion(0).Enabled = (Index = 1)
    optManualCompletion(1).Enabled = (Index = 1)
End Sub


