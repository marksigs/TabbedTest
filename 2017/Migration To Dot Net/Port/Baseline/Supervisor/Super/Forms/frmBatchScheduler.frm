VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{39D4636F-F04D-11D4-82A7-000102A3240F}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmBatchScheduler 
   Caption         =   "Batch Scheduler"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "frmBatchScheduler.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SetParameters...."
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin MSGOCX.MSGComboBox cboNoOfRetries 
      Height          =   315
      Left            =   5640
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSGOCX.MSGComboBox cboInterval 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
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
   Begin MSComCtl2.DTPicker DTPRunTime 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58916866
      CurrentDate     =   36910
   End
   Begin MSComCtl2.DTPicker DTPRunDate 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   36910
   End
   Begin MSGOCX.MSGEditBox txtBatchNumber 
      Height          =   315
      Left            =   5040
      TabIndex        =   6
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSGOCX.MSGEditBox txtDescription 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
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
   Begin MSGOCX.MSGComboBox cboProgramType 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "No Of Retries"
      Height          =   195
      Left            =   4200
      TabIndex        =   12
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Interval"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "RunTime"
      Height          =   195
      Left            =   4200
      TabIndex        =   8
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Run Date"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Program Type"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frmBatchScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bIsEdit As Boolean
Private m_clsTableAccess As TableAccess
Private Const MAX_ERR_NUM = 99999
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection


Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Public Sub PopulateScreen()
    On Error GoTo Failed
    Dim colTables As New Collection

    colTables.Add m_clsTableAccess

    g_clsFormProcessing.PopulateDBScreen colTables
    PopulateScreenFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim clsBatchJobsTable  As BatchJobsTable
    Set clsBatchJobsTable = m_clsTableAccess
    
    txtBatchNumber.Text = clsBatchJobsTable.GetBatchNo()
    txtDescription.Text = clsBatchJobsTable.GetDescription()
    DTPRunDate.Value = clsBatchJobsTable.GetRunDate()

    
    
    
    'txtMessageNumber.Text = clsErrorTable.GetMessageNumber()
    'clsErrorTable.GetMessageNumber txtMessageNumber, lblErrorCode
    'txtMessageText.Text = clsErrorTable.GetMessageText()
    'g_clsFormProcessing.HandleComboText cboMessageType, clsErrorTable.GetMessageType(), SET_CONTROL_VALUE

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add Me.txtMessageNumber.Text
    m_clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetAddState()
    m_clsTableAccess.AddRow
End Sub

Public Sub SetEditState()
    m_clsTableAccess.SetKeyMatchValues m_colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean

    bRet = True

    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    ValidateScreenData = bRet
End Function

Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As String
    Dim clsParam As ErrorMessageTable
    
    Set clsParam = m_clsTableAccess
    clsParam.SetMessageNumber txtMessageNumber.Text
    clsParam.SetMessageText txtMessageText.Text
    
    g_clsFormProcessing.HandleComboText cboMessageType, vTmp, GET_CONTROL_VALUE
    clsParam.SetMessageType vTmp
    m_clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub


Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsTableAccess = New BatchJobsTable
    
    If m_bIsEdit = True Then
        SetEditState
        PopulateScreen
    Else
        SetAddState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

