VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmErrors 
   Caption         =   "Error Message Add/Edit"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5055
      TabIndex        =   4
      Top             =   3075
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3735
      TabIndex        =   3
      Top             =   3075
      Width           =   1215
   End
   Begin MSGOCX.MSGTextMulti txtMessageText 
      Height          =   1395
      Left            =   2040
      TabIndex        =   2
      Top             =   1020
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2461
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
   Begin MSGOCX.MSGEditBox txtMessageNumber 
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
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
   Begin MSGOCX.MSGComboBox cboMessageType 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1755
      _ExtentX        =   3096
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
      ListText        =   "Warning|Error|Information"
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Message Type"
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label Label3 
      Caption         =   "Message Text"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label lblErrorCode 
      Caption         =   "Error Code"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   1755
   End
End
Attribute VB_Name = "frmErrors"
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
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Private Sub cboMessageType_Validate(Cancel As Boolean)
    Cancel = Not cboMessageType.ValidateData()
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
'Public Sub SetTableClass(clsTable As TableAccess)
'    Set m_clsTableAccess = clsTable
'End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user pressed the OK button. Performs necessary
'                   validation and saves any data that needs to be saved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
            SetReturnCode
            Hide
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
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
Private Function ValidateScreenData() As Boolean
    Dim bRet As Boolean

    bRet = True

    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If
    
    ValidateScreenData = bRet
End Function
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsTableAccess = New ErrorMessageTable
    
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
Public Sub SetAddState()
    m_clsTableAccess.AddRow
End Sub
Public Sub SetEditState()
    m_clsTableAccess.SetKeyMatchValues m_colKeys
End Sub
Private Sub txtDescription_GotFocus()
End Sub
Private Sub txtMessageNumber_Validate(Cancel As Boolean)
    Cancel = Not txtMessageNumber.ValidateData()
End Sub
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
    Dim clsErrorTable  As ErrorMessageTable
    Set clsErrorTable = m_clsTableAccess
    
    'txtMessageNumber.Text = clsErrorTable.GetMessageNumber()
    clsErrorTable.GetMessageNumber txtMessageNumber, lblErrorCode
    txtMessageText.Text = clsErrorTable.GetMessageText()
    g_clsFormProcessing.HandleComboText cboMessageType, clsErrorTable.GetMessageType(), SET_CONTROL_VALUE

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
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
Private Function DoesRecordExist() As Boolean
    Dim bRet As Boolean
    Dim sMessageNumber As String
    
    Dim col As New Collection
    
    sMessageNumber = txtMessageNumber.Text
    
    col.Add sMessageNumber
    
    bRet = m_clsTableAccess.DoesRecordExist(col)
    
    If bRet = True Then
        MsgBox "Message already exists - please enter a unique Message Number", vbInformation
        txtMessageNumber.SetFocus
    End If
    DoesRecordExist = bRet
End Function

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
