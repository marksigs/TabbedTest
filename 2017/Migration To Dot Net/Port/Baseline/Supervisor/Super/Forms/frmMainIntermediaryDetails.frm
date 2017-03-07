VERSION 5.00
Object = "{BCB21120-20F3-4664-94C3-5D74C6E52978}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmMainIntermediaryDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Intermediary Details"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "frmMainIntermediaryDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboType 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSGOCX.MSGEditBox txtIntermediary 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   660
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1260
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtIntermediary 
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSGOCX.MSGEditBox txtIntermediary 
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
      MaxLength       =   10
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Active To"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Active From"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Panel ID"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Intermediary Type"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmMainIntermediaryDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module : IntermediaryDetail
' Description   :
'
' Change history
' Prog      Date        Description
' STB       15/11/2001  Move transactional commencement/termination into the
'                       IntermediaryDetail class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'TextBox Constants
Private Const TEXT_ACTIVE_FROM = 1
Private Const TEXT_ACTIVE_TO = 2
Private Const TEXT_PANEL_ID = 0

Private m_sParentKey As String
Private m_vGUID As Variant
Private m_colKeys As Collection
Private Const m_sIntermediaryTypeCombo As String = "IntermediaryType"
Private m_ReturnCode As MSGReturnCode
Private m_clsIntermediary As IntermediaryTable
Private m_nIntermediaryType As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user clicks the OK button. If DoOKProcessing succeeds (i.e.,
'                   screen is validated and data saved), set the form return code to true and
'                   hide the form thus returning control back to the caller
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed

    Dim bRet As Boolean
    Dim bIsEdit As Boolean
    Dim frmInt As frmEditIntermediaries
    
    bRet = DoOKProcessing
    
    If bRet Then
        SaveScreenData
        Set frmInt = frmEditIntermediaries
        bIsEdit = False
        frmInt.SetIsEdit bIsEdit
        frmInt.SetParentKey m_sParentKey
        frmInt.SetIntermediaryType cboType.ListIndex
        frmInt.SetTableClass m_clsIntermediary
        Hide
'        g_clsDataAccess.BeginTrans
        frmInt.Show vbModal
    
        'Set the return code to that of the child form.
        SetReturnCode frmInt.GetReturnCode
        
        'Unload the edit form.
        Unload frmInt
        
'        If frmInt.GetReturnCode = MSGFailure And Not bIsEdit Then
'            g_clsDataAccess.RollbackTrans
'        End If
    End If
    
    Set frmInt = Nothing
        
    Exit Sub
Failed:
    Set frmInt = Nothing
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user cliks the Cancel button. All we do is hide the form, which
'                   passes control back to the caller, which checks the status of the form, and closes
'                   it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    On Error GoTo Failed

    Hide

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Load()
    
    On Error GoTo Failed

    'Assume failure until data is saved successfully.
    SetReturnCode MSGFailure

    Set m_clsIntermediary = New IntermediaryTable
    Set m_colKeys = New Collection
    
    PopulateScreenControls
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetIntermediaryType(nIntermediaryType As Integer)
    m_nIntermediaryType = nIntermediaryType
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Private Sub PopulateScreenControls()
    On Error GoTo Failed
    
    g_clsFormProcessing.PopulateCombo m_sIntermediaryTypeCombo, cboType, False
    g_clsFormProcessing.HandleComboExtra cboType, m_nIntermediaryType, SET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed

    Dim bRet As Boolean
    Dim clsIntermediary As Intermediary
    Dim colValues As Collection
    Dim sPanelID As String
        
    Set colValues = New Collection
    
    Set clsIntermediary = New Intermediary
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, True)
    
    
    If bRet Then
        sPanelID = txtIntermediary(TEXT_PANEL_ID).Text
        colValues.Add sPanelID
        bRet = Not TableAccess(m_clsIntermediary).DoesRecordExist(colValues, TableAccess(m_clsIntermediary).GetDuplicateKeys)
        If Not bRet Then
            g_clsErrorHandling.DisplayError "Panel ID must be unique"
            txtIntermediary(TEXT_PANEL_ID).SetFocus
        End If
    End If
    
    If bRet Then
        bRet = clsIntermediary.ValidateDate(txtIntermediary(TEXT_ACTIVE_FROM), txtIntermediary(TEXT_ACTIVE_TO))
    End If
    
    
    
    DoOKProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SaveScreenData()
    On Error GoTo Failed

    Dim clsGUID As GuidAssist
    Dim clsIntermediary As IntermediaryTable
    Dim vVal As Variant
    Dim dTmp As Date
    
    Set clsIntermediary = m_clsIntermediary
    Set clsGUID = New GuidAssist
        
    g_clsFormProcessing.CreateNewRecord clsIntermediary
    
    'Get/Set the GUID
    If m_colKeys.Count > 0 Then
        m_colKeys.Remove (1)
    End If
    
    m_vGUID = clsGUID.CreateGUID
    m_colKeys.Add m_vGUID
        
    clsIntermediary.SetIntermediaryGuid m_vGUID
     
    TableAccess(clsIntermediary).SetKeyMatchValues m_colKeys
    
    'PanelID
    clsIntermediary.SetPanelID txtIntermediary(TEXT_PANEL_ID).Text
        
    'Active From
    g_clsFormProcessing.HandleDate txtIntermediary(TEXT_ACTIVE_FROM), dTmp, GET_CONTROL_VALUE
    clsIntermediary.SetActiveFrom dTmp
    
    'Active To
    If Len(txtIntermediary(TEXT_ACTIVE_TO).ClipText) > 0 Then
        g_clsFormProcessing.HandleDate txtIntermediary(TEXT_ACTIVE_TO), dTmp, GET_CONTROL_VALUE
        clsIntermediary.SetActiveTo dTmp
    End If
    
    'InteremdiaryType
    g_clsFormProcessing.HandleComboExtra cboType, vVal, GET_CONTROL_VALUE
    clsIntermediary.SetIntermediaryType CStr(vVal)
    
    'TableAccess(clsIntermediary).Update
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetParentKey(sKey As String)
    m_sParentKey = sKey
End Sub
