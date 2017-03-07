VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditRegion 
   Caption         =   "Add/Edit Region"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmEditRegion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   4815
      TabIndex        =   5
      Top             =   1575
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtRegion 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   180
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      TextType        =   5
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
   Begin MSGOCX.MSGEditBox txtRegion 
      Height          =   315
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   1020
      Width           =   3975
      _ExtentX        =   7011
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
      MaxLength       =   30
   End
   Begin MSGOCX.MSGComboBox cboRegionType 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      Enabled         =   -1  'True
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
   Begin VB.Label lblLenderDetails 
      Caption         =   "Region Name"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "Region ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblLenderDetails 
      Caption         =   "RegionType"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   660
      Width           =   1395
   End
End
Attribute VB_Name = "frmEditRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditRegion
' Description   :
'
' Change history
' Prog      Date        AQR     Description
' STB       06/12/01    SYS1942 - Another button commits current transaction.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Private Const MAX_SHORT = 32767
Private Const REGION_ID = 0
Private Const REGION_NAME = 1
Private m_ReturnCode As MSGReturnCode
Private m_bIsEdit As Boolean
Private m_clsRegion As RegionTable
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

'Public Sub SetTableClass(clsTableAccess As TableAccess)
'    Set m_clsRegion = clsTableAccess
'End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function

Private Sub cboRegionType_Validate(Cancel As Boolean)
    Cancel = Not cboRegionType.ValidateData()
End Sub

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        bRet = g_clsFormProcessing.ClearScreenFields(Me)
    End If

    If bRet = True Then
        txtRegion(REGION_ID).SetFocus
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
            
        End If
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function
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
Private Function ValidateScreenData() As Boolean
    On Error GoTo Failed
    Dim nCount As Integer
    Dim bRet As Boolean

    bRet = True
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If

    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function DoesRecordExist()
    Dim bRet As Boolean
    Dim sRegion As String
    Dim clsRegion As New RegionTable
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection

    sRegion = Me.txtRegion(REGION_ID).Text

    If Len(sRegion) > 0 Then
        Set clsTableAccess = clsRegion
        colValues.Add sRegion
        
        bRet = clsTableAccess.DoesRecordExist(colValues)
    
        If bRet = True Then
            MsgBox "This region already exists"
            txtRegion(REGION_ID).SetFocus
        End If
    End If

    DoesRecordExist = bRet
End Function

Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    Set m_clsRegion = New RegionTable
    
    g_clsFormProcessing.PopulateCombo "RegionType", Me.cboRegionType
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Public Sub SetAddState()

End Sub

Public Sub SetEditState()
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim sGUID As Variant
    Dim sDepartmentID As String
    Dim colValues As New Collection

    ' First, the department table
    Set clsTableAccess = m_clsRegion
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData()
    cmdAnother.Enabled = False

    If Not rs Is Nothing Then
        If rs.RecordCount >= 0 Then
            PopulateScreenFields
        Else
            MsgBox "Region - no records to edit"
        End If
    End If
End Sub
Private Sub txtRegion_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtRegion(Index).ValidateData()
    
    If Cancel = False And Index = REGION_ID Then
        Dim sRegion As String
        sRegion = txtRegion(Index).Text
            
        If Len(sRegion) > 0 Then
            If CLng(sRegion) > MAX_SHORT Then
                MsgBox "Region ID must be between 0 and " & MAX_SHORT
                Cancel = True
            End If
        End If
    End If

End Sub
Public Function PopulateScreenFields() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vTmp As Variant
    
    ' Region ID
    txtRegion(REGION_ID).Text = m_clsRegion.GetRegionID()
    ' Region Name
    txtRegion(REGION_NAME).Text = m_clsRegion.GetRegionName()
    ' Region Type
    vTmp = m_clsRegion.GetRegionType()
    'g_clsFormProcessing.HandleComboText Me.cboRegionType, vTmp, SET_CONTROL_VALUE
    g_clsFormProcessing.HandleComboExtra Me.cboRegionType, vTmp, SET_CONTROL_VALUE
    
    PopulateScreenFields = True
    Exit Function
Failed:
    PopulateScreenFields = False
End Function
Public Sub SaveScreenData()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    
    If m_bIsEdit = False Then
        g_clsFormProcessing.CreateNewRecord m_clsRegion
    End If
    
    ' Region ID
    m_clsRegion.SetRegionID txtRegion(REGION_ID).Text
    
    ' Region Name
    m_clsRegion.SetRegionName txtRegion(REGION_NAME).Text
    
    ' Region Type
    g_clsFormProcessing.HandleComboExtra Me.cboRegionType, vTmp, GET_CONTROL_VALUE
    m_clsRegion.SetRegionType CStr(vTmp)
    Set clsTableAccess = m_clsRegion
    clsTableAccess.Update
    
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
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add txtRegion(REGION_ID).Text
    Set clsTableAccess = m_clsRegion
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    g_clsHandleUpdates.SaveChangeRequest m_clsRegion
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
