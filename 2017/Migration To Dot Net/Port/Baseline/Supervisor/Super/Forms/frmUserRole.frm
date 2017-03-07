VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmUserRole 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Role"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmUserRole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGEditBox TxtActiveTo 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin MSGOCX.MSGEditBox TxtActiveFrom 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin MSGOCX.MSGComboBox cboUserRole 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   3015
      _ExtentX        =   5318
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
   Begin MSGOCX.MSGComboBox cboUnit 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Active To"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Active From"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User Role"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "UnitId"
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmUserRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Private data

Private m_bIsEdit As Boolean
Private m_clsUserRoleTable As UserRoleTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_nCurrentIndex As Long
Private m_colExistingUnits  As Collection
Private m_sUserID As String
Private m_sUnitID As String

Private m_iRecordsAdded As Integer


Private m_bOk As Boolean

Private Sub cmdCancel_Click()
    
    Dim uResponse As VbMsgBoxResult
    
    On Error GoTo Failed

    'If the another button has been used, we should warn that more than one record
    'is about to be lost. Normally this would only rollback the last record added,
    'but this form is a popup which runs in a wider transaction.
    If m_iRecordsAdded > 1 Then
        uResponse = MsgBox("You have added more than one record. Proceeding will cancel all records added. Do you wish to proceed?", vbYesNo Or vbExclamation, "Loose Changes?")
    End If
    
    'If editing, only one record has been added or the user is willing to loose
    'multiple changes, then allow the cancel to proceed,
    If (m_iRecordsAdded < 2) Or (uResponse = vbYes) Then
        If Not m_bIsEdit Then
            CancelCurrentRecord
        End If
    
        Hide
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub




Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Validate data.
    bRet = DoOKProcessing

    If bRet Then
    
       ' If m_bIsEdit = False Then
       '     SetAddState
       ' End If
        
        'Save the control values to the underlying table.
        SaveScreenData
        
        'Indicate to the opener that everything has worked.
        SetReturnCode MSGSuccess
        
        'Hide the form and return control to the opener.
        Hide
    End If

    m_bOk = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Function DoOKProcessing() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, True)
        
    If bRet Then
        bRet = ValidateScreenData()
    End If
    
    DoOKProcessing = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function ValidateScreenData() As Boolean
    
    Dim dTo As Date
    Dim dFrom As Date
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    ValidateScreenData = False
    
'Valdate Dates:
    
    'Get the from date
    If Len(txtActiveFrom.ClipText) = 0 Then
        g_clsErrorHandling.DisplayError "Active From Date must be entered "
        txtActiveFrom.SetFocus
        Exit Function
    End If
    
    If Len(txtActiveFrom.ClipText) > 0 Then
        g_clsFormProcessing.HandleDate txtActiveFrom, dFrom, GET_CONTROL_VALUE
    Else
        dFrom = 0
    End If
    
    'Get the To Date
    If Len(txtActiveTo.ClipText) > 0 Then
        g_clsFormProcessing.HandleDate txtActiveTo, dTo, GET_CONTROL_VALUE
    Else
        dTo = 0
    End If
    
    If dFrom > dTo And dTo > 0 Then
        g_clsErrorHandling.DisplayError "Active To Date must be greater than the Active From Date"
        Exit Function
    End If

'Validate UserRole :
    
    If cboUserRole.Text = "" Then
        g_clsErrorHandling.DisplayError "User Role must be selected"
        cboUserRole.SetFocus
        Exit Function
    End If
    
    ValidateScreenData = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function



Private Sub SaveScreenData()
    
    On Error GoTo Failed
    Dim sTaskID As String
    Dim vVal As Variant
    Dim vStartDate As Variant
    Dim vEndDate As Variant
    
    ' Unit ID
    g_clsFormProcessing.HandleComboText cboUnit, vVal, GET_CONTROL_VALUE
    m_clsUserRoleTable.SetUnitID vVal
    m_sUnitID = CStr(vVal)
    
    g_clsFormProcessing.HandleDate txtActiveFrom, vStartDate, GET_CONTROL_VALUE
    m_clsUserRoleTable.SetActiveFrom vStartDate
    
    g_clsFormProcessing.HandleDate txtActiveTo, vEndDate, GET_CONTROL_VALUE
    m_clsUserRoleTable.SetActiveTo vEndDate
    
    g_clsFormProcessing.HandleComboExtra cboUserRole, vVal, GET_CONTROL_VALUE
    m_clsUserRoleTable.SetRole vVal
    
    g_clsFormProcessing.HandleComboText cboUserRole, vVal, GET_CONTROL_VALUE
    m_clsUserRoleTable.SetRoleText vVal
    
    ' User ID
    m_clsUserRoleTable.SetUserID m_sUserID
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    
    m_iRecordsAdded = 1
    
    Initialise

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub

Public Sub SetEditState()
    
    On Error GoTo Failed

    ' Validate we have the record
    If TableAccess(m_clsUserRoleTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit User Roles - Unable to locate User Role"
    End If
    
    ' If we get here, we have the data we need
    PopulateScreenFields

    'cmdAnother.Enabled = False
     
    cboUnit.Enabled = False
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenFields()

    On Error GoTo Failed

    Dim vVal As Variant

    ' Task
    vVal = m_clsUserRoleTable.GetUnitID()
    g_clsFormProcessing.HandleComboText cboUnit, vVal, SET_CONTROL_VALUE
    
    ' Default Task
    vVal = m_clsUserRoleTable.GetActiveFrom()
    g_clsFormProcessing.HandleDate txtActiveFrom, vVal, SET_CONTROL_VALUE
    
    ' Additional Task
    vVal = m_clsUserRoleTable.GetActiveTo()
    g_clsFormProcessing.HandleDate txtActiveTo, vVal, SET_CONTROL_VALUE
    
    ' Mandatory for this stage?
    vVal = m_clsUserRoleTable.GetRoleText()
    g_clsFormProcessing.HandleComboText cboUserRole, vVal, SET_CONTROL_VALUE
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Form_Initialize()
    'm_bIsEdit = False
    Set m_colExistingUnits = Nothing
End Sub
Private Sub Initialise()
    
    On Error GoTo Failed
    m_ReturnCode = MSGFailure
    
    PopulateCombos
        
    If m_bIsEdit Then
        SetEditState
    Else
       SetAddState
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub
Public Sub SetAddState()
    On Error GoTo Failed

    TableAccess(m_clsUserRoleTable).AddRow

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Public Sub SetTableAccess(clsTableAccess As TableAccess)
    
    On Error GoTo Failed
    Dim sFunctionName As String

    sFunctionName = "SetTableAccess"

    If Not clsTableAccess Is Nothing Then
        Set m_clsUserRoleTable = clsTableAccess
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, " " & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateCombos()
    
    On Error GoTo Failed

    'Unit ID
    PopulateUnitID
    
    'User Role
    g_clsFormProcessing.PopulateCombo "UserRole", cboUserRole
    
    Exit Sub
Failed:

    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateUnitID()
    On Error GoTo Failed
    
    Dim colFields As Collection
    Dim clsUserRoleTable As UserRoleTable
    
    Set clsUserRoleTable = New UserRoleTable
    
    clsUserRoleTable.GetUnits m_colExistingUnits
    
    Set colFields = clsUserRoleTable.GetComboFields()
    
    g_clsFormProcessing.PopulateComboFromTable cboUnit, clsUserRoleTable, colFields
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CancelCurrentRecord()
    On Error GoTo Failed
    Dim rsUserRole As ADODB.Recordset
    
    Set rsUserRole = TableAccess(m_clsUserRoleTable).GetRecordSet()

    If Not rsUserRole Is Nothing Then
        rsUserRole.CancelBatch adAffectCurrent
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetExistingUnits(colExistingUnits As Collection)
    Set m_colExistingUnits = colExistingUnits
End Sub

Public Sub SetUserID(sUserID As String)
    m_sUserID = sUserID
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu Then
        cmdCancel_Click
    End If
End Sub






