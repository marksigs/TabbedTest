VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmFindDepartments 
   Caption         =   "Find Departments"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   Icon            =   "frmFindDepartments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5235
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3945
      TabIndex        =   4
      Top             =   1335
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1335
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtDepartment 
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      TextType        =   4
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
   Begin MSGOCX.MSGComboBox cboChannel 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   300
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2955
      Begin VB.OptionButton optSearch 
         Caption         =   "Search by Department"
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   540
         Width           =   2235
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Search by Channel"
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   60
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmFindDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SEARCH_CHANNEL = 0
Private Const SEARCH_DEPARTMENT = 1
'Private m_clsChannel As DistributionChannelTable
Private m_ReturnCode As MSGReturnCode
Private m_clsDepartment As DepartmentTable
Private m_colKeys As Collection
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsDepartment = clsTableAccess
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    
    If Me.optSearch(SEARCH_CHANNEL).Value = True Then
        DoChannelSearch
    Else
        DoDepartmentSearch
    End If

    Set clsTableAccess = m_clsDepartment
    clsTableAccess.SetUpdated
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub DoChannelSearch()
    On Error GoTo Failed
    Dim sChannel As String
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    ' Make sure something is seleted from the channel combo
    
    'sChannel = Me.cboChannel.SelText
    g_clsFormProcessing.HandleComboExtra cboChannel, sChannel, GET_CONTROL_VALUE
    
    If Len(sChannel) > 0 Then
        m_clsDepartment.GetDepartmentsByChannel sChannel
        Set clsTableAccess = m_clsDepartment
        If clsTableAccess.RecordCount() = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "No Departments found with Channel: " + sChannel
        End If
    Else
        Me.cboChannel.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Channel must be selected from Channel Combo"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoDepartmentSearch()
    On Error GoTo Failed
    Dim sDepartment As String
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    
    ' Make sure something is seleted from the channel combo
    sDepartment = Me.txtDepartment.Text
    
    If Len(sDepartment) > 0 Then
        m_clsDepartment.GetDepartmentsByID sDepartment
        Set clsTableAccess = m_clsDepartment
                            
        If clsTableAccess.RecordCount() = 0 Then
            txtDepartment.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "No Departments  found with DepartmentID: " + sDepartment
        End If
    Else
        txtDepartment.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Department ID must be entered"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    SetReturnCode MSGFailure
    
    ' Default is Department
    optSearch(SEARCH_DEPARTMENT).Value = True

    PopulateChannel
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub PopulateChannel()
    'Set m_clsChannel = New DistributionChannelTable
    g_clsFormProcessing.PopulateChannel Me.cboChannel
End Sub
Private Sub optSearch_Click(Index As Integer)
    Select Case Index
    Case SEARCH_CHANNEL
        EnableChannelSearch
    Case SEARCH_DEPARTMENT
        EnableDepartmentSearch
    End Select
End Sub
Private Sub EnableChannelSearch()
    ' Enable channel combo, disable department text box
    cboChannel.Enabled = True
    txtDepartment.Text = ""
    txtDepartment.Enabled = False
    
    If Me.Visible = True Then
        cboChannel.SetFocus
    End If
End Sub
Private Sub EnableDepartmentSearch()
    ' Enable department text box, disable channel combo
    cboChannel.Enabled = False
    txtDepartment.Enabled = True
    
    If Me.Visible = True Then
        txtDepartment.SetFocus
    End If
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
