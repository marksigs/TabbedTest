VERSION 5.00

Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"

Begin VB.Form frmFindUnits 
   Caption         =   "Find Units"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   Icon            =   "frmFindUnits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGDataCombo cboDepartment 
      Height          =   315
      Left            =   3540
      TabIndex        =   12
      Top             =   420
      Width           =   1635
      _ExtentX        =   2884
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
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   300
      TabIndex        =   4
      Top             =   240
      Width           =   2955
      Begin VB.OptionButton optSearchType 
         Caption         =   "Search by Unit"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "Search by Region"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton optSearchType 
         Caption         =   "Search by Department"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   180
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&OK"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1860
         Width           =   1215
      End
      Begin MSGOCX.MSGEditBox MSGEditBox2 
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   900
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSGOCX.MSGComboBox MSGComboBox2 
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         Top             =   480
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4500
      TabIndex        =   3
      Top             =   1755
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtUnit 
      Height          =   315
      Left            =   3540
      TabIndex        =   1
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSGOCX.MSGComboBox cboRegion 
      Height          =   315
      Left            =   3540
      TabIndex        =   0
      Top             =   780
      Width           =   1695
      _ExtentX        =   2990
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
      Text            =   ""
   End
End
Attribute VB_Name = "frmFindUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SEARCH_DEPARTMENT = 0
Private Const SEARCH_REGION = 1
Private Const SEARCH_UNIT = 2
Private m_ReturnCode As MSGReturnCode
Private m_clsUnit As UnitTable
Private m_colKeys As Collection
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsUnit = clsTableAccess
End Sub
Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsTableAccess As TableAccess
    
    If optSearchType(SEARCH_DEPARTMENT).Value = True Then
        DoDepartmentSearch
    ElseIf optSearchType(SEARCH_REGION).Value = True Then
        DoRegionSearch
    ElseIf optSearchType(SEARCH_UNIT).Value = True Then
        DoUnitSearch
    End If

    Set clsTableAccess = m_clsUnit
    clsTableAccess.SetUpdated
    Hide

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub
Private Sub DoRegionSearch()
    On Error GoTo Failed
    Dim sRegion As String
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    ' Make sure something is seleted from the channel combo
    
    g_clsFormProcessing.HandleComboExtra Me.cboRegion, sRegion, GET_CONTROL_VALUE
    
    If Len(sRegion) > 0 Then
        m_clsUnit.GetUnitsByRegion sRegion
        Set clsTableAccess = m_clsUnit
        
        If clsTableAccess.RecordCount() = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "No Units found with Region: " + sRegion
        End If
    Else
        Me.cboRegion.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Region must be selected from Region Combo"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub DoUnitSearch()
    On Error GoTo Failed
    Dim sUnit As String
    Dim colValues As New Collection
    Dim clsTableAccess As TableAccess
    ' Make sure something is seleted from the channel combo
    sUnit = txtUnit.Text
    
    If Len(sUnit) > 0 Then
        m_clsUnit.GetUnitsByUnit sUnit
        Set clsTableAccess = m_clsUnit
        
        If clsTableAccess.RecordCount() = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, "No Units found with Unit: " + sUnit
        End If
    Else
        txtUnit.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Unit must be entered"
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
    sDepartment = cboDepartment.SelText
    
    If Len(sDepartment) > 0 Then
        m_clsUnit.GetUnitsByDepartment sDepartment
        Set clsTableAccess = m_clsUnit
        
        If clsTableAccess.RecordCount() = 0 Then
            cboDepartment.SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "No Units found with DepartmentID: " + sDepartment
        End If
    Else
        cboDepartment.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Department ID must be selected"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    
    SetReturnCode MSGFailure

    PopulateDepartment
    PopulateRegion
    
    optSearchType(SEARCH_DEPARTMENT).Value = True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub PopulateDepartment()
    Dim clsDepartment As New DepartmentTable
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim sDepartmentField As String
    
    Set clsTableAccess = clsDepartment
    
    Set rs = clsTableAccess.GetTableData()

    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Set cboDepartment.RowSource = rs
            sDepartmentField = clsDepartment.GetDepartmentField()
            cboDepartment.ListField = sDepartmentField
        Else
            g_clsErrorHandling.RaiseError errRecordNotFound, "Departments"
        End If
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, "Departments"
    End If
End Sub
Private Sub PopulateRegion()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim clsRegion As New RegionTable
    Dim clsTableAccess As TableAccess
    Dim colValues As New Collection
    Dim colIDS As New Collection
    
    Set clsTableAccess = clsRegion
    
    clsRegion.GetRegionsAsCollection colValues, colIDS

    cboRegion.SetListTextFromCollection colValues, colIDS
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub EnableDepartmentSearch(Optional bEnable As Boolean = True)
    ' Enable department text box, disable channel combo
    cboDepartment.Enabled = bEnable
    cboDepartment.SelText = ""
    
    If Me.Visible = True And bEnable = True Then
        cboDepartment.SetFocus
    End If
End Sub
Private Sub EnableRegionSearch(Optional bEnable As Boolean = True)
    cboRegion.Enabled = bEnable
    cboRegion.ListIndex = -1
    
    If Me.Visible = True And bEnable = True Then
        cboRegion.SetFocus
    End If
End Sub
Private Sub EnableUnitSearch(Optional bEnable As Boolean = True)
    ' Enable department text box, disable channel combo
    txtUnit.Enabled = bEnable
    txtUnit.Text = ""
    
    If Me.Visible = True And bEnable = True Then
        txtUnit.SetFocus
    End If
End Sub
Private Sub optSearchType_Click(Index As Integer)
    Select Case Index
    Case SEARCH_DEPARTMENT
        EnableDepartmentSearch
        EnableRegionSearch False
        EnableUnitSearch False
    
    Case SEARCH_REGION
        EnableRegionSearch
        EnableDepartmentSearch False
        EnableUnitSearch False

    Case SEARCH_UNIT
        EnableUnitSearch
        EnableRegionSearch False
        EnableDepartmentSearch False
    End Select
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

