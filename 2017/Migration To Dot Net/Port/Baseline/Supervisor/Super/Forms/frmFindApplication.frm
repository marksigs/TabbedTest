VERSION 5.00
Object = "{F2CAAEAC-D281-11D4-8274-000102A316E5}#5.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindApplication 
   Caption         =   "Find Application"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "frmFindApplication.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameSearch 
      Caption         =   "Search Criteria"
      Height          =   2715
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   4935
      Begin MSGOCX.MSGEditBox txtFindApp 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1875
         _ExtentX        =   3307
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
         MaxLength       =   12
      End
      Begin MSGOCX.MSGEditBox txtFindApp 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
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
         MaxLength       =   40
      End
      Begin MSGOCX.MSGEditBox txtFindApp 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Top             =   780
         Width           =   2775
         _ExtentX        =   4895
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
         MaxLength       =   30
      End
      Begin VB.Label lblSearch 
         Caption         =   "Surname"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label lblSearch 
         Caption         =   "Application Number"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label lblSearch 
         Caption         =   "Forename"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options"
      Height          =   1035
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   4935
      Begin VB.OptionButton optSearch 
         Caption         =   "Enter Search Criteria"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   660
         Width           =   1815
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Find All Applications"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmFindApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_clsTableAccess As ApplicationTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private Const FIND_APPLICATION As Integer = 0
Private Const FIND_FORENAME As Integer = 1
Private Const FIND_SURNAME As Integer = 2
Private Const SEARCH_ALL As Boolean = 0
Public Sub SetTableClass(clsTableAccess As TableAccess)
    Set m_clsTableAccess = clsTableAccess
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Private Sub cmdCancel_Click()
    Hide
    Unload Me
End Sub
Private Sub cmdFind_Click()
    On Error GoTo Failed
    BeginWaitCursor
    ValidateScreenData
    
    DoSearch
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub ValidateScreenData()
    On Error GoTo Failed
        
    Dim sForename As String
    Dim sSurname As String
    Dim sAppNo As String
        
    If optSearch(SEARCH_ALL).Value = False Then
        sForename = txtFindApp(FIND_FORENAME).Text
        sSurname = txtFindApp(FIND_SURNAME).Text
        sAppNo = txtFindApp(FIND_APPLICATION).Text
        
        If Len(sForename) = 0 And Len(sSurname) = 0 And Len(sAppNo) = 0 Then
            txtFindApp(FIND_APPLICATION).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "At least one search criteria must be entered"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Load()
    SetReturnCode MSGFailure
    
    optSearch(0).Value = True
    
    EndWaitCursor
End Sub
Private Sub DoSearch()
    On Error GoTo Failed
    Dim sAppNo  As String
    Dim sSurname As String
    Dim sForename As String
    Dim rs As ADODB.Recordset
    Dim nResponse As Integer
    Dim bSearch As Boolean
    
    bSearch = True
    
    If optSearch(SEARCH_ALL).Value = True Then
        nResponse = MsgBox("Searching all applications may take some time. Do you want to continue?", vbYesNo + vbQuestion)
    
        If nResponse = vbNo Then
            bSearch = False
        End If
    End If
            
    If bSearch Then
        ' Get the application/customer number
        
        sAppNo = txtFindApp(FIND_APPLICATION).Text
        sSurname = txtFindApp(FIND_SURNAME).Text
        sForename = txtFindApp(FIND_FORENAME).Text
        
        If Not m_clsTableAccess Is Nothing Then
            g_clsDataAccess.SaveSearcSQL True
            m_clsTableAccess.GetActiveApplications sAppNo, sSurname, sForename
            
            TableAccess(m_clsTableAccess).ValidateData
            frmMain.lvListView.MultiSelect = False
            g_clsFormProcessing.PopulateFromRecordset frmMain.lvListView, m_clsTableAccess
            SetReturnCode
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "FindApplication: Table class is empty"
        End If
    End If
    
    EndWaitCursor

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
Private Sub optSearch_Click(Index As Integer)
    On Error GoTo Failed
    If Index = SEARCH_ALL Then
        ' Search all
        EnableSearch False
    Else
        ' Use search criteria
        EnableSearch
        txtFindApp(FIND_APPLICATION).SetFocus
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub EnableSearch(Optional bEnable As Boolean = True)
    On Error GoTo Failed
    
    frameSearch.Enabled = bEnable
    
    
    If txtFindApp(FIND_APPLICATION).Enabled = True And bEnable = False Then
        txtFindApp(FIND_APPLICATION).Text = ""
        txtFindApp(FIND_SURNAME).Text = ""
        txtFindApp(FIND_FORENAME).Text = ""
    End If
    
    txtFindApp(FIND_APPLICATION).Enabled = bEnable
    txtFindApp(FIND_SURNAME).Enabled = bEnable
    txtFindApp(FIND_FORENAME).Enabled = bEnable

    lblSearch(FIND_APPLICATION).Enabled = bEnable
    lblSearch(FIND_SURNAME).Enabled = bEnable
    lblSearch(FIND_FORENAME).Enabled = bEnable


    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


