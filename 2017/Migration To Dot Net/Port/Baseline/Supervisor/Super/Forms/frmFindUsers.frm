VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmFindUsers 
   Caption         =   "Find Users"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   Icon            =   "frmFindUsers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   1500
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   5790
      TabIndex        =   4
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   5790
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
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
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   540
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   35
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
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
      Caption         =   "First Forename"
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Surname"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "User ID"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   150
      Width           =   1515
   End
End
Attribute VB_Name = "frmFindUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindUsers
' Description   : Finds all User details matching the search criteria
'
' Change history
' Prog      Date        Description
' DJP       27/02/03    Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'   BS     25/03/03     BM0282 Move listview lbltitle.Caption refresh
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Underlying table object.
Private m_clsUsers As OmigaUserTable

'Control indexes.
Private Const USER_ID     As Long = 0
Private Const SURNAME      As Long = 1
Private Const FORENAME As Long = 2


' Private data
Private m_lvResults As MSGListView
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdClear_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()

    Dim nThisControl As Integer
    
    On Error GoTo Failed

    'Clear all items from the listview.
    m_lvResults.ListItems.Clear
    'BS BM0282 25/03/03
    'Reset title as listview is being cleared
    frmMain.lblTitle(constListViewLabel).Caption = frmMain.tvwDB.SelectedItem()
    
    'Clear the contents of the criteria controls.
    For nThisControl = 0 To txtSearch.Count - 1
        txtSearch(nThisControl).Text = ""
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Hide the form and return control to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdSearch_Click
' Description   : Uses the current criteria and populates a list of matching
'                 records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSearch_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Ensure any mandatory criteria have been populated.
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, True)
    
    If bRet Then
        'If enough information has been supplied, then find the matching records.
        FindUsers
        Hide
                
        'BS BM0282 25/03/03
        'Reset title as listview is being refreshed
        frmMain.lblTitle(constListViewLabel).Caption = frmMain.tvwDB.SelectedItem()
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create any required table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    Set m_lvResults = frmMain.lvListView
    'Populate static lists and setup column headers.
    PopulateScreenControls

    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate static lists and setup column headers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()

    On Error GoTo Failed
    
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindUsers
' Description   : Returns all matching User records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindUsers()
    On Error GoTo Failed
    Dim clsUsers As OmigaUserTable
    Dim bFound As Boolean
    Dim sForename As String
    Dim sSurname As String
    Dim sUserID As String
    
    sForename = txtSearch(FORENAME).Text
    If sForename = "*" Then
        txtSearch(FORENAME).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "You must have at least one character to perform a wildcard search"
    End If
    
    sUserID = txtSearch(USER_ID).Text
    
    If sUserID = "*" Then
        txtSearch(USER_ID).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "You must have at least one character to perform a wildcard search"
    End If
    
    sSurname = txtSearch(SURNAME).Text
    
    If Len(sForename) = 0 And Len(sSurname) = 0 And Len(sUserID) = 0 Then
        txtSearch(USER_ID).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "At least one of User ID, Surname and First Forename must be entered"
    End If
    
    Set clsUsers = New OmigaUserTable
    
    clsUsers.GetUsers sUserID, sSurname, sForename
    
    If TableAccess(clsUsers).RecordCount > 0 Then
        g_clsFormProcessing.PopulateFromRecordset m_lvResults, clsUsers
    Else
        Set clsUsers = Nothing
        g_clsErrorHandling.RaiseError errGeneralError, "No records found for your search criteria"
    End If
    
    Set clsUsers = Nothing
    
    Exit Sub
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsUsers = Nothing
End Sub
Public Sub SetTableClass(clsUsers As OmigaUserTable)
    Set m_clsUsers = clsUsers
End Sub
