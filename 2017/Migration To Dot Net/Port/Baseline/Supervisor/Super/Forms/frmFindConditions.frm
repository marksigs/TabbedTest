VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmFindConditions 
   Caption         =   "Find Conditions"
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "frmFindConditions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6810
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboChannelID 
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Top             =   1350
      Width           =   2865
      _ExtentX        =   5054
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   4890
      TabIndex        =   4
      Top             =   450
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   4890
      TabIndex        =   5
      Top             =   930
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   1860
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   0
      Top             =   60
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
      Left            =   1740
      TabIndex        =   1
      Top             =   480
      Width           =   2865
      _ExtentX        =   5054
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
      MaxLength       =   50
   End
   Begin MSGOCX.MSGComboBox cboConditionType 
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Top             =   900
      Width           =   2865
      _ExtentX        =   5054
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
   Begin VB.Label lblSearch 
      Caption         =   "Condition Type"
      Height          =   315
      Index           =   3
      Left            =   90
      TabIndex        =   10
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Condition Reference"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   510
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Channel ID"
      Height          =   315
      Index           =   2
      Left            =   90
      TabIndex        =   7
      Top             =   1350
      Width           =   1515
   End
End
Attribute VB_Name = "frmFindConditions"
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
Private m_clsConditions As ConditionsTable

'Control indexes.
Private Const CONDITION_REF     As Long = 0
Private Const CONDITION_NAME      As Long = 1

' Private data
Private m_lvResults As MSGListView
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdClear_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()

    'BS BM0282 25/03/03
    'Clear all items from the listview (to be consistent with other options)
    m_lvResults.ListItems.Clear
    'Reset title as listview is being cleared
    frmMain.lblTitle(constListViewLabel).Caption = frmMain.tvwDB.SelectedItem()

    g_clsFormProcessing.ClearScreenFields Me
    
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
        FindConditions
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
    
    PopulateCombos
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateCombos()
    On Error GoTo Failed
    
    ' First the channel combo
    g_clsFormProcessing.PopulateChannel cboChannelID
    
    ' Then the Condition Type combo
    g_clsFormProcessing.PopulateCombo "CONDITIONTYPE", cboConditionType
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindUsers
' Description   : Returns all matching User records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindConditions()
    On Error GoTo Failed
    Dim clsConditions As ConditionsTable
    Dim bFound As Boolean
    Dim sType As String
    Dim sChannel As String
    Dim sConditionRef As String
    Dim sConditionName As String
    
    sConditionRef = txtSearch(CONDITION_REF).Text
    
    If sConditionRef = "*" Then
        txtSearch(CONDITION_REF).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "You must have at least one character to perform a wildcard search"
    End If
    
    If Len(sConditionRef) > 0 Then
        sConditionRef = g_clsSQLAssistSP.FormatWildcardedString(sConditionRef, bFound)
    End If
    
    sConditionName = txtSearch(CONDITION_NAME).Text
    
    If Len(sConditionName) > 0 Then
        sConditionName = g_clsSQLAssistSP.FormatWildcardedString(sConditionName, bFound)
    End If
    
    ' Get the Condition Type
    g_clsFormProcessing.HandleComboExtra cboConditionType, sType, GET_CONTROL_VALUE
    
    ' Get the Channel
    g_clsFormProcessing.HandleComboExtra cboChannelID, sChannel, GET_CONTROL_VALUE
    
    If Len(sChannel) > 0 Then
        sChannel = g_clsSQLAssistSP.FormatWildcardedString(sChannel, bFound)
    End If
    
    If Len(sConditionName) = 0 And Len(sConditionRef) = 0 And Len(sType) = 0 And Len(sChannel) = 0 Then
        txtSearch(CONDITION_REF).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "At least one of Condition Reference, Name, Type and Channel ID must be entered"
    End If
    
    Set clsConditions = New ConditionsTable
    
    clsConditions.GetConditions sConditionRef, sConditionName, sType, sChannel
    
    If TableAccess(clsConditions).RecordCount > 0 Then
        g_clsFormProcessing.PopulateFromRecordset m_lvResults, clsConditions
    Else
        Set clsConditions = Nothing
        g_clsErrorHandling.RaiseError errGeneralError, "No records found for your search criteria"
    End If
    
    Set clsConditions = Nothing
    
    Exit Sub
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsConditions = Nothing
End Sub
Public Sub SetTableClass(clsConditions As ConditionsTable)
    Set m_clsConditions = clsConditions
End Sub

