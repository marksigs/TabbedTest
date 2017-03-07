VERSION 5.00
Object = "{CA498F9F-7CC1-4596-BF1A-B795DCE7BBC6}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindPanel 
   Caption         =   "Find Third Parties"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   Icon            =   "frmFindPanel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Find"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2070
      Width           =   1095
   End
   Begin MSGOCX.MSGComboBox cboThirdPartyType 
      Height          =   315
      Left            =   1740
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
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
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   570
      Width           =   3435
      _ExtentX        =   6059
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
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   2
      Top             =   990
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
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   3
      Top             =   1410
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
      MaxLength       =   40
   End
   Begin VB.Label lblSearch 
      Caption         =   "Third Party Type"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Company Name"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   570
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Panel Id"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Label lblSearch 
      Caption         =   "Town"
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1410
      Width           =   1515
   End
End
Attribute VB_Name = "frmFindPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindPanel
' Description   : Finds all Panel details matching the search criteria
'
' Change history
' Prog      Date        Description
' DJP       18/02/03    Created
' DJP       24/02/03    Updated for spec change - update main listview instead of
'                       search listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'   BS     25/03/03     BM0282 Move listview lbltitle.Caption refresh
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Specific History:
'
' TK       30/11/2005  MAR81 Solicitor Panel Maintenance
'
'-------------------------------------------------------------------------------
Option Explicit

'Underlying table object.
Private m_clsPanel As PanelTable

'Control indexes.
Private Const PANEL_COMPANY_NAME     As Long = 0
Private Const PANEL_ID      As Long = 1
Private Const PANEL_TOWN As Long = 2

' Combo constants - these are hard coded in the database, so are here too (there are no validation types for these)
Private Const COMBO_TYPE_VALUER = 11
Private Const COMBO_TYPE_LEGAL_REP = 10

' Private data
Private m_enumThirdPartyType As ThirdPartyType
Private m_lvResults As MSGListView
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetListViewHeaders
' Description   :   Sets the field header titles on the listview
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetListViewHeaders(enumType As ThirdPartyType)
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim nWidth As Long
    m_lvResults.ColumnHeaders.Clear

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Company Name"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Active From"
    headers.Add lvHeaders

    lvHeaders.nWidth = 15
    lvHeaders.sName = "Active To"
    headers.Add lvHeaders

    lvHeaders.nWidth = 25
    lvHeaders.sName = "Name and Address Type"
    headers.Add lvHeaders

    ' Only display the Panel ID if we're a Legal Rep or a Valuer
    If enumType = COMBO_TYPE_LEGAL_REP Or enumType = COMBO_TYPE_VALUER Then
        lvHeaders.nWidth = 15
        lvHeaders.sName = "Panel ID"
        headers.Add lvHeaders
    End If

    ' TK 30/11/2005 MAR81 Only display the Status if we're a Legal Rep
    If enumType = COMBO_TYPE_LEGAL_REP Then
        lvHeaders.nWidth = 10
        lvHeaders.sName = "Legal Rep Status"
        headers.Add lvHeaders
    End If

    ' Only display the Status if we're a Legal Rep
    If enumType = COMBO_TYPE_LEGAL_REP Then
        lvHeaders.nWidth = 0
        lvHeaders.sName = "Panel UserId"
        headers.Add lvHeaders
    End If

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Town"
    headers.Add lvHeaders

    m_lvResults.AddHeadings headers
    'lvResults.LoadColumnDetails TypeName(Me)

    m_lvResults.HideSelection = False

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Function      : cmdEdit_Click
'' Description   : Opens the intermediary form in edit mode.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub cmdEdit_Click()
'
'    On Error GoTo Failed
'
'    EditPanel
'
'    Exit Sub
'
'Failed:
'    g_clsErrorHandling.DisplayError
'End Sub


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
        FindPanel
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

    SetComboState
    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub SetComboState()
    On Error GoTo Failed
    Dim bEnable As Boolean
    Dim nComboDefault As Long
        
    bEnable = False
    
    'Select the first combo item.
    If m_enumThirdPartyType = ThirdPartyLegalRepType Then
        nComboDefault = COMBO_TYPE_LEGAL_REP
    ElseIf m_enumThirdPartyType = ThirdPartyValuersType Then
        nComboDefault = COMBO_TYPE_VALUER
    Else
        bEnable = True
        nComboDefault = 0
    End If
    
    g_clsFormProcessing.HandleComboExtra cboThirdPartyType, nComboDefault, SET_CONTROL_VALUE
    cboThirdPartyType.Enabled = bEnable

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate static lists and setup column headers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()

    On Error GoTo Failed
    
    'Third Party Type combo.
    g_clsFormProcessing.PopulateCombo "THIRDPARTYTYPE", cboThirdPartyType
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindPanel
' Description   : Returns all matching Panel records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindPanel()
    On Error GoTo Failed
    Dim clsPanel As PanelTable
    Dim vIndex As Variant
    Dim bFound As Boolean
    Dim sTown As String
    Dim sPanelID As String
    Dim sCompanyName As String
    Dim enumSearchType As ThirdPartyType
    g_clsFormProcessing.HandleComboExtra cboThirdPartyType, vIndex, GET_CONTROL_VALUE
    
    If vIndex = "" Then
        cboThirdPartyType.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "A Third Party type must be selected"
    End If
    
    sTown = txtSearch(PANEL_TOWN).Text
    If sTown = "*" Then
        txtSearch(PANEL_TOWN).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "You must have at least one character to perform a wildcard search"
    End If
    
    If Len(sTown) > 0 Then
        sTown = g_clsSQLAssistSP.FormatWildcardedString(sTown, bFound)
    End If
    
    sCompanyName = txtSearch(PANEL_COMPANY_NAME).Text
    
    If sCompanyName = "*" Then
        txtSearch(PANEL_COMPANY_NAME).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "You must have at least one character to perform a wildcard search"
    End If
    
    If Len(sCompanyName) > 0 Then
        sCompanyName = g_clsSQLAssistSP.FormatWildcardedString(sCompanyName, bFound)
    End If
    
    sPanelID = txtSearch(PANEL_ID).Text
    
    If vIndex = COMBO_TYPE_LEGAL_REP Or vIndex = COMBO_TYPE_VALUER Then
        If Len(sTown) = 0 And Len(sPanelID) = 0 And Len(sCompanyName) = 0 Then
            txtSearch(PANEL_COMPANY_NAME).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "At least one of Town, Panel ID and Company Name must be entered"
        End If
    Else
        If Len(sTown) = 0 And Len(sCompanyName) = 0 Then
            txtSearch(PANEL_COMPANY_NAME).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "At least one of Town and Company Name must be entered"
        End If
    End If
    
    Select Case vIndex
        Case COMBO_TYPE_LEGAL_REP
            enumSearchType = ThirdPartyLegalRepType
        Case COMBO_TYPE_VALUER
            enumSearchType = ThirdPartyValuersType
        Case Else
            enumSearchType = ThirdPartyLocalType
    End Select
    
    Set clsPanel = New PanelTable
    clsPanel.SetSearchMode enumSearchType
    
    clsPanel.GetPanelFromSearch sCompanyName, sPanelID, sTown, CLng(vIndex)
    
    If TableAccess(clsPanel).RecordCount > 0 Then
        g_clsFormProcessing.PopulateFromRecordset m_lvResults, clsPanel
    Else
        Set clsPanel = Nothing
        g_clsErrorHandling.RaiseError errGeneralError, "No records found for your search criteria"
    End If
    
    Set clsPanel = Nothing
    
    Exit Sub
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboThirdPartyType_Click
' Description   :   Called when the Third Party combo is changed
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboThirdPartyType_Click()
    On Error GoTo Failed
    Dim vType As Variant
    Dim bEnable As Boolean
    
    bEnable = True
    ' Depending on the type selected, we need to set the listview headers
    g_clsFormProcessing.HandleComboExtra cboThirdPartyType, vType, GET_CONTROL_VALUE
    
    If Not IsEmpty(vType) Then
        If Len(vType) > 0 Then
            SetListViewHeaders CLng(vType)
            'BS BM0282
            'Don't clear listview to be consistent with other options
            'm_lvResults.ListItems.Clear
            If vType <> COMBO_TYPE_LEGAL_REP And vType <> COMBO_TYPE_VALUER Then
                bEnable = False
                txtSearch(PANEL_ID).Text = ""
            End If
            txtSearch(PANEL_TOWN).Text = ""
            txtSearch(PANEL_COMPANY_NAME).Text = ""
            txtSearch(PANEL_ID).Enabled = bEnable
            lblSearch(PANEL_ID).Enabled = bEnable
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsPanel = Nothing
End Sub
Public Sub SetTableClass(clsPanel As PanelTable)
    Set m_clsPanel = clsPanel
End Sub
Public Sub SetThirdPartyType(enumType As ThirdPartyType)
    m_enumThirdPartyType = enumType
End Sub
