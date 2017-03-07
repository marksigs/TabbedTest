VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindIntermediary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intermediary Search"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "frmFindIntermediary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGComboBox cboIntermediaryType 
      Height          =   315
      Left            =   1860
      TabIndex        =   15
      Top             =   240
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
      Mandatory       =   -1  'True
      Text            =   ""
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6420
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Find"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   720
      Width           =   5235
      _ExtentX        =   9234
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
   Begin VB.Frame fraResults 
      Caption         =   "Search Results"
      Height          =   3795
      Left            =   180
      TabIndex        =   9
      Top             =   2460
      Width           =   8535
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   3360
         Width           =   915
      End
      Begin MSGOCX.MSGListView lvResults 
         Height          =   2715
         Left            =   360
         TabIndex        =   6
         Top             =   540
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4789
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   1
      Left            =   1860
      TabIndex        =   1
      Top             =   1140
      Width           =   5235
      _ExtentX        =   9234
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
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   2
      Left            =   1860
      TabIndex        =   2
      Top             =   1560
      Width           =   5235
      _ExtentX        =   9234
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
      MaxLength       =   45
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   3
      Left            =   1860
      TabIndex        =   3
      Top             =   1980
      Width           =   5235
      _ExtentX        =   9234
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
   Begin VB.Label lblIntermediary 
      Caption         =   "Company Name"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Town"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   1980
      Width           =   1515
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Surname"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Forename"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label lblIntermediary 
      Caption         =   "Intermediary Type"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmFindIntermediary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindIntermediary
' Description   : Finds all intermediaries matching the search criteria
'
' Change history
' Prog      Date        Description
' AA        26/06/2001  Created
' DJP       27/06/01    SQL Server port
' STB       21/12/01    SYS2550 - Integrate Intermediaries.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table object.
Private m_clsIntermediaryTable As IntermediaryTable

'Control indexes.
Private Const INTERMEDIARY_FORENAME     As Long = 0
Private Const INTERMEDIARY_SURNAME      As Long = 1
Private Const INTERMEDIARY_COMPANY_NAME As Long = 2
Private Const INTERMEDIARY_TOWN         As Long = 3


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdClear_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()

    Dim nThisControl As Integer
    
    On Error GoTo Failed

    'Clear all items from the listview.
    ClearLV
    
    'Clear the contents of the criteria controls.
    For nThisControl = 0 To txtSearch.Count - 1
        txtSearch(nThisControl).Text = ""
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEdit_Click
' Description   : Opens the intermediary form in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
    
    On Error GoTo Failed

    EditIntermediary

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
        FindIntermediary
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
    
    'Create the underlying table object.
    Set m_clsIntermediaryTable = New IntermediaryTable
    
    'Populate static lists and setup column headers.
    PopulateScreenControls

    'Select the first combo item.
    cboIntermediaryType.ListIndex = 0

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
    
    'Intermediary Type combo.
    g_clsFormProcessing.PopulateCombo "IntermediaryType", cboIntermediaryType, False

    'Setup the column headers in the listview.
    SetLVHeaders

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLVHeaders
' Description   : Construct the column headers.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLVHeaders()
        
    Dim colheaders As Collection
    Dim lvHeaders As listViewAccess
    
    On Error GoTo Failed
    
    Set colheaders = New Collection
    
    'Name.
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Name"
    colheaders.Add lvHeaders
    
    'Panel ID.
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Panel ID"
    colheaders.Add lvHeaders
    
    'Town.
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Town"
    colheaders.Add lvHeaders
    
    'Add the headers (collection) into the listview.
    lvResults.AddHeadings colheaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboIntermediaryType_Click
' Description   :   Checks if the selected item is PaymentProcessing, using the
'                   validation type Returned from the combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboIntermediaryType_Click()
    
    Dim uIntermediaryType As IntermediaryTypeEnum
    
    On Error GoTo Failed
    
    'Get the Int Type Enum of the selected item.
    g_clsFormProcessing.HandleComboExtra cboIntermediaryType, uIntermediaryType, GET_CONTROL_VALUE
    
    'Only enable the company controls if not individual.
    EnableCompanyControls (uIntermediaryType <> IndividualType)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableCompanyControls
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableCompanyControls(Optional ByVal bEnable As Boolean = True)
    
    On Error GoTo Failed
    
    If bEnable Then
        txtSearch(INTERMEDIARY_SURNAME).Enabled = False
        lblIntermediary(INTERMEDIARY_SURNAME).Enabled = False
        
        txtSearch(INTERMEDIARY_FORENAME).Enabled = False
        lblIntermediary(INTERMEDIARY_FORENAME).Enabled = False
        
        txtSearch(INTERMEDIARY_COMPANY_NAME).Enabled = True
        lblIntermediary(INTERMEDIARY_COMPANY_NAME).Enabled = True
    Else
        txtSearch(INTERMEDIARY_SURNAME).Enabled = True
        lblIntermediary(INTERMEDIARY_SURNAME).Enabled = True
        
        txtSearch(INTERMEDIARY_FORENAME).Enabled = True
        lblIntermediary(INTERMEDIARY_FORENAME).Enabled = True
        
        txtSearch(INTERMEDIARY_COMPANY_NAME).Enabled = False
        lblIntermediary(INTERMEDIARY_COMPANY_NAME).Enabled = False
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindIntermediary
' Description   : Returns all matching intermediary records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindIntermediary()
    
    Dim uIntermediaryType As IntermediaryTypeEnum
    Dim clsFindIntermediary As IntermediarySearch
    
    On Error GoTo Failed
    
    'Create a class used to search.
    Set clsFindIntermediary = New IntermediarySearch

    'Set the Company Name criteria.
    If txtSearch(INTERMEDIARY_COMPANY_NAME).Enabled = True Then
        clsFindIntermediary.SetCompanyName txtSearch(INTERMEDIARY_COMPANY_NAME).Text
    End If
    
    'Set the forename criteria.
    If txtSearch(INTERMEDIARY_FORENAME).Enabled = True Then
        clsFindIntermediary.SetForename txtSearch(INTERMEDIARY_FORENAME).Text
    End If
    
    'Set the surname criteria.
    If txtSearch(INTERMEDIARY_SURNAME).Enabled = True Then
        clsFindIntermediary.SetSurname txtSearch(INTERMEDIARY_SURNAME).Text
    End If
    
    'Set the town criteria.
    clsFindIntermediary.SetTown txtSearch(INTERMEDIARY_TOWN).Text
    
    'Get the selected combo item.
    g_clsFormProcessing.HandleComboExtra cboIntermediaryType, uIntermediaryType, GET_CONTROL_VALUE
    
    'Set the intermediary-type criteria.
    clsFindIntermediary.SetType uIntermediaryType
    
    'Perform a search and populate the table object.
    clsFindIntermediary.FindIntermediaries m_clsIntermediaryTable
    
    'Populate the listview from the data.
    PopulateLV
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateLV
' Description   : Populate the LVResults LV with the search results
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateLV()
        
    Dim nThisInt As Long
    Dim lvItem As MSComctlLib.ListItem
    Dim uIntermediaryType As IntermediaryTypeEnum
    
    On Error GoTo Failed

    'Clear existing rows.
    ClearLV
    
    'Add each record into the listview.
    If TableAccess(m_clsIntermediaryTable).RecordCount > 0 Then
        For nThisInt = 1 To TableAccess(m_clsIntermediaryTable).RecordCount
            Set lvItem = lvResults.ListItems.Add
            
            'Unique key
            lvItem.Key = m_clsIntermediaryTable.GetIntermediaryGUID
            
            'Because the key can only be a string value, the guid must be stored in the tag as a byte()
            lvItem.Tag = m_clsIntermediaryTable.GetIntermediaryGUID
            
            'Get the selected intermediary type.
            g_clsFormProcessing.HandleComboExtra cboIntermediaryType, uIntermediaryType, GET_CONTROL_VALUE
            
            If uIntermediaryType = IndividualType Then
                lvItem.Text = m_clsIntermediaryTable.GetIntermediaryName(INDIVIDUAL)
            Else
                lvItem.Text = m_clsIntermediaryTable.GetIntermediaryName(INTERMEDIARY_COMPANY)
            End If
            
            lvItem.ListSubItems.Add , , m_clsIntermediaryTable.GetPanelID
            lvItem.ListSubItems.Add , , m_clsIntermediaryTable.GetTown
            
            TableAccess(m_clsIntermediaryTable).MoveNext
        Next
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ClearLV
' Description   : Clear the results listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearLV()
    
    Dim nThisInt As Long

    On Error GoTo Failed

    For nThisInt = lvResults.ListItems.Count To 1 Step -1
        lvResults.ListItems.Remove nThisInt
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    Set m_clsIntermediaryTable = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_DblClick
' Description   : Open the selected intermediary.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_DblClick()
    
    On Error GoTo Failed

    EditIntermediary
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_ItemClick
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo Failed
    
    cmdEdit.Enabled = True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EditIntermediary
' Description   : Loads frmEditIntermediary passing through the GUID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditIntermediary()
        
    Dim vGuid As Variant
    Dim colKeys As Collection
    Dim frmEdit As frmEditIntermediaries
    
    On Error GoTo Failed
    
    If Not lvResults.SelectedItem Is Nothing Then
        Set frmEdit = frmEditIntermediaries
        Set colKeys = New Collection
        
        vGuid = lvResults.SelectedItem.Tag
        
        frmEdit.SetIsEdit
        
        colKeys.Add vGuid
        
        frmEdit.SetKeys colKeys
        
        frmEdit.Show vbModal
        
        Unload frmEdit
        
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
