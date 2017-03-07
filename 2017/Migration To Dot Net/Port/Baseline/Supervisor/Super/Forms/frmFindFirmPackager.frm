VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindFirmPackager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Packager Firm"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmFindFirmPackager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtFirmName 
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   5235
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Find"
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin MSGOCX.MSGEditBox txtFirmAddress 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
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
      Begin MSGOCX.MSGEditBox txtUnitId 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblFirmName 
         AutoSize        =   -1  'True
         Caption         =   "Firm Name"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit Id"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Search Results"
      Height          =   3675
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   8535
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   915
      End
      Begin MSGOCX.MSGListView lvResults 
         Height          =   2715
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4789
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "frmFindFirmPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindFirmPackager
' Description   : Finds all introducers matching the search criteria
'
' Change history
' Prog      Date        Description
' PB        27/10/06    EP2_13 - Copied from frmFindFirmPackager to frmFindFirmPackager
'                       and modified to use for firm packager.
' TW        18/11/2006  EP2_132 ECR20/21 Header corrections
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' TW        13/03/2007  EP2_1934 - DBM341 - Satellite Packagers - Allow firm name of up to 50 characters.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table object.
Private m_objFirmPackagerTable As FirmPackagerTable

Public Function GetTableClass() As IntroducerTable
    Set GetTableClass = m_objFirmPackagerTable
End Function
Public Sub SetTableClass(TableClass As FirmPackagerTable)
    Set m_objFirmPackagerTable = TableClass
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdClear_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()

    Dim nThisControl As Integer
    Dim ctrThisControl As Control
    
    On Error GoTo Failed

    'Clear all items from the listview.
    ClearLV
    
    'Clear the contents of the criteria controls.
    'For nThisControl = 0 To txtSearch.Count - 1
    For Each ctrThisControl In Me.Controls
        If TypeOf ctrThisControl Is MSGEditBox Then
            ctrThisControl.Text = ""
        End If
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEdit_Click
' Description   : Opens the introducer form in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
    
    On Error GoTo Failed

    EditFirmPackager

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
    
    BeginWaitCursor
    FindFirmPackager
    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


Private Sub Form_Initialize()
    '
    'Create the underlying table object.
    Set m_objFirmPackagerTable = New FirmPackagerTable
    '
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Create any required table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    'Populate static lists and setup column headers.
    PopulateScreenControls

    'Select the first combo item.
    'cboIntroducerType.ListIndex = 0
    
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
        
    Dim colHeaders As Collection
    Dim lvHeaders As listViewAccess
    
    On Error GoTo Failed
    
    Set colHeaders = New Collection
    
' TW 16/01/2007 EP2_859
    'FSA Ref
'    lvHeaders.nWidth = 10
'    lvHeaders.sName = "FSA Ref"
'    colHeaders.Add lvHeaders
    
    'Unit Id
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Unit Id"
    colHeaders.Add lvHeaders
    
    'Company Name
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Firm Name"
    colHeaders.Add lvHeaders
    
    'Address
    lvHeaders.nWidth = 49
    lvHeaders.sName = "Address"
    colHeaders.Add lvHeaders
' TW 16/01/2007 EP2_859 End
    
    'Add the headers (collection) into the listview.
    lvResults.AddHeadings colHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindFirmPackager
' Description   : Returns all matching introducer records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindFirmPackager()
    
    Dim objFirmPackagerSearch As FirmPackagerSearch
    
    On Error GoTo Failed
    
    'Create a class used to search.
    Set objFirmPackagerSearch = New FirmPackagerSearch
    
' TW 16/01/2007 EP2_859
'    'Set the FSA ref criteria.
'    If Me.txtFSAReference.Enabled = True Then
'        objFirmPackagerSearch.SetFSARef Me.txtFSAReference.Text
'    End If
    
    'Set the Unit criteria.
    If Me.txtUnitId.Enabled = True Then
        objFirmPackagerSearch.SetUnit Me.txtUnitId.Text
    End If
' TW 16/01/2007 EP2_859 End
    
    'Set the firm name criteria.
    If Me.txtFirmName.Enabled = True Then
        objFirmPackagerSearch.SetFirmName txtFirmName.Text
    End If
    
    'Set address
    If Me.txtFirmAddress.Enabled Then
        objFirmPackagerSearch.SetAddress txtFirmAddress.Text
    End If
    
    'Perform a search and populate the table object.
    objFirmPackagerSearch.Find m_objFirmPackagerTable
    
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
    
    On Error GoTo Failed

    'Clear existing rows.
    ClearLV
    
    'Add each record into the listview.
    If TableAccess(m_objFirmPackagerTable).RecordCount > 0 Then
        For nThisInt = 1 To TableAccess(m_objFirmPackagerTable).RecordCount
            Set lvItem = lvResults.ListItems.Add
            
            'Unique key
            lvItem.Key = "fp" & m_objFirmPackagerTable.GetPrincipalFirmID

            'Because the key can only be a string value, the id must be stored in the tag as a byte()
            lvItem.Tag = m_objFirmPackagerTable.GetPrincipalFirmID
            
' TW 16/01/2007 EP2_859
'            lvItem.Text = m_objFirmPackagerTable.GetFSARef
            lvItem.Text = m_objFirmPackagerTable.GetUnitId
            
' TW 16/01/2007 EP2_859 End
            
            lvItem.ListSubItems.Add , , m_objFirmPackagerTable.GetName
            
            lvItem.ListSubItems.Add , , m_objFirmPackagerTable.GetAddress
            
            TableAccess(m_objFirmPackagerTable).MoveNext
        Next
' TW 16/01/2007 EP2_859
    Else
        MsgBox "No records found matching your Search Criteria", vbExclamation
' TW 16/01/2007 EP2_859 End
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
    Set m_objFirmPackagerTable = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_DblClick
' Description   : Open the selected introducer.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_DblClick()
    
    On Error GoTo Failed

    EditFirmPackager
    
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
' Function      : EditFirmPackager
' Description   : Loads frmEditFirmPackager passing through the GUID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditFirmPackager()
        
    Dim varKey As Variant
    Dim colKeys As Collection
    Dim frmEdit As frmEditPackager
    
    On Error GoTo Failed
    
    If Not lvResults.SelectedItem Is Nothing Then
        Set frmEdit = frmEditPackager
        Set colKeys = New Collection
        
        varKey = lvResults.SelectedItem.Tag
        
        frmEdit.SetIsEdit True
        
        colKeys.Add varKey
        
        frmEdit.SetKeys colKeys
        
        frmEdit.Show vbModal
        
        Unload frmEdit
        
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


