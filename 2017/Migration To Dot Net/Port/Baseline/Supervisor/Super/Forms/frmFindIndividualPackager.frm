VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindIndividualPackager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Individual Packager"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmFindIndividualPackager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Find"
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin MSGOCX.MSGEditBox txtFirmName 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
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
      Begin MSGOCX.MSGEditBox txtIntroducerId 
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Introducer Id"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   900
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
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   570
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Search Results"
      Height          =   3795
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
         Top             =   3360
         Width           =   915
      End
      Begin MSGOCX.MSGListView lvResults 
         Height          =   2835
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5001
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frmFindIndividualPackager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindIndividualPackager
' Description   : Finds all introducers matching the search criteria
'
' Change history
' Prog      Date        Description
' PB        27/10/06    EP2_13 - Copied from frmFindIntermediary to frmFindIndividualPackager
'                       and modified to use for individual packager.
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table object.
Private m_objIndividualPackagerTable As IndividualPackagerTable

'Control indexes.
Private Const INTRODUCER_FORENAME     As Long = 0
Private Const INTRODUCER_SURNAME      As Long = 1
Private Const INTRODUCER_COMPANY_NAME As Long = 2
Private Const INTRODUCER_TOWN         As Long = 3

Public Function GetTableClass() As IntroducerTable
    Set GetTableClass = m_objIndividualPackagerTable
End Function
Public Sub SetTableClass(TableClass As IndividualPackagerTable)
    Set m_objIndividualPackagerTable = TableClass
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

    EditIndividualPackager

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
    FindIndividualPackager
    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


Private Sub Form_Initialize()
    '
    'Create the underlying table object.
    Set m_objIndividualPackagerTable = New IndividualPackagerTable
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
    
    'Introducer Type combo.
    'g_clsFormProcessing.PopulateCombo "IntroducerType", cboIntroducerType, False

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
    
    'Id
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Id"
    colHeaders.Add lvHeaders
    
    'Name
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Name"
    colHeaders.Add lvHeaders
    
    
    'Address
    lvHeaders.nWidth = 59
    lvHeaders.sName = "Address"
    colHeaders.Add lvHeaders
        
    'Add the headers (collection) into the listview.
    lvResults.AddHeadings colHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableCompanyControls
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableCompanyControls(Optional ByVal bEnable As Boolean = True)
    
    On Error GoTo Failed
    
    'Me.lblFullName.Enabled = Not bEnable
    'Me.txtFullName.Enabled = Not bEnable
    
    'Me.lblCompanyName.Enabled = bEnable
    'Me.txtCompanyName.Enabled = bEnable

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindIndividualPackager
' Description   : Returns all matching introducer records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindIndividualPackager()
    
    Dim objIndividualPackagerSearch As IndividualPackagerSearch
    
    On Error GoTo Failed
    
    'Create a class used to search.
    Set objIndividualPackagerSearch = New IndividualPackagerSearch

    'Set the Introducer Id criteria.
    If Me.txtIntroducerId.Enabled = True Then
        objIndividualPackagerSearch.SetIntroducerID Me.txtIntroducerId.Text
    End If
    
    'Set the firm name criteria.
    If Me.txtFirmName.Enabled = True Then
        objIndividualPackagerSearch.SetFirmName txtFirmName.Text
    End If
    
    'Set address
    If Me.txtFirmAddress.Enabled Then
        objIndividualPackagerSearch.SetAddress txtFirmAddress.Text
    End If
    
    'Perform a search and populate the table object.
    objIndividualPackagerSearch.Find m_objIndividualPackagerTable
    
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
    If TableAccess(m_objIndividualPackagerTable).RecordCount > 0 Then
        For nThisInt = 1 To TableAccess(m_objIndividualPackagerTable).RecordCount
            Set lvItem = lvResults.ListItems.Add
            
            'Unique key
            lvItem.Key = "ip" & m_objIndividualPackagerTable.GetIndividualPackagerId
            
            'Because the key can only be a string value, the id must be stored in the tag as a byte()
            lvItem.Tag = m_objIndividualPackagerTable.GetIndividualPackagerId
            
'            lvItem.Text = m_objIndividualPackagerTable.GetForename
'
'            lvItem.ListSubItems.Add , , m_objIndividualPackagerTable.GetSurname
            
            lvItem.Text = m_objIndividualPackagerTable.GetIndividualPackagerId
            lvItem.ListSubItems.Add , , m_objIndividualPackagerTable.GetUserName
            lvItem.ListSubItems.Add , , m_objIndividualPackagerTable.GetAddress
            
            
            TableAccess(m_objIndividualPackagerTable).MoveNext
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
    Set m_objIndividualPackagerTable = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_DblClick
' Description   : Open the selected introducer.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_DblClick()
    
    On Error GoTo Failed

    EditIndividualPackager
    
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
' Function      : EditIndividualPackager
' Description   : Loads frmEditIndividualPackager passing through the GUID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditIndividualPackager()
        
    Dim varKey As Variant
    Dim colKeys As Collection
    Dim frmEdit As frmEditIndividualPackager
    
    On Error GoTo Failed
    
    If Not lvResults.SelectedItem Is Nothing Then
        Set frmEdit = frmEditIndividualPackager
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


