VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindClubs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Clubs"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "frmFindClubs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Find"
         Height          =   375
         Left            =   7200
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSGOCX.MSGEditBox txtFirmName 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblFirmName 
         AutoSize        =   -1  'True
         Caption         =   "Club Name"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Search Results"
      Height          =   3795
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   8535
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   3360
         Width           =   915
      End
      Begin MSGOCX.MSGListView lvResults 
         Height          =   2835
         Left            =   240
         TabIndex        =   4
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
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
End
Attribute VB_Name = "frmFindClubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindClub
' Description   : Finds all introducers matching the search criteria
'
' Change history
' Prog      Date        Description
' PB        03/11/06    EP2_13 - Copied from frmFindFirmBroker to frmFindClub
'                       and modified to use for Club.
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table object.
Private m_objClubTable As ClubTable

Public Function GetTableClass() As IntroducerTable
    Set GetTableClass = m_objClubTable
End Function
Public Sub SetTableClass(TableClass As ClubTable)
    Set m_objClubTable = TableClass
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

    EditClub

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
    FindClub
    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


Private Sub Form_Initialize()
    '
    'Create the underlying table object.
    Set m_objClubTable = New ClubTable
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
    
    'Name
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Name"
    colHeaders.Add lvHeaders
    
' TW 16/01/2007 EP2_859
'    'Description
'    lvHeaders.nWidth = 40
'    lvHeaders.sName = "Description"
'    colHeaders.Add lvHeaders
'
'    'Building name
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "House Name"
'    colHeaders.Add lvHeaders
'
'    'Building number
'    lvHeaders.nWidth = 10
'    lvHeaders.sName = "House Number"
'    colHeaders.Add lvHeaders
'
'    'Street
'    lvHeaders.nWidth = 15
'    lvHeaders.sName = "Street"
'    colHeaders.Add lvHeaders
'
'    'Postcode
'    lvHeaders.nWidth = 10
'    lvHeaders.sName = "Postcode"
'    colHeaders.Add lvHeaders
    
    'Address
    lvHeaders.nWidth = 74
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
' Function      : FindClub
' Description   : Returns all matching introducer records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindClub()
    
    Dim objClubSearch As ClubSearch
    
    On Error GoTo Failed
    
    'Create a class used to search.
    Set objClubSearch = New ClubSearch
    
    'Set the Name criteria.
    If Me.txtFirmName.Enabled = True Then
        objClubSearch.SetName Me.txtFirmName.Text
    End If
    
    'Set address
    If Me.txtFirmAddress.Enabled Then
        objClubSearch.SetAddress Me.txtFirmAddress.Text
    End If
    
    'Perform a search and populate the table object.
    objClubSearch.Find m_objClubTable
    
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
    If TableAccess(m_objClubTable).RecordCount > 0 Then
        For nThisInt = 1 To TableAccess(m_objClubTable).RecordCount
            Set lvItem = lvResults.ListItems.Add
            
            'Unique key
            lvItem.Key = "ASSOC" & m_objClubTable.GetClubId ' GetIntroducerGUID
            
            'Because the key can only be a string value, the id must be stored in the tag as a byte()
            lvItem.Tag = m_objClubTable.GetClubId
            
            lvItem.Text = m_objClubTable.GetName
            
' TW 16/01/2007 EP2_859
'            lvItem.ListSubItems.Add , , m_objClubTable.GetDescription
'            lvItem.ListSubItems.Add , , m_objClubTable.GetBuildingHouseName
'            lvItem.ListSubItems.Add , , m_objClubTable.GetBuildingHouseNumber
'            lvItem.ListSubItems.Add , , m_objClubTable.GetStreet
'            lvItem.ListSubItems.Add , , m_objClubTable.GetPostCode
            lvItem.ListSubItems.Add , , m_objClubTable.GetAddress
' TW 16/01/2007 EP2_859 End
            
            TableAccess(m_objClubTable).MoveNext
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
    Set m_objClubTable = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_DblClick
' Description   : Open the selected introducer.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_DblClick()
    
    On Error GoTo Failed

    EditClub
    
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
' Function      : EditClub
' Description   : Loads frmEditClub passing through the GUID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditClub()
        
    Dim varKey As Variant
    Dim colKeys As Collection
    Dim frmEdit As frmEditAssociation
    
    On Error GoTo Failed
    
    If Not lvResults.SelectedItem Is Nothing Then
        Set frmEdit = frmEditAssociation
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



