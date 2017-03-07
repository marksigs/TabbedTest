VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmFindDABroker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find DA Broker"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmFindDABroker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Search Criteria"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Find"
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear"
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin MSGOCX.MSGEditBox txtFirmName 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
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
         Left            =   1800
         TabIndex        =   3
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
      Begin MSGOCX.MSGEditBox txtFSAReference 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
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
      Begin MSGOCX.MSGEditBox txtIntroducerId 
         Height          =   315
         Left            =   1800
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
         TabIndex        =   14
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblFSAReference 
         AutoSize        =   -1  'True
         Caption         =   "FSA Reference"
         Height          =   195
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblIntroducerName 
         AutoSize        =   -1  'True
         Caption         =   "Introducer Name"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   570
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "Search Results"
      Height          =   3795
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   8535
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   915
      End
      Begin MSGOCX.MSGListView lvResults 
         Height          =   2715
         Left            =   240
         TabIndex        =   6
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
      TabIndex        =   8
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "frmFindDABroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindDABroker
' Description   : Finds all intermediaries matching the search criteria
'
' Change history
' Prog      Date        Description
' PB        08/11/2006  Created
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
' TW        01/02/2007  EP2_1036 - Principal Firms/Network display and search (Follow-on)
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table object.
Private m_objDABrokerTable As DABrokerTable

'Control indexes.
Private Const DABROKER_FORENAME     As Long = 0
Private Const DABROKER_SURNAME      As Long = 1
Private Const DABROKER_COMPANY_NAME As Long = 2
Private Const DABROKER_TOWN         As Long = 3

Public Sub SetTableClass(objDABrokerTable As TableAccess)
    Set m_objDABrokerTable = objDABrokerTable
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdClear_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClear_Click()

    Dim ctrControl As Control
    
    On Error GoTo Failed

    'Clear all items from the listview.
    ClearLV
    
    'Clear the contents of the criteria controls.
    For Each ctrControl In Me.Controls
        If TypeOf ctrControl Is MSGEditBox Then
            ctrControl.Text = ""
        End If
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEdit_Click
' Description   : Opens the DABroker form in edit mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
    
    On Error GoTo Failed

    EditDABroker

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
        FindDABroker
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
    Set m_objDABrokerTable = New DABrokerTable
    
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
    
' TW 16/01/2007 EP2_859
    'FSA Ref
    lvHeaders.nWidth = 10
    lvHeaders.sName = "FSA Ref"
    colHeaders.Add lvHeaders
' TW 16/01/2007 EP2_859 End
    
    'Name
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Name"
    colHeaders.Add lvHeaders
    
    'Address
    lvHeaders.nWidth = 49
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
    
   
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindDABroker
' Description   : Returns all matching DABroker records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FindDABroker()
    
    'Dim uDABrokerType As DABrokerTypeEnum
    Dim clsFindDABroker As DABrokerSearch
    
    On Error GoTo Failed
    
    'Create a class used to search.
    Set clsFindDABroker = New DABrokerSearch

' TW 16/01/2007 EP2_859
    'Set the Introducer Id criteria.
    If Me.txtIntroducerId.Enabled = True Then
        clsFindDABroker.SetIntroducerID Me.txtIntroducerId.Text
    End If
' TW 16/01/2007 EP2_859 End
    
    'Set the FSA ref criteria.
    If Me.txtFSAReference.Enabled = True Then
        clsFindDABroker.SetFSARef Me.txtFSAReference.Text
    End If
    
    'Set the firm name criteria.
    If Me.txtFirmName.Enabled = True Then
        clsFindDABroker.SetFirmName txtFirmName.Text
    End If
    
    'Set address
    If Me.txtFirmAddress.Enabled Then
        clsFindDABroker.SetAddress txtFirmAddress.Text
    End If
        
    'Perform a search and populate the table object.
    clsFindDABroker.Find m_objDABrokerTable
    
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
    'Dim uDABrokerType As DABrokerTypeEnum
    
    On Error GoTo Failed

    'Clear existing rows.
    ClearLV
    
    'Add each record into the listview.
    If TableAccess(m_objDABrokerTable).RecordCount > 0 Then
        For nThisInt = 1 To TableAccess(m_objDABrokerTable).RecordCount
            Set lvItem = lvResults.ListItems.Add
            
            'Unique key
            lvItem.Key = "intro" & m_objDABrokerTable.GetIntroducerID & m_objDABrokerTable.GetFSARef
            
            'Because the key can only be a string value, the guid must be stored in the tag as a byte()
            lvItem.Tag = m_objDABrokerTable.GetIntroducerID
            
            lvItem.Text = m_objDABrokerTable.GetIntroducerID
' TW 16/01/2007 EP2_859
' TW 01/02/2007 EP2_1036
'            lvItem.ListSubItems.Add , , m_objARBrokerTable.GetFSARef
            lvItem.ListSubItems.Add , , "" 'Individual FSA Ref needed when available
' TW 01/02/2007 EP2_1036 End
' TW 16/01/2007 EP2_859 End
            lvItem.ListSubItems.Add , , m_objDABrokerTable.GetUserName
            lvItem.ListSubItems.Add , , m_objDABrokerTable.GetAddress
            
            TableAccess(m_objDABrokerTable).MoveNext
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
    Set m_objDABrokerTable = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvResults_DblClick
' Description   : Open the selected DABroker.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvResults_DblClick()
    
    On Error GoTo Failed

    EditDABroker
    
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
' Function      : EditDABroker
' Description   : Loads frmEditDABroker passing through the GUID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditDABroker()
        
    Dim vGUID As Variant
    Dim colKeys As Collection
    Dim frmEdit As frmEditIndividualDABroker
    
    On Error GoTo Failed
    
    If Not lvResults.SelectedItem Is Nothing Then
        Set frmEdit = frmEditIndividualDABroker
        Set colKeys = New Collection
        
        vGUID = lvResults.SelectedItem.Tag
        
        frmEdit.SetIsEdit True
        
        colKeys.Add vGUID
        
        frmEdit.SetKeys colKeys
        
        frmEdit.Show vbModal
        
        Unload frmEdit
        
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


