VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditPrintingPack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pack Control"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   Icon            =   "frmEditPrintingPack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnother 
      Caption         =   "A&nother"
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame fraPack 
      Height          =   7635
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9495
      Begin MSGOCX.MSGComboBox cboPackDestination 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.CheckBox chkMultipleQuotes 
         Caption         =   "Pack contains multiple quotes"
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move Down"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         Picture         =   "frmEditPrintingPack.frx":0442
         TabIndex        =   9
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move Up"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8160
         Picture         =   "frmEditPrintingPack.frx":0884
         TabIndex        =   8
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   7080
         Width           =   1095
      End
      Begin MSGOCX.MSGMulti txtPackDescription 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1160
         Width           =   5895
         _ExtentX        =   10398
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
         Text            =   ""
         Mandatory       =   -1  'True
         MaxLength       =   256
      End
      Begin MSGOCX.MSGEditBox txtPackName 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         MaxLength       =   32
      End
      Begin MSGOCX.MSGEditBox txtPackId 
         Height          =   315
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         Mandatory       =   -1  'True
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
      Begin MSGOCX.MSGListView lvPackMembers 
         Height          =   4695
         Left            =   180
         TabIndex        =   5
         Top             =   2280
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   8281
         HideSelection   =   0   'False
         AllowColumnReorder=   0   'False
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Pack Destination"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label lblPackMembers 
         AutoSize        =   -1  'True
         Caption         =   "Pack Members"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2020
         Width           =   1065
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Pack Id"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   420
         Width           =   555
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Pack Description"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1220
         Width           =   1215
      End
      Begin VB.Label lblPrintingDocuments 
         AutoSize        =   -1  'True
         Caption         =   "Pack Name"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   820
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmEditPrintingPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
' Form          :   frmEditPrintingPack
' Description   :   Form which allows the user to edit and add Print Pack Control details
' Change history
' Prog      Date        Description
' GHun      14/10/2005  MAR202 Created
' RF        14/12/2005  MAR867 Complete pack handling changes
' HMA       11/01/2006  MAR966 Implement Move Up and Move Down buttons
'------------------------------------------------------------------------------------
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
' TW        19/03/2007  EP2_1276 - Inactive template should not be available for selection in a pack.
'                       Pass selected pack number to frmAddPackMember.
' TW        01/05/2007  EP2_2595 - Incomplete list of Template Names to create Pack
'------------------------------------------------------------------------------------
Option Explicit

Private m_ReturnCode As MSGReturnCode
Private m_clsPackControlTable As PackControlTable
Private m_clsPackMemberTable As PackMemberTable
Private m_bIsEdit As Boolean
Private m_colKeys As Collection

'ListView items have their keys prefixed with this value.
Private Const KEY_PREFIX As String = "Pack"

Private Const PACK_DESTINATION_COMBO_GROUP As String = "PackDestination"

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Private Sub Add()
On Error GoTo Failed
    Dim frmPackMember As frmAddPackMember
    Dim enumReturn As MSGReturnCode
    Dim collNewPackMemberData As Collection
    Dim lvItem As MSComctlLib.ListItem
    Dim strHostTemplateID As String

    'Create a pack member popup.
    Set frmPackMember = New frmAddPackMember
    
' TW 19/03/2007 EP2_1276
' TW 01/05/2007 EP2_2595
'    frmPackMember.SetKeys m_colKeys
    If m_bIsEdit Then
        frmPackMember.SetKeys m_colKeys
    Else
        frmPackMember.SetKeys Nothing
    End If
' TW 01/05/2007 EP2_2595 End
' TW 19/03/2007 EP2_1276 End
    
    'Show the form.
    frmPackMember.Show vbModal
    enumReturn = frmPackMember.GetReturnCode()
    If enumReturn = MSGSuccess Then
        ' update the list view
        Set collNewPackMemberData = frmPackMember.GetScreenData()
        Set lvItem = lvPackMembers.ListItems.Add
        strHostTemplateID = collNewPackMemberData("HOSTTEMPLATEID")
        lvItem.Text = strHostTemplateID
        lvItem.Key = GetKey(lvPackMembers.ListItems.Count + 1) ' this should be unique but HostTemplateID may not be
        lvItem.Tag = strHostTemplateID
        lvItem.SubItems(1) = GetHostTemplateName(strHostTemplateID)
        lvItem.SubItems(2) = GetHostTemplateDescription(strHostTemplateID)
    End If
    
    Unload frmPackMember
    
    Set frmPackMember = Nothing

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub Delete()
On Error GoTo Failed
    If Not lvPackMembers.SelectedItem Is Nothing Then
        lvPackMembers.RemoveLine lvPackMembers.SelectedItem.Index
        
        cmdDelete.Enabled = False
        cmdMoveDown.Enabled = False
        cmdMoveUp.Enabled = False
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError " Delete failed: " & Err.DESCRIPTION
End Sub

Private Sub Another()
On Error GoTo Failed
    g_clsErrorHandling.RaiseError 1, "Not implemented"
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub MoveDown()
On Error GoTo Failed

    Dim i As Integer
    Dim tempText As String
    Dim tempKey1 As String
    Dim tempKey2 As String
    Dim tempTag As String
    Dim tempSub1 As String
    Dim tempSub2 As String
    
    i = lvPackMembers.SelectedItem.Index
    
    tempKey1 = lvPackMembers.ListItems.Item(i).Key
    tempKey2 = lvPackMembers.ListItems.Item(i + 1).Key
    
    tempText = lvPackMembers.ListItems.Item(i + 1).Text
    tempTag = lvPackMembers.ListItems.Item(i + 1).Tag
    tempSub1 = lvPackMembers.ListItems.Item(i + 1).SubItems(1)
    tempSub2 = lvPackMembers.ListItems.Item(i + 1).SubItems(2)

    'Clear keys temporarily first
    lvPackMembers.ListItems.Item(i).Key = "Temp1"
    lvPackMembers.ListItems.Item(i + 1).Key = "Temp2"

    lvPackMembers.ListItems.Item(i + 1).Text = lvPackMembers.ListItems.Item(i).Text
    lvPackMembers.ListItems.Item(i + 1).Key = tempKey1
    lvPackMembers.ListItems.Item(i + 1).Tag = lvPackMembers.ListItems.Item(i).Tag
    lvPackMembers.ListItems.Item(i + 1).SubItems(1) = lvPackMembers.ListItems.Item(i).SubItems(1)
    lvPackMembers.ListItems.Item(i + 1).SubItems(2) = lvPackMembers.ListItems.Item(i).SubItems(2)
    
    lvPackMembers.ListItems.Item(i).Text = tempText
    lvPackMembers.ListItems.Item(i).Key = tempKey2
    lvPackMembers.ListItems.Item(i).Tag = tempTag
    lvPackMembers.ListItems.Item(i).SubItems(1) = tempSub1
    lvPackMembers.ListItems.Item(i).SubItems(2) = tempSub2
    
    Set lvPackMembers.SelectedItem = lvPackMembers.ListItems.Item(i + 1)
    lvPackMembers.SetFocus
    
    'If we have reached the bottom of the list, disable the Move Down button
    If ((i + 1) = lvPackMembers.ListItems.Count) Then
        cmdMoveDown.Enabled = False
    Else
        cmdMoveDown.Enabled = True
    End If

    If ((i + 1) = lvPackMembers.ListItems.Item(1).Index) Then
        cmdMoveUp.Enabled = False
    Else
        cmdMoveUp.Enabled = True
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub MoveUp()
On Error GoTo Failed
    
    Dim i As Integer
    Dim tempText As String
    Dim tempKey1 As String
    Dim tempKey2 As String
    Dim tempTag As String
    Dim tempSub1 As String
    Dim tempSub2 As String
    
    i = lvPackMembers.SelectedItem.Index
    
    tempKey1 = lvPackMembers.ListItems.Item(i).Key
    tempKey2 = lvPackMembers.ListItems.Item(i - 1).Key
    
    tempText = lvPackMembers.ListItems.Item(i - 1).Text
    tempTag = lvPackMembers.ListItems.Item(i - 1).Tag
    tempSub1 = lvPackMembers.ListItems.Item(i - 1).SubItems(1)
    tempSub2 = lvPackMembers.ListItems.Item(i - 1).SubItems(2)

    'Clear keys temporarily first
    lvPackMembers.ListItems.Item(i).Key = "Temp1"
    lvPackMembers.ListItems.Item(i - 1).Key = "Temp2"

    lvPackMembers.ListItems.Item(i - 1).Text = lvPackMembers.ListItems.Item(i).Text
    lvPackMembers.ListItems.Item(i - 1).Key = tempKey1
    lvPackMembers.ListItems.Item(i - 1).Tag = lvPackMembers.ListItems.Item(i).Tag
    lvPackMembers.ListItems.Item(i - 1).SubItems(1) = lvPackMembers.ListItems.Item(i).SubItems(1)
    lvPackMembers.ListItems.Item(i - 1).SubItems(2) = lvPackMembers.ListItems.Item(i).SubItems(2)
    
    lvPackMembers.ListItems.Item(i).Text = tempText
    lvPackMembers.ListItems.Item(i).Key = tempKey2
    lvPackMembers.ListItems.Item(i).Tag = tempTag
    lvPackMembers.ListItems.Item(i).SubItems(1) = tempSub1
    lvPackMembers.ListItems.Item(i).SubItems(2) = tempSub2
    
    Set lvPackMembers.SelectedItem = lvPackMembers.ListItems.Item(i - 1)
    lvPackMembers.SetFocus
    
    'If we have reached the top of the list, disable the Move Up button
    If ((i - 1) = lvPackMembers.ListItems.Item(1).Index) Then
        cmdMoveUp.Enabled = False
    Else
        cmdMoveUp.Enabled = True
    End If
            
    If ((i - 1) = lvPackMembers.ListItems.Count) Then
        cmdMoveDown.Enabled = False
    Else
        cmdMoveDown.Enabled = True
    End If
            
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub Form_Load()
On Error GoTo Failed

    Set m_clsPackControlTable = New PackControlTable
    Set m_clsPackMemberTable = New PackMemberTable
    
    PopulateScreenControls
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub PopulateScreenControls()
On Error GoTo Failed
        
    ' populate combo
    g_clsFormProcessing.PopulateCombo PACK_DESTINATION_COMBO_GROUP, cboPackDestination
    
    'Setup the column headers.
    SetListViewHeaders
               
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " PopulateScreenControls: " & Err.DESCRIPTION
End Sub

Private Sub SetEditState()
On Error GoTo Failed
    
    Dim clsTableAccess As TableAccess
    
    ' initialise pack control table
    Set clsTableAccess = m_clsPackControlTable
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    clsTableAccess.GetTableData
    
    If clsTableAccess.RecordCount = 0 Then
        g_clsErrorHandling.RaiseError , "Empty Record Set Returned"
    End If
    
    txtPackId.Enabled = False
    
    PopulateScreenFields
    
    ' initialise pack member table
    m_clsPackMemberTable.GetTableData m_clsPackControlTable.GetPackControlNumber()
    
    PopulateListView
       
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " SetEditState: " & Err.DESCRIPTION
End Sub

Private Sub PopulateScreenFields()
On Error GoTo Failed
    With m_clsPackControlTable
        txtPackId.Text = .GetPackControlNumber()
        txtPackName.Text = .GetPackControlName()
        txtPackDescription.Text = .GetPackControlDescription()
        g_clsFormProcessing.HandleComboExtra _
            cboPackDestination, .GetPackDestination(), SET_CONTROL_VALUE
        chkMultipleQuotes.Value = .GetUseMultipleQuotes()
    End With
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " PopulateScreenFields: " & Err.DESCRIPTION
End Sub

Private Sub SetAddState()
On Error GoTo Failed
    
    g_clsFormProcessing.CreateNewRecord m_clsPackControlTable
              
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " SetAddState: " & Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()
On Error GoTo Failed
    
    Dim vTmp As Variant
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim intCnt As Integer
    Dim blnPackMemberTableIsOpen As Boolean
    
    blnPackMemberTableIsOpen = False
               
    With m_clsPackControlTable
        .SetPackControlNumber txtPackId.Text
        .SetPackControlName txtPackName.Text
        .SetPackControlDescription txtPackDescription.Text
        g_clsFormProcessing.HandleComboExtra cboPackDestination, vTmp, GET_CONTROL_VALUE
        .SetPackDestination CStr(vTmp)
        .SetUseMultipleQuotes chkMultipleQuotes.Value
    End With
    
    Set clsTableAccess = m_clsPackControlTable
    TableAccess(m_clsPackControlTable).Update
    
    ' delete and recreate all pack members ...
    Set clsTableAccess = m_clsPackMemberTable
    ' delete
    If TableAccess(m_clsPackMemberTable).RecordCount > 0 Then
        blnPackMemberTableIsOpen = True
        TableAccess(m_clsPackMemberTable).DeleteAllRows
    End If
    ' create
    For intCnt = 1 To lvPackMembers.ListItems.Count
    
        blnPackMemberTableIsOpen = True
        
        TableAccess(m_clsPackMemberTable).AddRow

        With m_clsPackMemberTable
            .SetPackControlNumber m_clsPackControlTable.GetPackControlNumber()
            .SetPackMemberNumber intCnt
            
            ' currently only templates within packs are supported
            .SetPackMemberControlNumber Null
            .SetPackMemberType 2
            
            .SetHostTemplateId lvPackMembers.ListItems.Item(intCnt).Text
        End With
        
    Next
    
    If blnPackMemberTableIsOpen = True Then
        TableAccess(m_clsPackMemberTable).Update
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " SaveScreenData: " & Err.DESCRIPTION
End Sub

Private Function DoOKProcessing() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)

    If bRet = True Then
        bRet = ValidateScreenData()

        If bRet = True Then
            SaveScreenData
            SaveChangeRequest
        End If
        
    End If

    DoOKProcessing = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    DoOKProcessing = False
End Function

Private Sub cmdOK_Click()
On Error GoTo Failed
    Dim bRet As Boolean

    bRet = DoOKProcessing()

    If bRet = True Then
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub SaveChangeRequest()
On Error GoTo Failed
    Dim sDesc As String
    Dim colMatchValues As Collection

    Set colMatchValues = New Collection

    sDesc = txtPackId.Text

    colMatchValues.Add sDesc
    
    TableAccess(m_clsPackControlTable).SetKeyMatchValues colMatchValues

    g_clsHandleUpdates.SaveChangeRequest m_clsPackControlTable, sDesc

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " SaveChangeRequest: " & Err.DESCRIPTION
End Sub

Private Function ValidateScreenData() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    If m_bIsEdit = False Then
        bRet = Not DoesRecordExist()
    End If

    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " ValidateScreenData: " & Err.DESCRIPTION
End Function

Private Function DoesRecordExist() As Boolean
On Error GoTo Failed
    Dim bRet As Boolean
    Dim sPackID As String
    Dim colValues As Collection
    Set colValues = New Collection

    sPackID = Trim(txtPackId.Text)
    If Len(sPackID) > 0 Then
        colValues.Add sPackID

        bRet = TableAccess(m_clsPackControlTable).DoesRecordExist(colValues)

        If bRet = True Then
           g_clsErrorHandling.DisplayError "Pack ID must be unique"
           txtPackId.SetFocus
        End If
    End If
    DoesRecordExist = bRet
    Set colValues = Nothing
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " DoesRecordExist: " & Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetListViewHeaders
' Description   : Setup the column colHeaders on the ListViews.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetListViewHeaders()
    
    Dim colHeaders As Collection
    Dim lvHeaders As listViewAccess
    
    On Error GoTo Failed
    
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Template Id"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Template Name"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 54
    lvHeaders.sName = "Template Description"
    colHeaders.Add lvHeaders
    
    'Add the column header to the insurance type listview.
    lvPackMembers.AddHeadings colHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " SetListViewHeaders: " & Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the list view
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PopulateListView()
On Error GoTo Failed

    Dim lngCnt As Long
    Dim lngNumRecs As Long
    Dim lvItem As MSComctlLib.ListItem
    Dim clsTableAccess As TableAccess
    Dim strHostTemplateID As String

    While lvPackMembers.ListItems.Count > 0
        lvPackMembers.ListItems.Remove 1
    Wend

    Set clsTableAccess = m_clsPackMemberTable
    
    lngNumRecs = TableAccess(m_clsPackMemberTable).RecordCount
    If lngNumRecs > 0 Then
        TableAccess(m_clsPackMemberTable).MoveFirst
    
        For lngCnt = 1 To TableAccess(m_clsPackMemberTable).RecordCount
            With m_clsPackMemberTable
                Set lvItem = lvPackMembers.ListItems.Add
                strHostTemplateID = .GetHostTemplateID()
                lvItem.Text = strHostTemplateID
                lvItem.Key = GetKey(.GetPackMemberNumber) ' this should be unique but HostTemplateID may not be
                lvItem.Tag = strHostTemplateID
                lvItem.SubItems(1) = GetHostTemplateName(strHostTemplateID)
                lvItem.SubItems(2) = GetHostTemplateDescription(strHostTemplateID)
            End With
            TableAccess(m_clsPackMemberTable).MoveNext
        Next
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " PopulateListView: " & Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetKey
' Description   : Retuns a key to be used to index a listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetKey(ByVal vval As Variant) As String
On Error GoTo Failed
    GetKey = KEY_PREFIX & CStr(vval)
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " GetKey: " & Err.DESCRIPTION
End Function

Private Sub lvPackMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvPackMembers.SelectedItem Then
        cmdDelete.Enabled = True
        'If there is more than one item in the list enable move up and move down
        If lvPackMembers.ListItems.Count > 1 Then
            If Item.Index = lvPackMembers.ListItems.Item(1).Index Then
            'If Item.Index = 1 Then
                cmdMoveUp.Enabled = False
            Else
                cmdMoveUp.Enabled = True
            End If
            
            If Item.Index = lvPackMembers.ListItems.Count Then
                cmdMoveDown.Enabled = False
            Else
                cmdMoveDown.Enabled = True
            End If
        End If
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Function GetHostTemplateName(ByVal vstrHOSTTEMPLATEID As String) As Variant
On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim strHostTemplateName As String
    
    strHostTemplateName = ""
    
    If Len(vstrHOSTTEMPLATEID) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, _
            " GetHostTemplateName: HostTemplate ID is empty"
    End If
    
    sSearch = "SELECT HOSTTEMPLATENAME " _
        & " FROM HOSTTEMPLATE " _
        & " WHERE HOSTTEMPLATEID = '" & vstrHOSTTEMPLATEID & "'"
        
    Set rs = g_clsDataAccess.ExecuteCommand(sSearch)
    
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.fields(0).Value) Then
            strHostTemplateName = rs.fields(0).Value
        End If
    End If
    
    GetHostTemplateName = strHostTemplateName
        
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " GetHostTemplateName failed: " & Err.DESCRIPTION
    
End Function

Private Function GetHostTemplateDescription(ByVal vstrHOSTTEMPLATEID As String) As Variant
On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    Dim strHostTemplateDesc As String
    
    strHostTemplateDesc = ""
    
    If Len(vstrHOSTTEMPLATEID) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, _
            " GetHostTemplateDescription: HostTemplate ID is empty"
    End If
    
    sSearch = "SELECT HOSTTEMPLATEDESCRIPTION " _
        & " FROM HOSTTEMPLATE " _
        & " WHERE HOSTTEMPLATEID = '" & vstrHOSTTEMPLATEID & "'"
        
    Set rs = g_clsDataAccess.ExecuteCommand(sSearch)
    
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.fields(0).Value) Then
            strHostTemplateDesc = rs.fields(0).Value
        End If
    End If
    
    GetHostTemplateDescription = strHostTemplateDesc
        
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, " GetHostTemplateDescription failed: " & Err.DESCRIPTION
    
End Function

Private Sub cmdAdd_Click()
    Add
End Sub
Private Sub cmdAnother_Click()
    Another
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    Delete
End Sub
Private Sub cmdMoveDown_Click()
    MoveDown
End Sub
Private Sub cmdMoveUp_Click()
    MoveUp
End Sub


