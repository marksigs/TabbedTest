VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIntTargetDetails 
   Caption         =   "Add/Edit Intermediary Target Details"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   Icon            =   "frmEditIntTargetDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTarget 
      Caption         =   "Target Details"
      Height          =   4335
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   7635
      Begin VB.CommandButton cmdDeleteTarget 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditTarget 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1740
         TabIndex        =   2
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddTarget 
         Caption         =   "Add"
         Height          =   375
         Left            =   420
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin MSGOCX.MSGListView lvTargetDetails 
         Height          =   3135
         Left            =   360
         TabIndex        =   0
         Top             =   420
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   5530
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditIntTargetDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditIntTargetDetails
' Description   : Allows the add/edit of target data for an interemdiary.
'
' Change history
' Prog      Date        Description
' AA        26/06/01    Created
' STB       27/11/01    Cancelling frmEditTargetDetails and then this form no-
'                       longer errors.
' STB       07/12/01    SYS2550 SQL Server support.
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'These table object references are set by the interemediary form
'prior to opening this form.
Private m_clsTargetTable As IntermediaryTargetTable

'NOTE: The Cancel button has been removed from this screen. Whilst there is
'still the facility to cancel from both the Target Popup screen AND the main
'Intermediaries screen, there is no simple mechanism to implement it at this
'level of the screen hierarchy (within the time constraints allowed).


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAddTarget_Click
' Description   : Opens the Target Details form ready for adding a new target.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddTarget_Click()
    
    On Error GoTo Failed
    
    'Show the target details form in Add mode.
    ShowTargetPopup False
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdDeleteTarget_Click
' Description   : Prepares the selected Target for deletion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDeleteTarget_Click()
    
    Dim lSequence As Long
    Dim clsDetails As PopulateDetails
    
    On Error GoTo Failed
    
    'Only attempt to remove the item from the listview if there is one selected.
    If Not lvTargetDetails.SelectedItem Is Nothing Then
        'Get the populate details object corresponding to the current selection.
        Set clsDetails = lvTargetDetails.GetExtra(lvTargetDetails.SelectedItem.Index)
        
        'Obtain its sequence number (part of its key).
        lSequence = clsDetails.GetExtra()
                
        'Place a filter on the recordset so only it is available.
        m_clsTargetTable.FindRecord CStr(lSequence)

        'Delete the record from the recordset.
        TableAccess(m_clsTargetTable).GetRecordSet.Delete

        'Remove the filter.
        TableAccess(m_clsTargetTable).CancelFilter
        
        'Remove the item from the listview.
        lvTargetDetails.ListItems.Remove (lvTargetDetails.SelectedItem.Index)
        
        'Update the edit/delete buttons to reflect the current selection.
        EnableEditControls False
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdEditTarget_Click
' Description   : Open the Target Details form to edit the currently selected record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditTarget_Click()
    
    On Error GoTo Failed
    
    'Show the target details form in Add mode.
    ShowTargetPopup True
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Remove any records from the underlying table object which are in the delete
'                 collection. All other changes have already taken place to the underlying
'                 table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    On Error GoTo Failed
         
    'Hide this form and return control to the opener.
    Hide

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Delete any records from the recordset which are about to be added.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    
    On Error GoTo Failed
        
    'Hide the form and return control to the opener.
    Hide
        
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the form's status code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Set the success/failure code of the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Update the screen controls with values from the underlying tables.
'                 This is constant data (combo lists, listview contents, etc).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    
    'Setup the column headers on the Targets listview.
    SetLVHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLVHeaders
' Description   : Add the column headers to the Target ListView.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLVHeaders()
    
    Dim colHeaders As Collection
    Dim lvcolHeaders As listViewAccess
        
    On Error GoTo Failed
    
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    'ActiveFrom header.
    lvcolHeaders.nWidth = 25
    lvcolHeaders.sName = "ActiveFrom"
    colHeaders.Add lvcolHeaders
    
    'ActiveTo header.
    lvcolHeaders.nWidth = 25
    lvcolHeaders.sName = "ActiveTo"
    colHeaders.Add lvcolHeaders
    
    'Target header.
    lvcolHeaders.nWidth = 40
    lvcolHeaders.sName = "Target"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the target listview.
    lvTargetDetails.AddHeadings colHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Setup the form so it can be used.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
        
    'Setup the screen controls (listheaders and the like).
    PopulateScreenControls
    
    'Configure the screen controls according to form state.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Populate the screen controls from the underlying data.
    SetScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup the controls to function in an 'Add' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup the controls to function in an 'Edit' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTarget
' Description   : Store a reference to the Target table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTarget(ByRef clsTargetTable As IntermediaryTargetTable)
    Set m_clsTargetTable = clsTargetTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates the form with the values from the table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetScreenFields(Optional bClear As Boolean = False)
    
    On Error GoTo Failed
    
    'Sorts the records by FromDate and then Sequence Number.
    m_clsTargetTable.OrderRecords
    
    'Populate the listview from the underlying data.
    g_clsFormProcessing.PopulateFromRecordset lvTargetDetails, m_clsTargetTable, bClear
    
    'Only enable the edit/delete controls if an item is selected.
    EnableEditControls (Not lvTargetDetails.SelectedItem Is Nothing)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ShowTargetPopup
' Description   : Displays the Target popup screen with either the current record (in edit
'                 mode) or a new record (in add mode).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowTargetPopup(ByVal bIsEdit As Boolean)
                
    Dim lSequence As Long
    Dim bFilter As Boolean
    Dim frmTarget As frmEditTargetDetails
    Dim clsDetails As PopulateDetails
    
    On Error GoTo Failed
    
    'Use a local variable to talk to the form (pointless).
    Set frmTarget = frmEditTargetDetails
    
    'Default to no record filter?
    bFilter = False
    
    'Associate the table, edit/add mode and keys collection with the popup form.
    frmTarget.SetTarget m_clsTargetTable
    frmTarget.SetIsEdit bIsEdit
    frmTarget.SetKeys m_colKeys
        
    'If we're editing the selected record.
    If bIsEdit Then
        'Get a populated details object for the current listview selection.
        Set clsDetails = lvTargetDetails.GetExtra(lvTargetDetails.SelectedItem.Index)
        
        'Get the sequence number for the selected Target.
        lSequence = clsDetails.GetExtra()
        
        'Filter the recordset so that only this record is visible.
        m_clsTargetTable.FindRecord CStr(lSequence)
        
        'Flag the filter as being on.
        bFilter = True
    End If
    
    'Show the Target Details form modally.
    frmTarget.Show vbModal
        
    If bFilter Then
        'Remove the filter, if there was one.
        TableAccess(m_clsTargetTable).CancelFilter
    End If
    
    'Was OK or cancel pressed on the popup form
    If frmTarget.GetReturnCode = MSGSuccess Then
        'If changes were made, then refresh the listview.
        SetScreenFields True
    End If
    
    'Unload the Target Details dialog.
    Unload frmTarget

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy-up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
    
    'Release object references.
    Set m_colKeys = Nothing
    Set m_clsTargetTable = Nothing
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvTargetDetails_DblClick
' Description   : Routine double-clicks to cmdEdit_Clicks.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvTargetDetails_DblClick()
    cmdEditTarget = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableEditControls
' Description   : Enable/Disable both the edit and delete buttons.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableEditControls(Optional ByVal bEnable As Boolean = False)
    
    On Error GoTo Failed
    
    cmdEditTarget.Enabled = bEnable
    cmdDeleteTarget.Enabled = bEnable

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvTargetDetails_MouseUp
' Description   : Update the edit/delete button's enabled state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvTargetDetails_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Only enable the edit/delete controls if an item is selected.
    EnableEditControls (Not lvTargetDetails.SelectedItem Is Nothing)

End Sub
