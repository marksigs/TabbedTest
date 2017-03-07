VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIntCrossSelling 
   Caption         =   "Add/Edit Cross Selling Details"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   Icon            =   "frmEditIntCrossSelling.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MSGOCX.MSGListView lvInsuranceType 
      Height          =   2595
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4577
      Sorted          =   -1  'True
      AllowColumnReorder=   0   'False
   End
End
Attribute VB_Name = "frmEditIntCrossSelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditInCrossSelling
' Description   :   Form used to configure which insurance types cannot be sold by
'                   the intermediary.
'
' Change history
' Prog      Date        Description
' STB       07/12/01    SYS2550 SQL Server support.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'These table object references are set by the interemediary form
'prior to opening this form.
Private m_clsCrossSelling As IntCrossSellingTable

'An intermediary helper-class to provide us with some 'global' utils.
Private m_clsIntermediary As Intermediary


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    'Validate and save the record.
    bRet = DoOkProcessing

    If bRet Then
        'Hide the form and return control to the caller.
        Hide
    End If

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLVHeaders
' Description   : Setup the column colHeaders on the ListViews.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLVHeaders()
    
    Dim colHeaders As Collection
    Dim lvHeaders As listViewAccess
    
    On Error GoTo Failed
    
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    'Correspondence column header.
    lvHeaders.nWidth = 90
    lvHeaders.sName = "Insurance Type"
    colHeaders.Add lvHeaders
    
    'Add the column header to the insurance type listview.
    lvInsuranceType.AddHeadings colHeaders
    
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
    
    'Create a helper-class with some 'global' util routines in.
    Set m_clsIntermediary = New Intermediary
    
    'Populate 'static' data.
    PopulateScreenControls
    
    'Configure the screen controls according to form state.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Select items based upon the underlying table object's contents.
    CheckSelectedItems
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetCrossSelling
' Description   : Store a reference to the cross-selling table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetCrossSelling(ByRef clsCrossSelling As IntCrossSellingTable)
    Set m_clsCrossSelling = clsCrossSelling
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Update the screen controls with values from the underlying tables.
'                 This is constant data (combo lists, listview contents, etc).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    
    On Error GoTo Failed
    
    'Setup the column headers.
    SetLVHeaders
    
    'The MSGListView won't update until run-time so we have to set these properties in code.
    lvInsuranceType.Checkboxes = True

    'Add a row to the list view for each item in the specified Combo Group.
    m_clsIntermediary.PopulateLV "IntermediaryInsuranceType", lvInsuranceType

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIsEdit
' Description   : Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(ByVal bIsEdit As Boolean)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup the controls to function in an 'Edit' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup the controls to function in an 'Add' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckSelectedItems
' Description   : Checks all items with a ValueID in the table relating to the LV
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckSelectedItems()
    
    Dim nValueID As Long
    Dim nThisItem As Integer
    Dim lvItem As MSComctlLib.ListItem
    
    On Error GoTo Failed
        
    'Ensure there is data to add.
    If TableAccess(m_clsCrossSelling).RecordCount > 0 Then
        'Ensure we're starting from the begining.
        TableAccess(m_clsCrossSelling).MoveFirst
        
        'Add each record into the listview.
        For nThisItem = 1 To TableAccess(m_clsCrossSelling).RecordCount
            'Get the unique id for the current record.
            nValueID = m_clsCrossSelling.GetInsuranceType
            
            'Get a reference to the item in the ListView. Use the helper-class
            'to get the unique id (which will be prefixed with INT).
            Set lvItem = lvInsuranceType.ListItems.Item(m_clsIntermediary.GetKey(nValueID))
            
            'Check the item.
            lvItem.Checked = True
            
            'Move onto the next record.
            TableAccess(m_clsCrossSelling).MoveNext
        Next
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOkProcessing
' Description   : Validate and save the current record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
    
    Dim bSuccess As Boolean
    
    On Error GoTo Failed

    'Validate all mandatory fields have been supplied.
    bSuccess = g_clsFormProcessing.DoMandatoryProcessing(Me)
        
    If bSuccess Then
        'If the data is valid, copy the values into the table object(s)
        SaveScreenData
    End If
    
    'Return True (success) or False (Failure) to the opener.
    DoOkProcessing = bSuccess
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : For each checked item in the LV, write the valueid to the tableclass
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()

    Dim vValueID As Variant
    Dim colValueIDs As Collection
    
    On Error GoTo Failed
    
    'If there any existing, selected items, delete them all.
    If TableAccess(m_clsCrossSelling).RecordCount > 0 Then
        TableAccess(m_clsCrossSelling).DeleteAllRows
    End If
    
    'Get a collection of selected ValueIDs.
    Set colValueIDs = m_clsIntermediary.GetCheckedItemsAsCollection(lvInsuranceType)
    
    For Each vValueID In colValueIDs
        'Add a row for each selected item.
        TableAccess(m_clsCrossSelling).AddRow
        
        'Set the field values now.
        m_clsCrossSelling.SetInsuranceType vValueID
        m_clsCrossSelling.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
    Next

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Unload
' Description   : Release object references and tidy up.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)

    'Release object references.
    Set m_colKeys = Nothing
    Set m_clsCrossSelling = Nothing
    Set m_clsIntermediary = Nothing

End Sub
