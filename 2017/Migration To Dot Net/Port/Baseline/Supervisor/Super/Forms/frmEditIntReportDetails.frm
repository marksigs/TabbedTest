VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditIntReportDetails 
   Caption         =   "Add/Edit Correspondence and Report Details"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmEditIntReportDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Frame fraDailyReportDays 
      Caption         =   "Daily Report Days"
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   3660
      Width           =   6315
      Begin MSGOCX.MSGListView lvReportDays 
         Height          =   2235
         Left            =   360
         TabIndex        =   2
         Top             =   420
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3942
         AllowColumnReorder=   0   'False
         Checkboxes      =   -1  'True
      End
   End
   Begin VB.Frame fraCorrespondenceType 
      Caption         =   "Correspondence Type"
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6315
      Begin MSGOCX.MSGListView lvCorrespondenceType 
         Height          =   2055
         Left            =   360
         TabIndex        =   0
         Top             =   420
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3625
         AllowColumnReorder=   0   'False
         Checkboxes      =   -1  'True
      End
   End
   Begin MSGOCX.MSGComboBox cboReportFrequency 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   3180
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.Label lblIntReport 
      Caption         =   "Report Frequency"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   3180
      Width           =   1515
   End
End
Attribute VB_Name = "frmEditIntReportDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditIntReportDetails
' Description   :   Form which allows the user add and edit Intermediary Report
'                   and Correspondance details.
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

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'These table object references are set by the interemediary form
'prior to opening this form.
Private m_clsIntermediaryTable As IntermediaryTable
Private m_clsCorrespondence As IntCorrespondenceTable
Private m_clsReportDays As IntermediaryReportDaysTable

'An intermediary helper-class to provide us with some 'global' utils.
Private m_clsIntermediary As Intermediary

'Indicates if we're dealing with the correspondance listview or the report one.
Private Enum CorrespondenceType
    TypeReport
    TypeCorrespondence
End Enum


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the current record.
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
' Function      : cmdCancel_Click
' Description   : 'Hide the form and return control to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Hide
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
' Description   : Return the forms status to the opener.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return code to success/failure.
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
    
    On Error GoTo Failed
    
    'Populate the Report frequency combo.
    g_clsFormProcessing.PopulateCombo "IntermediaryReportFrequency", cboReportFrequency, False
    
    'Setup the column headers.
    SetLVHeaders
        
    'The MSGListView won't update until run-time so we have to set these properties in code.
    lvReportDays.Checkboxes = True
    lvCorrespondenceType.Checkboxes = True
        
    'Add a row to the correspondence list view for each item in the specified Combo Group.
    m_clsIntermediary.PopulateLV "IntermediaryCorrespondenceType", lvCorrespondenceType
    
    'Add a row to the report list view for each item in the specified Combo Group.
    m_clsIntermediary.PopulateLV "DaysOfWeek", lvReportDays
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetLVHeaders
' Description   : Setup the column colHeaders on the ListViews.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetLVHeaders()
    
    Dim colHeaders As Collection
    Dim lvcolHeaders As listViewAccess
    
    On Error GoTo Failed
        
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    'Correspondence column header.
    lvcolHeaders.nWidth = 90
    lvcolHeaders.sName = "Type"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the correspondence listview.
    lvCorrespondenceType.AddHeadings colHeaders
    
    'Remove the item from the collection, we'll re-use it.
    colHeaders.Remove 1
    
    'Dailey Report column header.
    lvcolHeaders.nWidth = 90
    lvcolHeaders.sName = "Day"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the report listview.
    lvReportDays.AddHeadings colHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckSelectedItems
' Description   : Checks all items with a ValueID in the table which are present in the
'                 MSGListView specified.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckSelectedItems(ByRef lv As MSGListView, ByRef clsTableAccess As TableAccess, ByVal uType As CorrespondenceType)
    
    Dim nValueID As Long
    Dim nThisItem As Integer
    Dim lvItem As MSComctlLib.ListItem
    
    On Error GoTo Failed
    
    'Ensure there is data to add.
    If clsTableAccess.RecordCount > 0 Then
        'Ensure we're starting from the begining.
        clsTableAccess.MoveFirst
        
        'Add each record into the listview.
        For nThisItem = 1 To clsTableAccess.RecordCount
            'Get the unique id for the current record.
            If uType = TypeCorrespondence Then
                nValueID = m_clsCorrespondence.GetType
            Else
                nValueID = m_clsReportDays.GetDay
            End If
            
            'Get a reference to the item in the ListView. Use the helper-class
            'to get the unique id (which will be prefixed with INT).
            Set lvItem = lv.ListItems.Item(m_clsIntermediary.GetKey(nValueID))
            
            'Check the item.
            lvItem.Checked = True
            
            'Move onto the next record.
            clsTableAccess.MoveNext
        Next
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Setup the controls to function in an 'Add' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    'Stub
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Setup the controls to function in an 'Edit' state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    'Stub
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Setup the form so it can be used.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Form_Load()
    
    On Error GoTo Failed

    'Create an intermediary helper-class to provide us with some 'global' utils.
    Set m_clsIntermediary = New Intermediary
        
    'Populate combos and 'static' data.
    PopulateScreenControls
        
    'Configure the screen controls according to form state.
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    'Populate the controls with data from the underlying table object(s).
    SetScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetCorrespondence
' Description   : Store a reference to the correspondance table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetCorrespondence(ByRef clsCorrespondence As IntCorrespondenceTable)
    Set m_clsCorrespondence = clsCorrespondence
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetDailyReport
' Description   : Store a reference to the dailey report table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetDailyReport(ByRef clsReport As IntermediaryReportDaysTable)
    Set m_clsReportDays = clsReport
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIntermediary
' Description   : Store a reference to the base intermediary table table.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIntermediary(ByRef clsIntermediary As IntermediaryTable)
    Set m_clsIntermediaryTable = clsIntermediary
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
' Description   : Copy screen values into the underlying tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    
    Dim vVal As Variant
    
    On Error GoTo Failed

    'Broker the call to more specific handler routines.
    SaveReportDays
    SaveCorrespondence
        
    'Store the selected value from the report frequency combo in a tempory variable.
    g_clsFormProcessing.HandleComboExtra cboReportFrequency, vVal, GET_CONTROL_VALUE
    
    'Update the underlying field value with the frequency.
    m_clsIntermediaryTable.SetReportFrequency vVal

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveReportDays
' Description   : Updates the table class with all report day records.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveReportDays()
        
    Dim vValueID As Variant
    Dim colValueIDs As Collection
        
    On Error GoTo Failed
        
    'If there any existing, selected items, delete them all.
    If TableAccess(m_clsReportDays).RecordCount > 0 Then
        TableAccess(m_clsReportDays).DeleteAllRows
    End If
    
    'Get a collection of selected ValueIDs.
    Set colValueIDs = m_clsIntermediary.GetCheckedItemsAsCollection(lvReportDays)
        
    For Each vValueID In colValueIDs
        'Add a row for each selected item.
        TableAccess(m_clsReportDays).AddRow
        
        'Set the field values now.
        m_clsReportDays.SetDay vValueID
        m_clsReportDays.SetIntermediaryGuid m_colKeys.Item(INTERMEDIARY_KEY)
    Next

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveCorrespondence
' Description   : For each checked item in the LV, write the ValueID to the tableclass
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveCorrespondence()

    Dim vValueID As Variant
    Dim colValueIDs As Collection
    
    On Error GoTo Failed
    
    'If there any existing, selected items, delete them all.
    If TableAccess(m_clsCorrespondence).RecordCount > 0 Then
        TableAccess(m_clsCorrespondence).DeleteAllRows
    End If
    
    'Get a collection of selected ValueIDs.
    Set colValueIDs = m_clsIntermediary.GetCheckedItemsAsCollection(lvCorrespondenceType)
    
    For Each vValueID In colValueIDs
        'Add a row for each selected item.
        TableAccess(m_clsCorrespondence).AddRow
        
        'Set the field values now.
        m_clsCorrespondence.SetType vValueID
        m_clsCorrespondence.SetIntermediaryGuid m_colKeys.Item(INTERMEDIARY_KEY)
    Next
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Update the screen controls with data from the underlying table objects.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetScreenFields()
    
    Dim vVal As Variant
    
    On Error GoTo Failed
    
    'Set the report frequency combo.
    vVal = m_clsIntermediaryTable.GetReportFrequency
    g_clsFormProcessing.HandleComboExtra cboReportFrequency, vVal, SET_CONTROL_VALUE
    
    'Check Selected Resport Days.
    CheckSelectedItems lvReportDays, m_clsReportDays, TypeReport
        
    'Check Selected Correspondence Types
    CheckSelectedItems lvCorrespondenceType, m_clsCorrespondence, TypeCorrespondence
    
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
    Set m_clsReportDays = Nothing
    Set m_clsIntermediary = Nothing
    Set m_clsCorrespondence = Nothing
    Set m_clsIntermediaryTable = Nothing

End Sub
