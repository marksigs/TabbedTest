VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MSGListView 
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   3690
   ScaleWidth      =   7740
   Begin MSComctlLib.ListView lvListView 
      Height          =   2700
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   4763
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "MSGListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Control : MSGListView
' Description  : ActiveX control to provider extra functionality to the ListView control
'
' Change history
' Prog      Date        Description
' AA        20/06/2001  SYS2000 - Added Checkboxes property to the control
' DJP       07/11/2001  SYS2831 - Put FieldData and listViewAccess in here instead
'                       of lvHeadings.dll
' DJP       14/11/2001  SYS2831 - When resizing resize listview to slightly less than the width
'                       of the User Control otherwise the right border is missing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'Property Variables:
Dim m_Checkboxes As Boolean
Private m_SortKey As Integer
Private m_SortOrder As ListSortOrderConstants
Private m_TagObject As Object
Private m_Extra As Object

'Event Declarations:
Event DeletePressed()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvListView,lvListView,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lvListView,lvListView,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=lvListView,lvListView,-1,ItemClick
Attribute ItemClick.VB_Description = "Occurs when a ListItem object is clicked or selected"
Event DblClick() 'MappingInfo=lvListView,lvListView,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event Click() 'MappingInfo=lvListView,lvListView,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

'Default Property Values:
Const m_def_Checkboxes = 0
Const m_def_SortKey = 0
Const m_def_SortOrder = 0

' Constants
Private Const COL_POS As String = "ColumnPosition"
Private Const COL_WIDTH As String = "ColumnWidth"

' Private data
Private m_sRegLocation As String
Private m_clsErrorHandling As ErrorHandling

Public Type listViewAccess
    sName As String
    nWidth As Integer
End Type

Public Type FieldData
    sField As String
    bRequired As Boolean
    bVisible As Boolean
    bDisabled As Boolean
    sDefault As String
    sError As String
    sTitle As String
    colComboValues As Collection
    colComboIDS As Collection
    sOtherField As String
    bDateField As Boolean
    sMinValue As String
    sMaxValue As String
End Type

' Gets the index of the currently selection column. sVal will contain the index, the function
' returns true if the index is returned, false if not

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Control : MSGListView
' Description  : ActiveX control to provider extra functionality to the ListView control
'
' BMIDS Specific Change history
' Prog      Date        Description
' DB        06/01/2003  BM0060 - Added new method IsColDate to check if date column clicked on and if yes, refresh
'                                the index list to reflect this.
' BS        08/04/2003  BM0497 Amended date/numeric column sort
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Function GetSelectedColumn(nColumn As Integer, sVal As String) As Boolean
    Dim colLine As Collection
    Dim bRet As Boolean
    On Error GoTo Failed

    bRet = False
    Set colLine = GetLine()
    
    If Not colLine Is Nothing Then
        If nColumn < lvListView.ColumnHeaders.Count Then
            sVal = colLine(nColumn)
            bRet = True
        End If
    End If
    
    GetSelectedColumn = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Function GetLine(Optional nIndex As Integer = -1, Optional objExtra As Object = Nothing) As Collection
    On Error GoTo Failed
    Dim listLine As ListItem
    Dim nSelectedLine As Integer
    Dim colLine As New Collection
    
    If (nIndex <> -1) Then
        Set listLine = lvListView.ListItems(nIndex)
    Else
        Set listLine = lvListView.SelectedItem
    End If
    
    If (Not listLine Is Nothing) Then
        
        Dim nEntry As Integer
        Dim nNumEntries As Integer
        
        If IsObject(listLine.Tag) Then
            If Not listLine.Tag Is Nothing Then
                Set objExtra = listLine.Tag
            End If
        End If
        
        nNumEntries = lvListView.ColumnHeaders.Count
        
        For nEntry = 0 To nNumEntries - 1
            If (nEntry = 0) Then
                colLine.Add listLine.Text
            Else
                colLine.Add listLine.SubItems(nEntry)
            End If
        Next
        Set GetLine = colLine
    Else
        Set GetLine = Nothing
    End If
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Sub AddLine(colLine As Collection, Optional objExtra As Object = Nothing)
    On Error GoTo Failed
    Dim nEntry As Long
    Dim nNumEntries As Long
    Dim listLine As ListItem
    Dim sStripLineFeed As String
    Dim sStripCarriageReturn As String
    
    Set listLine = UserControl.lvListView.ListItems.Add()
    Set listLine.Tag = objExtra
     
    nNumEntries = colLine.Count

    For nEntry = 1 To nNumEntries
        
        If InStr(1, colLine(nEntry), Chr(13)) > 0 Then
            'There are some Carraige returns in the text, so replace all LineFeed Characters with ""
            sStripLineFeed = Replace(colLine(nEntry), Chr(10), "")
        Else
            'No Carraige Returns
            sStripLineFeed = colLine(nEntry)
        End If
        
        sStripCarriageReturn = Replace(sStripLineFeed, Chr(13), " ")

        If (nEntry = 1) Then
            listLine.Text = sStripCarriageReturn
        Else
            listLine.SubItems(nEntry - 1) = sStripCarriageReturn
        End If
    Next
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Function GetExtra(nLine As Integer) As Object
    Set GetExtra = UserControl.lvListView.ListItems(nLine).Tag
End Function

Public Sub AddHeadings(colHeadings As Collection)
    On Error GoTo Failed
    Dim nTotal As Integer
    Dim nCount As Integer
    Dim dWidth As Double
    Dim nColumnWidth As Integer
    Dim sColName As String
    Dim dOnePercent As Double
    Dim colHeading As listViewAccess
    
    dOnePercent = UserControl.lvListView.Width / 100
    
    nTotal = colHeadings.Count
    
    For nCount = 1 To nTotal
        colHeading = colHeadings(nCount)
        sColName = colHeading.sName
        nColumnWidth = colHeading.nWidth
        dWidth = nColumnWidth * dOnePercent
        
        ' Set the header with the information we have obtained
        SetFieldHeader sColName, dWidth
    Next nCount
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Sub SetFieldHeader(sName As String, nWidth As Double)
    On Error GoTo Failed
    Dim colX As ColumnHeader
    
    Set colX = lvListView.ColumnHeaders.Add()
    
    colX.Text = sName
    colX.Width = nWidth
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Function SaveColumnDetails() As String
    On Error GoTo Failed
    Dim colHeader As ColumnHeader
    Dim nPosition As Long
    Dim nWidth As Long
    
    If Len(m_sRegLocation) > 0 Then
        For Each colHeader In lvListView.ColumnHeaders
            nPosition = colHeader.Position
            nWidth = colHeader.Width
        
            SetKeyValue HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, m_sRegLocation & COL_WIDTH & colHeader.Index, CStr(nWidth), REG_SZ
            SetKeyValue HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, m_sRegLocation & COL_POS & colHeader.Index, CStr(nPosition), REG_SZ
        Next
    End If
    
    SaveColumnDetails = m_sRegLocation
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Sub LoadColumnDetails(sRegLocation As String)
    On Error GoTo Failed
    Dim colHeader As ColumnHeader
    Dim sWidth As String
    Dim sPosition As String
    
    m_sRegLocation = sRegLocation
    
    For Each colHeader In lvListView.ColumnHeaders
        sWidth = QueryValue(HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, sRegLocation & COL_WIDTH & colHeader.Index)
        sPosition = QueryValue(HKEY_LOCAL_MACHINE, SUPERVISOR_KEY, sRegLocation & COL_POS & colHeader.Index)
        
        If IsNumeric(sWidth) Then
            colHeader.Width = CLng(sWidth)
        End If
        
        If IsNumeric(sPosition) Then
            colHeader.Position = CLng(sPosition)
        End If
    Next
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub

Public Function GetValueFromName(sName As String) As String
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sVal As String
    Dim nCol As Integer
    Dim nCount As Integer
    Dim sField As String
    Dim colX As ColumnHeader ' Declare variable.
        
    nCount = lvListView.ColumnHeaders.Count()

    For Each colX In lvListView.ColumnHeaders
        sField = colX.Text

        If UCase(sField) = UCase(sName) Then
            bRet = GetSelectedColumn(colX.Index, sVal)
            Exit For
        End If
    Next

    GetValueFromName = sVal
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Private Sub UserControl_Initialize()
    Set m_clsErrorHandling = New ErrorHandling
End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim strName As String
    Dim lngItem As Long
    Dim bNumericCol As Boolean
    Dim bDateSort As Boolean
    Dim nItem As Double
    'BS BM0497 08/04/03
    Dim nColNo As Long
    Dim varitm As ListItem
    Dim strFormat As String
    'BS BM0497 End 08/04/03
    
    g_nSortCol = ColumnHeader.Index - 1
    
    If lvListView.SortOrder = lvwAscending And g_nPreviousCol = ColumnHeader.Index Then
        'The user has clicked the same column header twice, therefore do a descending sort
        g_nSortOrder = srtDes
        lvListView.SortOrder = lvwDescending
    Else
        'Sort the column ascending
        g_nSortOrder = srtAsc
        lvListView.SortOrder = lvwAscending
    End If
    
    bNumericCol = IsColNumeric(ColumnHeader.Index)
    'DB 06/01/03 BM0060 - Added new call to check whether date column clicked on.
    bDateSort = IsColDate(ColumnHeader.Index)
    
    'BS BM0497 08/04/03
    'bNumericCol is true if the column is a date but bDateSort is false if the column is numeric,
    'so if column is character then bNumericCol and bDateSort will both be false
    
    lvListView.Sorted = True
    
'DoStringSort:
    
    'If Not bNumericCol Or bDateSort Then
    If Not bNumericCol And Not bDateSort Then
DoStringSort:
        'lvListView.Sorted = True
        'BS BM0497 End 08/04/03
        lvListView.SortKey = ColumnHeader.Index - 1
    Else
        On Error GoTo DoStringSort
        'bDateSort is set to true if the sort fails. This way the column will be sorted by string
        'using the on error event.
        'BS BM0497 08/04/03
        'Using the sort APIs correctly sorts the ListView but the index values of the items in the
        'ListItems collection are not updated and are therefore out of sync with the ListView.
        'The index values are used for promotion and deletion, so the wrong rows were being
        'promoted/deleted. As the ListView is multi-select, the only solution is to use the character
        'sort which does update the index values. To do this a temporary column is added to the
        'ListView and the values in the selected date/numeric column are formatted and copied to
        'the new column. The ListView is then sorted alphabetically on this new column, and then
        'the column is deleted.
        
'        bDateSort = True
'        lvListView.Sorted = False
'        SendMessage lvListView.hWnd, LVM_SORTITEMS, lvListView.hWnd, AddressOf SortListViewItems
        
        lvListView.ColumnHeaders.Add , "DummySort", ""
        nColNo = lvListView.ColumnHeaders("DummySort").SubItemIndex
        
        If bDateSort Then
            'Column contains dates
            For Each varitm In lvListView.ListItems
                If ColumnHeader.Index > 1 Then
                    varitm.SubItems(nColNo) = Format(CDate(varitm.SubItems(ColumnHeader.Index - 1)), "yyyymmddhhmmss")
                Else
                    varitm.SubItems(nColNo) = Format(CDate(varitm.Text), "yyyymmddhhmmss")
                End If
            Next
        Else
            'Column contains numeric values
            strFormat = String$(20, "0") & "." & String$(10, "0")

            For Each varitm In lvListView.ListItems
                If ColumnHeader.Index > 1 Then
                    If Len(varitm.SubItems(ColumnHeader.Index - 1)) = 0 Then
                        'Value is blank, so set sort column to blank so they are sorted differently to 0
                        varitm.SubItems(nColNo) = ""
                    ElseIf CDbl(varitm.SubItems(ColumnHeader.Index - 1)) >= 0 Then
                            'Value is positive
                            varitm.SubItems(nColNo) = Format(CDbl(varitm.SubItems(ColumnHeader.Index - 1)), strFormat)
                        Else
                            'Value is negative, so need to inverse it to get the correct order
                            varitm.SubItems(nColNo) = "&" & InvNumber(Format(0 - CDbl(varitm.SubItems(ColumnHeader.Index - 1)), strFormat))
                    End If
                Else
                     If Len(varitm.Text) = 0 Then
                        'Value is blank, so set sort column to blank so they are sorted differently to 0
                        varitm.SubItems(nColNo) = ""
                    ElseIf CDbl(varitm.Text) >= 0 Then
                            'Value is positive
                            varitm.SubItems(nColNo) = Format(CDbl(varitm.Text), strFormat)
                        Else
                            'Value is negative, so need to inverse it to get the correct order
                            varitm.SubItems(nColNo) = "&" & InvNumber(Format(0 - CDbl(varitm.Text), strFormat))
                    End If
                End If
            Next
        End If
        
        lvListView.SortKey = nColNo
        'BS BM0497 End 08/04/03
    End If
    On Error GoTo Failed
    lvListView.Refresh
    'BS BM 08/04/03
    'Commented back out - it doesn't actually do anything we want, it's just an example of how
    'to use the APIs to get values from the ListView control.
    'DB 06/01/03 BM0060 - Commented back in as per CBS.
    'Not needed for how the listview is used in supervisor
'    If bNumericCol And Not bDateSort Then
'         For lngItem = 0 To lvListView.ListItems.Count - 1
'          ListView_GetListItem lngItem, lvListView.hWnd, strName, nItem
'        Next
'    End If

    'Remove temporary sort column if present
    If lvListView.ColumnHeaders.Count - 1 = nColNo Then
        lvListView.ColumnHeaders.Remove ("DummySort")
    End If

    'BS BM0497 End 08/04/03
    
    g_nPreviousCol = ColumnHeader.Index
    Exit Sub

Failed:
    m_clsErrorHandling.DisplayError
End Sub

'BS BM0497 08/04/03
'Formats negative numbers to ensure correct sort order
Private Function InvNumber(ByVal Number As String) As String
    Dim nPos As Integer
    For nPos = 1 To Len(Number)
        Select Case Mid$(Number, nPos, 1)
            Case "-": Mid$(Number, nPos, 1) = " "
            Case "0": Mid$(Number, nPos, 1) = "9"
            Case "1": Mid$(Number, nPos, 1) = "8"
            Case "2": Mid$(Number, nPos, 1) = "7"
            Case "3": Mid$(Number, nPos, 1) = "6"
            Case "4": Mid$(Number, nPos, 1) = "5"
            Case "5": Mid$(Number, nPos, 1) = "4"
            Case "6": Mid$(Number, nPos, 1) = "3"
            Case "7": Mid$(Number, nPos, 1) = "2"
            Case "8": Mid$(Number, nPos, 1) = "1"
            Case "9": Mid$(Number, nPos, 1) = "0"
        End Select
    Next
    InvNumber = Number
End Function
'BS BM0497 End 08/04/03

Private Function IsColNumeric(nColIndx As Integer) As Boolean

    Dim bRet As Boolean
    Dim sItem As String
    Dim iCnt As Integer
    
    bRet = False
    sItem = ""
    
    If lvListView.ListItems.Count > 0 Then
        If nColIndx > 1 Then
            'Get the item text
            sItem = lvListView.ListItems(1).SubItems(nColIndx - 1)
        Else
            'Get the SubItem text
            sItem = lvListView.ListItems(1).Text
        End If
    End If
    
    If IsDate(sItem) Or IsNumeric(sItem) Or sItem = Space(Len(sItem)) Then
        'This can be converted to a date!!!!!
        bRet = True
    Else
        'This is a string!!!!!
        bRet = False
    End If
    
    IsColNumeric = bRet
    
End Function

Private Sub lvListView_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Failed
    If KeyCode = vbKeyReturn Then
        If Not lvListView.SelectedItem Is Nothing Then
            RaiseEvent DblClick
        End If
    ElseIf KeyCode = vbKeyDelete Then
        RaiseEvent DeletePressed
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub UserControl_Resize()
    UserControl.lvListView.Width = UserControl.Width - 100
    UserControl.lvListView.Height = UserControl.Height

End Sub

Public Property Get View() As ListViewConstants
Attribute View.VB_Description = "Returns/sets the current view of the ListView control."
    View = lvListView.View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
    lvListView.View() = New_View
    PropertyChanged "View"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lvListView.View = PropBag.ReadProperty("View", 3)
    Set m_Extra = PropBag.ReadProperty("Extra", Nothing)
    lvListView.Enabled = PropBag.ReadProperty("Enabled", True)
    lvListView.HideSelection = PropBag.ReadProperty("HideSelection", True)
    lvListView.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
    lvListView.Sorted = PropBag.ReadProperty("Sorted", False)
    Set m_TagObject = PropBag.ReadProperty("TagObject", Nothing)
    m_SortKey = PropBag.ReadProperty("SortKey", m_def_SortKey)
    m_SortOrder = PropBag.ReadProperty("SortOrder", m_def_SortOrder)
    lvListView.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    lvListView.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder", True)
    Set SmallIcons = PropBag.ReadProperty("SmallIcons", Nothing)
    Set Icons = PropBag.ReadProperty("Icons", Nothing)
    m_Checkboxes = PropBag.ReadProperty("Checkboxes", m_def_Checkboxes)
End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Failed
    
    If Len(m_sRegLocation) > 0 Then
        ' Save the position and with of the columns in the registry
        SaveColumnDetails
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("View", lvListView.View, 3)
    Call PropBag.WriteProperty("Extra", m_Extra, Nothing)
    Call PropBag.WriteProperty("Enabled", lvListView.Enabled, True)
    Call PropBag.WriteProperty("HideSelection", lvListView.HideSelection, True)
    Call PropBag.WriteProperty("HideColumnHeaders", lvListView.HideColumnHeaders, False)
    Call PropBag.WriteProperty("Sorted", lvListView.Sorted, False)
    Call PropBag.WriteProperty("TagObject", m_TagObject, Nothing)
    Call PropBag.WriteProperty("SortKey", m_SortKey, m_def_SortKey)
    Call PropBag.WriteProperty("SortOrder", m_SortOrder, m_def_SortOrder)
    Call PropBag.WriteProperty("MultiSelect", lvListView.MultiSelect, False)

    Call PropBag.WriteProperty("AllowColumnReorder", lvListView.AllowColumnReorder, True)
    Call PropBag.WriteProperty("SmallIcons", SmallIcons, Nothing)
    Call PropBag.WriteProperty("Icons", Icons, Nothing)
    Call PropBag.WriteProperty("Checkboxes", m_Checkboxes, m_def_Checkboxes)
End Sub

Public Property Get SelectedItem() As Object
    Set SelectedItem = lvListView.SelectedItem
    
End Property

Public Property Set SelectedItem(ByVal New_SelectedItem As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    Set lvListView.SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
End Property

Private Sub lvListView_Click()
    RaiseEvent Click
End Sub

Public Function RemoveLine(nLine As Long) As Variant
    lvListView.ListItems.Remove (nLine)
End Function

Public Function ListItems() As ListItems
    Set ListItems = lvListView.ListItems
End Function

Public Property Get Extra() As Object
    Set Extra = m_Extra
End Property

Public Property Set Extra(ByVal New_Extra As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    Set m_Extra = New_Extra
    PropertyChanged "Extra"
End Property

Private Sub lvListView_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = lvListView.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lvListView.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the ListView loses focus"
    HideSelection = lvListView.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    lvListView.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

Private Sub lvListView_ItemClick(ByVal Item As ListItem)
    RaiseEvent ItemClick(Item)
End Sub

Public Property Get ColumnHeaders() As IColumnHeaders
Attribute ColumnHeaders.VB_Description = "Returns a reference to a collection of ColumnHeader objects."
    Set ColumnHeaders = lvListView.ColumnHeaders
End Property

Public Property Get HideColumnHeaders() As Boolean
Attribute HideColumnHeaders.VB_Description = "Returns/sets whether or not a ListView control's column headers are hidden in Report view."
    HideColumnHeaders = lvListView.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    lvListView.HideColumnHeaders() = New_HideColumnHeaders
    PropertyChanged "HideColumnHeaders"
End Property

Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lvListView.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lvListView.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

Public Property Get TagObject() As Object
Attribute TagObject.VB_MemberFlags = "400"
    Set TagObject = m_TagObject
End Property

Public Property Set TagObject(ByVal New_TagObject As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    Set m_TagObject = New_TagObject
    PropertyChanged "TagObject"
End Property

Public Property Get SortKey() As Integer
Attribute SortKey.VB_Description = "Returns/sets the current sort key."
    SortKey = m_SortKey
End Property

Public Property Let SortKey(ByVal New_SortKey As Integer)
    m_SortKey = New_SortKey
    PropertyChanged "SortKey"
End Property

Public Property Get SortOrder() As ListSortOrderConstants
Attribute SortOrder.VB_Description = "Returns/sets whether or not the ListItems will be sorted in ascending or descending order."
    SortOrder = m_SortOrder
End Property

Public Property Let SortOrder(ByVal New_SortOrder As ListSortOrderConstants)
    m_SortOrder = New_SortOrder
    PropertyChanged "SortOrder"
End Property

Private Sub UserControl_InitProperties()
    m_SortKey = m_def_SortKey
    m_SortOrder = m_def_SortOrder
    m_Checkboxes = m_def_Checkboxes
End Sub

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
    MultiSelect = lvListView.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    lvListView.MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets whether a user can reorder columns in report view."
    AllowColumnReorder = lvListView.AllowColumnReorder
End Property

Public Property Let AllowColumnReorder(ByVal New_AllowColumnReorder As Boolean)
    lvListView.AllowColumnReorder() = New_AllowColumnReorder
    PropertyChanged "AllowColumnReorder"
End Property

Public Property Get SmallIcons() As Object
Attribute SmallIcons.VB_Description = "Returns/sets the images associated with the SmallIcons property of a ListView control."
    Set SmallIcons = lvListView.SmallIcons
End Property

Public Property Set SmallIcons(ByVal New_SmallIcons As Object)
    Set lvListView.SmallIcons = New_SmallIcons
    PropertyChanged "SmallIcons"
End Property

Public Property Get Icons() As Object
Attribute Icons.VB_Description = "Returns/sets the images associated with the Icon properties of a ListView control."
    Set Icons = lvListView.Icons
End Property

Public Property Set Icons(ByVal New_Icons As Object)
    Set lvListView.Icons = New_Icons
    PropertyChanged "Icons"
End Property

Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the list."
    Checkboxes = m_Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
    m_Checkboxes = New_Checkboxes
    PropertyChanged "Checkboxes"
    lvListView.Checkboxes = New_Checkboxes
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsColDate
' Description   : Gets the value of the first item in the listview, and checks if it is
'                 a Date field. Returns false if string
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsColDate(nColIndx As Integer) As Boolean

    Dim bRet As Boolean
    Dim sItem As String
    Dim iCnt As Integer
    
    bRet = False
    sItem = ""
    
    If lvListView.ListItems.Count > 0 Then
        If nColIndx > 1 Then
            'Get the item text
            sItem = lvListView.ListItems(1).SubItems(nColIndx - 1)
        Else
            'Get the SubItem text
            sItem = lvListView.ListItems(1).Text
        End If
    End If
    
    If IsDate(sItem) Then
        'This can be converted to a date!!!!!
        bRet = True
    Else
        'This is a string!!!!!
        bRet = False
    End If
    
    IsColDate = bRet
    
End Function

