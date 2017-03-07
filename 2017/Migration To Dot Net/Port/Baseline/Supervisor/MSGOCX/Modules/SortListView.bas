Attribute VB_Name = "SortListView"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        : SortListView
' Description   : Sorting Routine for listview control.
'                 Extract from MSDN Article 'HOWTO: Sort a ListView Control by Date [Q170884]'
'                 This article can be found in the documentation folder in SS
'                 The listview will only sort by string, so this is required to enable
'                 numeric/date sorting.
'
' Change history
' Prog      Date        Description
'  AA       18/01/01    Added module
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit
    'Column to sort
    Public g_nSortCol As Integer
    'Which order to sort the listview
    Public g_nSortOrder As SortOrder
    'The index of the column last pressed
    Public g_nPreviousCol As Integer
    'Text of the current compare item in the listview
    Private m_sItem As String
    
    'Structures
    'Defines the type of sort
    Public Enum SortOrder
        srtAsc = 0
        srtDes
    End Enum
         
    Public Type POINT
      x As Long
      y As Long
    End Type
 
    Public Type LV_FINDINFO
      flags As Long
      psz As String
      lParam As Long
      pt As POINT
      vkDirection As Long
    End Type
    'Contains all the information about the listview item that the API needs
    Public Type LV_ITEM
      mask As Long
      iItem As Long
      iSubItem As Long
      State As Long
      stateMask As Long
      pszText As Long
      cchTextMax As Long
      iImage As Long
      lParam As Long
      iIndent As Long
    End Type

     'Constants
    Private Const LVFI_PARAM = 1
    Private Const LVIF_TEXT = &H1

    Private Const LVM_FIRST = &H1000
    Private Const LVM_FINDITEM = LVM_FIRST + 13
    Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
    Public Const LVM_SORTITEMS = LVM_FIRST + 48
 
    'API declarations
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SortListViewItems
' Description   : Compares the values in the listview returned by ListView_GetItemData,
'                 and returns an appropriate value, specified by the compare
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function SortListViewItems(ByVal lngParam1 As Long, ByVal lngParam2 As Long, ByVal hWnd As Long) As Long
 
    Dim strName1 As String
    Dim strName2 As String
    Dim nVal1 As Double
    Dim nVal2 As Double
    On Error GoTo Failed
    'Get the value of the first item to be sorted as Double, works in reverse order,
    'ie if the listview contains 3 value 10,20,30, then 20,30 will be compared, then 10, 20
    ListView_GetItemData lngParam1, hWnd, strName1, nVal1
    'Get the value of the next item to be sorted as Double
    ListView_GetItemData lngParam2, hWnd, strName2, nVal2
 
    'Compare the values
    'Return 0 ==> Less Than
    '       1 ==> Equal
    '       2 ==> Greater Than
 
    If nVal1 < nVal2 Then
        If g_nSortOrder = srtAsc Then
            SortListViewItems = 0
        Else
            SortListViewItems = 2
        End If
    ElseIf nVal1 = nVal2 Then
        SortListViewItems = 1
    Else
        If g_nSortOrder = srtAsc Then
            SortListViewItems = 2
        Else
            SortListViewItems = 0
        End If
    End If
    Exit Function
Failed:
    Err.Raise Err.Number, , Err.Description
    Err.Clear
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ListView_GetItemData
' Description   : Gets the value of the item in the listview and converts ot to the appropriate data type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ListView_GetItemData(lngParam As Long, hWnd As Long, strName As String, nItem As Double)
     
    Dim objFind As LV_FINDINFO
    Dim lngIndex As Long
    Dim objItem As LV_ITEM
    Dim baBuffer(32) As Byte
    Dim lngLength As Long
    Dim vtmpListItem As Variant
    Dim dDate As Date
    On Error GoTo Failed
    
     '
    ' Convert the input parameter to an index in the list view
    '
    objFind.flags = LVFI_PARAM
    objFind.lParam = lngParam
    lngIndex = SendMessage(hWnd, LVM_FINDITEM, -1, VarPtr(objFind))
 
       '
    ' Obtain the name of the specified list view item
    '
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = 0
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                            VarPtr(objItem))
    strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)
    
     '
    ' Obtain the modification date of the specified list view item
    '
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = g_nSortCol
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                            VarPtr(objItem))
    If lngLength > 0 Then
        m_sItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        vtmpListItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        If m_sItem <> Space(Len(m_sItem)) Then
            'This
            vtmpListItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
            
            If IsDate(vtmpListItem) And InStr(1, vtmpListItem, ".") = 0 Then
                'Ths can be converted to a date
                dDate = CDate(vtmpListItem)
                nItem = CDbl(dDate)
            Else
                'This number is to big to be a date
                nItem = CDbl(vtmpListItem)
            End If
        Else
            nItem = 0
        End If
    Else
        nItem = 0
    End If
    
    Exit Sub
Failed:
    Err.Raise Err.Number, Err.Source, Err.Description
    Err.Clear
 End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ListView_GetListItem
' Description   : Sorts the ListItems Collection (Not need for Supervisor - AA 18/01/01)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ListView_GetListItem(lngIndex As Long, hWnd As Long, strName As String, nItem As Double)
    Dim objItem As LV_ITEM
    Dim baBuffer(32) As Byte
    Dim lngLength As Long
    Dim dDate As Date
    Dim vtmpListItem As Variant
    On Error GoTo Failed

    objItem.mask = LVIF_TEXT
    objItem.iSubItem = 0
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, VarPtr(objItem))
    strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)

    objItem.mask = LVIF_TEXT
    objItem.iSubItem = g_nSortCol
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                             VarPtr(objItem))
    If lngLength > 0 Then
        m_sItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        vtmpListItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
        If m_sItem <> Space(Len(m_sItem)) Then
            'This
            vtmpListItem = Left$(StrConv(baBuffer, vbUnicode), lngLength)
            
            If IsDate(vtmpListItem) And InStr(1, vtmpListItem, ".") = 0 Then
                'Ths can be converted to a date
                dDate = CDate(vtmpListItem)
                nItem = CDbl(dDate)
            Else
                'This number is to big to be a date
                nItem = CDbl(vtmpListItem)
            End If
        Else
            nItem = 0
        End If
    Else
        nItem = 0
    End If
    
    Exit Sub
Failed:
    Err.Raise Err.Number, Err.Source, Err.Description
    Err.Clear
End Sub


