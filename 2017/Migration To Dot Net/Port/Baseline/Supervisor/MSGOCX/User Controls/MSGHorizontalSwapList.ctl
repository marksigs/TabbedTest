VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MSGHorizontalSwapList 
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   ScaleHeight     =   3465
   ScaleWidth      =   8445
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSGOCX.MSGListView MSGListView1 
      Height          =   2715
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4789
      Sorted          =   -1  'True
   End
   Begin VB.CommandButton cmdExtra 
      Height          =   375
      Index           =   0
      Left            =   3060
      TabIndex        =   6
      Top             =   2100
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAddSecond 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.CommandButton cmdAddAllToSecond 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   555
   End
   Begin VB.CommandButton cmdAddAllToFirst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1740
      Width           =   555
   End
   Begin VB.CommandButton cmdAddFirst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2340
      Width           =   555
   End
   Begin MSComctlLib.ImageList imlButtons 
      Left            =   3300
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   7
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":0000
            Key             =   "AddOneSecondV"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":0112
            Key             =   "AddAllFirstV"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":029C
            Key             =   "AddAllSecondH"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":043E
            Key             =   "AddAllSecondV"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":05C8
            Key             =   "AddOneFirstH"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":06DE
            Key             =   "AddOneFirstV"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":07D8
            Key             =   "AddOneSecondH"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MSGHorizontalSwapList.ctx":08EE
            Key             =   "AddAllFirstH"
         EndProperty
      EndProperty
   End
   Begin MSGOCX.MSGListView MSGListView2 
      Height          =   2715
      Left            =   3960
      TabIndex        =   8
      Top             =   420
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4789
      Sorted          =   -1  'True
   End
   Begin VB.Label lblListTitle 
      Caption         =   "Selected Items"
      Height          =   255
      Index           =   1
      Left            =   4020
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblListTitle 
      Caption         =   "Available Items"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "MSGHorizontalSwapList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserControl   : MSGHorizontalSwapList
' Description   : ActiveX Control allowing two listviews to swap values between each other
'                 using four buttons

' Change history
' Prog      Date        Description
' DJP       06/11/01    SYS2831 Updated for Core inheritence
' DJP       06/11/01    SYS2831 Updated for Core inheritence - fix problem of buttons not
'                       drawing correctly.
' DJP       27/11/01    SYS2912 SQL Server locking problem.
' STB       29/01/02    SYS2957 Security Enhancement - DoesValueExist() doesn't
'                       ignore listrows with 1+ columns.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Default Property Values:
Const m_def_HasChanged = 0
Const m_def_MoveToSecond = True
Private m_bAllowSizeChange As Boolean

Private WithEvents btnObj As CommandButton
Attribute btnObj.VB_VarHelpID = -1
Event ExtraClick(nIndex As Integer)
Event FirstItemClicked(colRow As Collection)
Event SecondItemClicked(colRow As Collection, objExtra As Object)

'Property Variables:
Dim m_HasChanged As Boolean
Dim m_Extra As Object
Dim m_MoveToSecond As Boolean
' Private data
Private m_bSwapEnabled As Boolean
Private m_clsErrorHandling As ErrorHandling
Private m_bAllowReorder As Boolean
' Enums
Private Enum MoveDirection
    MoveDirectionUp = 1
    MoveDirectionDown
End Enum

' Constants
Private Const LEFT_TITLE As Integer = 0
Private Const RIGHT_TITLE As Integer = 1

'User Defined Access Functions
Public Sub SetFirstTitle(sTitle As String)
    lblListTitle(LEFT_TITLE).Caption = sTitle
End Sub
Public Sub SetFirstSorted(Optional bSorted As Boolean = True)
    MSGListView1.Sorted = bSorted
End Sub
Public Sub SetSecondSorted(Optional bSorted As Boolean = True)
    MSGListView2.Sorted = bSorted
End Sub
Public Sub AllowReorder(Optional bAllowReorder As Boolean = False)
    On Error GoTo Failed
    
    cmdMoveUp.Visible = True
    cmdMoveDown.Visible = True
    m_bAllowReorder = bAllowReorder
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub cmdAddAllToSecond_Click()
    On Error GoTo Failed
    m_HasChanged = True
    CopyAllToList MSGListView1, MSGListView2
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub cmdAddAllToFirst_Click()
    On Error GoTo Failed
    m_HasChanged = True
    CopyAllToList MSGListView2, MSGListView1
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub cmdAddFirst_Click()
    On Error GoTo Failed
    m_HasChanged = True
    CopyToList MSGListView2, MSGListView1
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
' Add's the selection(s) from the left listbox to the right listbox.
Private Sub cmdAddSecond_Click()
    On Error GoTo Failed
    m_HasChanged = True
    CopyToList MSGListView1, MSGListView2
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub cmdExtra_Click(Index As Integer)
    RaiseEvent ExtraClick(Index)
End Sub
Private Sub cmdMoveUp_Click()
    On Error GoTo Failed
    
    HandleMove MoveDirectionUp
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub cmdMoveDown_Click()
    On Error GoTo Failed
    
    HandleMove MoveDirectionDown
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub HandleMove(enumDirection As MoveDirection)
    On Error GoTo Failed
    Dim nCurrentPos As Long
    Dim lstItems As ListItems
    Dim lstSelectedItem As ListItem
    Dim lstNewPos As ListItem
    
    Set lstItems = MSGListView2.ListItems
    Set lstSelectedItem = MSGListView2.SelectedItem
    nCurrentPos = lstSelectedItem.Index
    
    If enumDirection = MoveDirectionUp Then
        ' Need to move down one
        Set lstNewPos = lstItems.Item(nCurrentPos - 1)
        SwapItems lstNewPos, lstSelectedItem
    Else
        Set lstNewPos = lstItems.Item(nCurrentPos + 1)
        SwapItems lstSelectedItem, lstNewPos
    End If
    MSGListView2.HideSelection = False
    Set MSGListView2.SelectedItem = lstNewPos
    'MSGListView2.SetFocus
    
    HandleMoveUpDown lstNewPos
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub SwapItems(lstItemOne As ListItem, lstItemTwo As ListItem)
    On Error GoTo Failed
    Dim sVal As String
    Dim sTmpVal As String
    Dim objTmpRef As Object
    
    sVal = lstItemOne.Text
    sTmpVal = lstItemTwo.Text

    lstItemTwo.Text = sVal
    lstItemOne.Text = sTmpVal
    
    Set objTmpRef = lstItemOne.Tag
    Set lstItemOne.Tag = lstItemTwo.Tag
    Set lstItemTwo.Tag = objTmpRef

    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub MSGListView1_Click()
    On Error GoTo Failed
    Dim Item As ListItem
    Set Item = MSGListView1.SelectedItem
    
    If (Not Item Is Nothing) And m_bSwapEnabled Then
        cmdAddSecond.Enabled = True
        cmdAddFirst.Enabled = False
    Else
        cmdAddSecond.Enabled = False
        cmdAddFirst.Enabled = False
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub MSGListView2_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim Item As ListItem
    Dim colLine As New Collection
    Dim objExtra As Object
    
    Set Item = MSGListView2.SelectedItem
    
    If Not Item Is Nothing And m_bSwapEnabled Then
        cmdAddFirst.Enabled = True
        cmdAddSecond.Enabled = False
    
        ' DJP Phase 2 Task Management
        HandleMoveUpDown Item
    
        ' Raise an event to say an item has been selected
        Set colLine = MSGListView2.GetLine(-1, objExtra)
        
        RaiseEvent SecondItemClicked(colLine, objExtra)
    Else
        cmdAddFirst.Enabled = False
        cmdAddSecond.Enabled = False
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub HandleMoveUpDown(itemSelected As ListItem)
    On Error GoTo Failed
    Dim nPos As Long
    Dim nCount As Long
    Dim lstItems As ListItems
    Dim bEnableUp As Boolean
    Dim bEnableDown As Boolean
    
    ' Need to figure out where we are in the list, and enable/disable accordingly
    Set lstItems = MSGListView2.ListItems
    nCount = lstItems.Count
    nPos = itemSelected.Index
    
    bEnableUp = False
    bEnableDown = False
    
    If nCount > 1 Then
        If nPos < nCount Then
            If nPos = 1 Then
                bEnableDown = True
                bEnableUp = False
            Else
                ' Enable Up button
                bEnableUp = True
                bEnableDown = True
            End If
        Else
            bEnableUp = True
            bEnableDown = False
        End If
    End If
    
    cmdMoveUp.Enabled = bEnableUp
    cmdMoveDown.Enabled = bEnableDown
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub UserControl_Initialize()
    On Error GoTo Failed
    
    Set m_clsErrorHandling = New ErrorHandling
    UserControl.Width = MSGListView2.Left + MSGListView2.Width + 150
    ' DJP SYS2831
    SetControlsPosition
    m_bSwapEnabled = True
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError "SWAPList"
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Failed
    Dim nMaxWidth As Integer
    Dim nMaxHeight As Integer
    Dim nButtonHeight As Integer
    Dim nButtonGap As Integer
    Dim nTop As Integer
    
    nMaxHeight = UserControl.Height / 100 * 75
    nButtonHeight = nMaxHeight / 100 * 17
    nButtonGap = nButtonHeight / 100 * 20
    
    MSGListView1.Height = nMaxHeight
    MSGListView2.Height = nMaxHeight
    
    cmdAddSecond.Height = nButtonHeight
    nTop = cmdAddSecond.Height + cmdAddSecond.Top + nButtonGap
    
    cmdAddAllToSecond.Top = nTop
    cmdAddAllToSecond.Height = nButtonHeight
    nTop = nTop + nButtonHeight + nButtonGap
        
    cmdAddAllToFirst.Top = nTop
    cmdAddAllToFirst.Height = nButtonHeight
    nTop = nTop + nButtonHeight + nButtonGap
    
    cmdAddFirst.Top = nTop
    cmdAddFirst.Height = nButtonHeight
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Public Function GetFirstCount() As Integer
    On Error GoTo Failed
    Dim nCount As Integer
    
    nCount = MSGListView1.ListItems.Count
    GetFirstCount = nCount
    Exit Function
Failed:
    m_clsErrorHandling.DisplayError
End Function
Public Function GetSecondCount() As Integer
    On Error GoTo Failed
    Dim nCount As Integer
    
    nCount = MSGListView2.ListItems.Count
    GetSecondCount = nCount
    Exit Function
Failed:
    m_clsErrorHandling.DisplayError
End Function
Public Function GetLineFirst(nIndex As Integer, Optional objExtra As Object = Nothing) As Collection
    Set GetLineFirst = MSGListView1.GetLine(nIndex, objExtra)
End Function
Public Function GetLineSecond(nIndex As Integer, Optional objExtra As Object = Nothing) As Collection
    Set GetLineSecond = MSGListView2.GetLine(nIndex, objExtra)
End Function
Public Sub SetSecondTitle(sTitle As String)
    lblListTitle(RIGHT_TITLE).Caption = sTitle
End Sub

Public Function SetFirstColumnHeaders(Headers As Collection) As Boolean
    Dim bRet As Boolean

    bRet = SetHeaders(MSGListView1, Headers)

End Function

Public Function SetSecondColumnHeaders(lHeaders As Collection) As Boolean
    Dim bRet As Boolean
    bRet = SetHeaders(MSGListView2, lHeaders)
End Function
Private Function SetHeaders(lvListToSet As MSGListView, lHeaders As Collection) As Boolean
    lvListToSet.AddHeadings lHeaders
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SetReportType(nReport As Integer) As Variant
    UserControl.MSGListView1.View = nReport
    UserControl.MSGListView2.View = nReport
End Function
Public Function AddLineFirst(line As Collection, Optional objExtra As Object = Nothing) As Boolean
    UserControl.MSGListView1.AddLine line, objExtra
    SetButtonsState
End Function
Public Function AddLineSecond(line As Collection, Optional objExtra As Object = Nothing) As Boolean
    UserControl.MSGListView2.AddLine line, objExtra
    cmdAddAllToFirst.Enabled = True
    SetButtonsState
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true0
Public Property Get MoveToSecond() As Boolean
    MoveToSecond = m_MoveToSecond
End Property

Public Property Let MoveToSecond(ByVal New_MoveToSecond As Boolean)
    m_MoveToSecond = New_MoveToSecond
    PropertyChanged "MoveToSecond"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MoveToSecond = m_def_MoveToSecond
    m_HasChanged = m_def_HasChanged
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MoveToSecond = PropBag.ReadProperty("MoveToSecond", m_def_MoveToSecond)
    Set m_Extra = PropBag.ReadProperty("Extra", Nothing)
    m_HasChanged = PropBag.ReadProperty("HasChanged", m_def_HasChanged)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MoveToSecond", m_MoveToSecond, m_def_MoveToSecond)
    Call PropBag.WriteProperty("Extra", m_Extra, Nothing)
    Call PropBag.WriteProperty("HasChanged", m_HasChanged, m_def_HasChanged)
End Sub
Friend Sub SetButtonsState()
    On Error GoTo Failed
    
    If m_bSwapEnabled Then
        If MSGListView1.ListItems.Count = 0 Then
            cmdAddAllToSecond.Enabled = False
            cmdAddSecond.Enabled = False
        Else
            cmdAddAllToSecond.Enabled = True
            cmdAddSecond.Enabled = False
        End If
    
        If MSGListView2.ListItems.Count = 0 Then
            cmdAddAllToFirst.Enabled = False
            cmdAddFirst.Enabled = False
        Else
            cmdAddAllToFirst.Enabled = True
            cmdAddFirst.Enabled = False
        End If
    Else
        cmdAddFirst.Enabled = False
        cmdAddSecond.Enabled = False
        cmdAddAllToFirst.Enabled = False
        cmdAddAllToSecond.Enabled = False
    End If
    
    cmdMoveUp.Enabled = False
    cmdMoveDown.Enabled = False
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub CopyAllToList(lstSource As MSGListView, lstDest As MSGListView)
    On Error GoTo Failed
    Dim Item As ListItem
    Dim newitem As ListItem
    
    On Error GoTo Failed
    
    For Each Item In lstSource.ListItems
        AddItem Item, lstDest
    Next

    lstSource.ListItems.Clear
    SetButtonsState

    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub AddItem(srcItem As ListItem, lstDest As MSGListView)
    Dim newitem As ListItem
    On Error GoTo Failed
    
    If (Not srcItem Is Nothing) Then
        Dim nSubItemCount
        Dim subItem As ListSubItem
        Dim srcSubItem As ListSubItem
        
        Set newitem = lstDest.ListItems.Add
        MSGListView1.ListItems
        newitem = srcItem
        If IsObject(srcItem.Tag) Then
            Set newitem.Tag = srcItem.Tag
        End If
        
        For Each srcSubItem In srcItem.ListSubItems
            Set subItem = newitem.ListSubItems.Add()
            subItem = srcSubItem
        Next
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub CopyToList(lstSource As MSGListView, lstDest As MSGListView)
    Dim Item As ListItem
    Dim bRet As Boolean
    Dim newitem As ListItem
    On Error GoTo Failed
    Set Item = lstSource.SelectedItem
    
    If (Not Item Is Nothing) Then
        AddItem Item, lstDest
        lstSource.RemoveLine Item.Index
        SetButtonsState
    End If
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
Private Sub SetControlsPosition()
    SetControlHorizontal
End Sub

Private Sub SetControlHorizontal()
    Set cmdAddSecond.Picture = imlButtons.ListImages("AddOneSecondH").Picture
    Set cmdAddAllToSecond.Picture = imlButtons.ListImages("AddAllSecondH").Picture
    Set cmdAddFirst.Picture = imlButtons.ListImages("AddOneFirstH").Picture
    Set cmdAddAllToFirst.Picture = imlButtons.ListImages("AddAllFirstH").Picture
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,2,0
Public Property Get Extra() As Object
Attribute Extra.VB_MemberFlags = "400"
    Set Extra = m_Extra
End Property

Public Property Set Extra(ByVal New_Extra As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    Set m_Extra = New_Extra
    PropertyChanged "Extra"

End Property

Public Function AddExtraButton(sCaption As String, nTop As Integer, nLeft As Integer) As Integer
    Dim nCount As Integer
    On Error GoTo Failed
    nCount = cmdExtra.Count()
    
    ' A new one.
    Load cmdExtra(nCount)
    cmdExtra(nCount).Visible = True
    cmdExtra(nCount).Enabled = True
    AddExtraButton = nCount
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

Public Sub ShowExtraButton(nIndex As Integer, Optional bShow As Boolean = True)
    On Error GoTo Failed
    cmdExtra(nIndex).Visible = bShow
    Exit Sub
Failed:
    MsgBox ("Unable to Enable/disable control (" + CStr(nIndex) + "). Error is " + Err.Description)
End Sub

Public Sub EnableExtraButton(nIndex As Integer, Optional bEnable As Boolean = True)
    On Error GoTo Failed
    cmdExtra(nIndex).Enabled = bEnable
    Exit Sub
Failed:
    MsgBox ("Unable to Enable/disable control (" + CStr(nIndex) + "). Error is " + Err.Description)
End Sub

Public Sub SetExtraCaption(nIndex As Integer, sCaption As String)
    On Error GoTo Failed
    cmdExtra(nIndex).Caption = sCaption
    
    Exit Sub
Failed:
    MsgBox ("Unable to set control (" + CStr(nIndex) + ") caption. Error is: " + Err.Description)
End Sub
Public Sub SetExtraPosition(nIndex As Integer, nTop As Integer, nLeft As Integer)
    With cmdExtra(nIndex)
        .Top = nTop
        .Left = nLeft
        .Width = cmdExtra(0).Width
        .Height = cmdExtra(0).Height
    End With
    Exit Sub
Failed:
    MsgBox ("Unable to set control (" + CStr(nIndex) + ") position. Error is: " + Err.Description)
End Sub
Public Sub ClearFirst()
    MSGListView1.ListItems.Clear
End Sub
Public Sub ClearSecond()
    MSGListView2.ListItems.Clear
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoesValueExist
' Description   : Returns TRUE if sValue is contained in the (second) swaplist, FALSE if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DoesValueExist(sValue As String) As Boolean
    On Error GoTo Failed
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sSecondValue As String
    Dim colValue As Collection
    
    nCount = GetSecondCount()
    
    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        
        Set colValue = GetLineSecond(nThisItem)
        
        If colValue.Count > 0 Then
            sSecondValue = colValue(1)
        
            If sSecondValue = sValue Then
                bFound = True
            End If
        End If
        
        nThisItem = nThisItem + 1
    Wend
    
    DoesValueExist = bFound
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,0
Public Property Get HasChanged() As Boolean
Attribute HasChanged.VB_MemberFlags = "400"
    HasChanged = m_HasChanged
End Property
Public Property Let HasChanged(ByVal New_HasChanged As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_HasChanged = New_HasChanged
    PropertyChanged "HasChanged"
End Property
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableButtons
' Description   : Enables or disables the state of the four swap buttons, depending on bEnable.
'                 Defaults to Enable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnableButtons(Optional bEnable As Boolean = True)
    On Error GoTo Failed

    m_bSwapEnabled = bEnable
    SetButtonsState
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
