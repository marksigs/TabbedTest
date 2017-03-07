VERSION 5.00
Begin VB.UserControl MSGComboBox 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   330
   ScaleWidth      =   1635
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1635
   End
End
Attribute VB_Name = "MSGComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserControl   : MSGComboBox
' Description   : ActiveX control that does extra combo processing, including mandatory
'                 processing
' Change history
' Prog      Date        Description
' DJP       14/11/01    SYS2996 <None> becomes <Select>
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
' TW        28/11/2007  VR973 - Improved Mandatory checks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

 'Default Property Values:
Const m_def_Mandatory = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ListText As String = vbNullString   'BMIDS768 GHun
'Property Variables:
Dim m_ListText As String
Dim m_Mandatory As Boolean
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer

' Private data
Private m_colExtra As Collection
Private m_clsErrorHanding As ErrorHandling
'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:

'Default Property Values:
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = combo1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    combo1.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    Cancel = Not ValidateData()
End Sub

Private Sub UserControl_Initialize()
    Set m_clsErrorHanding = New ErrorHandling
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Mandatory = m_def_Mandatory
    m_ListText = m_def_ListText
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    On Error Resume Next
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ListText = PropBag.ReadProperty("ListText", m_def_ListText)
    SetListText m_ListText
    combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    combo1.SelLength = PropBag.ReadProperty("SelLength", 0)
    
    combo1.SelStart = PropBag.ReadProperty("SelStart", 0)
    combo1.SelText = PropBag.ReadProperty("SelText", "")
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    combo1.DataMember = PropBag.ReadProperty("DataMember", "")
    'Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)   'BMIDS768 GHun
    combo1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

    combo1.Text = PropBag.ReadProperty("Text", "Combo1")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    combo1.ItemData(Index) = PropBag.ReadProperty("ItemData" & Index, 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    On Error Resume Next
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ListIndex", combo1.ListIndex, 0)
    Call PropBag.WriteProperty("SelLength", combo1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", combo1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", combo1.SelText, "")
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataMember", combo1.DataMember, "")
    'Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)  'BMIDS768 GHun
    Call PropBag.WriteProperty("ToolTipText", combo1.ToolTipText, "")
    Call PropBag.WriteProperty("ListText", m_ListText, m_def_ListText)
    Call PropBag.WriteProperty("Text", combo1.Text, "Combo1")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("ItemData" & Index, combo1.ItemData(Index), 0)
End Sub

Private Sub UserControl_Resize()
    UserControl.combo1.Width = UserControl.Width
    UserControl.Height = UserControl.combo1.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = combo1.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    combo1.ListIndex = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    On Error Resume Next
    SelLength = combo1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    On Error Resume Next
    combo1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    On Error Resume Next
    SelStart = combo1.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    On Error Resume Next
    combo1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    On Error Resume Next
    Dim sText As String
    
    sText = combo1.Text
    
    If sText <> COMBO_NONE Then
        SelText = sText
    End If
End Property
Public Property Let SelText(ByVal New_SelText As String)
    On Error Resume Next
    combo1.SelText = New_SelText
    PropertyChanged "SelText"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property
Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    
    If New_Mandatory = False Then
        ' Clear the current selection, and change the colour to white (incase it's red)
        combo1.BackColor = vbWhite
        combo1.ListIndex = -1
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Style
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
    Style = combo1.Style
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,DataSource
Public Property Get DataSource() As adodb.Recordset
Attribute DataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set DataSource = combo1.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As adodb.Recordset)
    Set combo1.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    DataMember = combo1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    combo1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23
Public Sub DataFieldCombo(sField As String)
    combo1.DataField = sField
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    combo1.AddItem Item, Index
End Sub
Public Function ValidateData(Optional bReportError As Boolean = True) As Boolean
    Dim bRet As Boolean
' TW 28/11/2007 VR973
'    bRet = True
'    CheckMandatory
    bRet = CheckMandatory
' TW 28/11/2007 VR973 End

    ValidateData = bRet
End Function
' Checks to see if the combo has had anything selected in it. If not, highlight
' in red
Public Function CheckMandatory(Optional bHighLight As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim sVal As String
    
    bRet = True
    
    If m_Mandatory = True And combo1.Enabled = True Then
        If Len(Text()) > 0 Then
            combo1.BackColor = vbWhite
            bRet = True
        Else
            bRet = False
            If bHighLight = True Then
                combo1.BackColor = vbRed '&H80FF&
            End If
        End If
    End If

    CheckMandatory = bRet
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = combo1.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    combo1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,List
Public Property Get List(ByVal Index As Integer) As Variant
    List = combo1.List(Index)
End Property
Public Property Let List(ByVal Index As Integer, ByVal New_List As Variant)
    combo1.List(Index) = New_List
    PropertyChanged "List"
End Property
' Takes a collection and populates the combobox. Optionally takes a collection
' of strings which can be added to each item in the combo
Public Sub SetListTextFromCollection(colValues As Collection, Optional colIDS As Collection = Nothing)
    Dim nCount As Integer
    Dim nThisValue As Integer
    Dim sValue As String
    
    nCount = colValues.Count
    combo1.Clear
    Set m_colExtra = New Collection
    
    If nCount > 0 Then
        For nThisValue = 1 To nCount
            sValue = colValues(nThisValue)
            combo1.AddItem sValue
            
            If Not colIDS Is Nothing Then
                If colIDS.Count > 0 Then
                    m_colExtra.Add colIDS(nThisValue), CStr(nThisValue)
                End If
            End If
        Next
    End If
End Sub
' Returns the extra string associated with the index passed in, if there
' is an extra value
Public Function GetExtra(nIndex As Integer) As String
    On Error GoTo Failed
    
    If Not m_colExtra Is Nothing Then
        If nIndex + 1 <= m_colExtra.Count Then
            GetExtra = m_colExtra(CStr(nIndex + 1))
        End If
    Else
        m_clsErrorHanding.RaiseError errGeneralError, "MSGComboBox.GetExtra - no extra values"
    End If
    
    Exit Function
Failed:
    m_clsErrorHanding.RaiseError Err.Number, Err.Description
End Function
' Sets the text in the combo using a | delimited string. It's not possible to
' have a drop down combo at design time where we can enter them, so each item
' is separated by a | in one string
Private Function SetListText(sStr As String) As String
    Dim nPos As Variant
    Dim nStart As Variant
    Dim nLen As Long
    Dim nEnd As Variant
    Dim bDone As Boolean
    Dim sLine As String
    
    nLen = Len(sStr)
    bDone = False
    nStart = 1
    combo1.Clear
    
    If Len(sStr) > 0 Then
        While bDone = False
            nEnd = InStr(nStart, sStr, "|")
        
            If Not IsNull(nStart) Then
                If nEnd > 0 Then
                    sLine = Mid(sStr, nStart, nEnd - nStart)
                    nStart = nEnd + 1
                Else
                    sLine = Right(sStr, nLen - (nStart - 1))
                    bDone = True
                End If
            Else
                sLine = sStr
                bDone = True
            End If
    
            combo1.AddItem sLine
        Wend
    End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ListText() As String
Attribute ListText.VB_Description = "Pipe separated list of strings for the combobox."
    ListText = m_ListText
End Property
Public Property Let ListText(ByVal New_ListText As String)
    m_ListText = New_ListText
    SetListText New_ListText
    PropertyChanged "ListText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
    On Error Resume Next
    Dim sText As String
    
    sText = combo1.Text
    
    If sText <> COMBO_NONE Then
        Text = sText
    End If

End Property

Public Property Let Text(ByVal New_Text As String)
    combo1.Text() = New_Text
    PropertyChanged "Text"
End Property

Private Sub Combo1_Click()
    RaiseEvent Click
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ItemData
Public Property Get ItemData(ByVal Index As Integer) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a ComboBox or ListBox control."
    ItemData = combo1.ItemData(Index)
End Property
Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
    combo1.ItemData(Index) = New_ItemData
    PropertyChanged "ItemData"
End Property
Public Sub SetComboFocus()
    combo1.SetFocus
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

