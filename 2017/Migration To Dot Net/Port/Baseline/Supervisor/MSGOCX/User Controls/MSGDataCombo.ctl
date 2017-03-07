VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.UserControl MSGDataCombo 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   ScaleHeight     =   300
   ScaleWidth      =   1230
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
   End
End
Attribute VB_Name = "MSGDataCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserControl   : MSGDataCombo
' Description   : ActiveX combo control that's hooked up to the database, that inclues extra
'                 processing such as mandatory processing

' Change history
' Prog      Date        Description
' DJP       14/11/01    SYS2996 <None> becomes <Select>
' TW        28/11/2007  VR973 - Improved Mandatory checks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'Default Property Values:
Const m_def_SelText = ""
Const m_def_ComboDatafield1 = "0"
Const m_def_Mandatory = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Style = 0
Const m_def_Text = ""
'Property Variables:
Dim m_SelText As String
Dim m_ComboDatafield1 As String
Dim m_Mandatory As Boolean
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_DataMembers As DataMembers
Dim m_Style As StyleConstants
Dim m_Text As String

'Event Declarations:
Event Click(Area As Integer) 'MappingInfo=DataCombo1,DataCombo1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
'Event Click()
Event Change()
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Sub DoStuff()

End Sub
'
'
'Private Sub DataCombo1_Click(Area As Integer)
'    RaiseEvent Click
'End Sub

Private Sub DataCombo1_Change()
    RaiseEvent Change
End Sub

Private Sub UserControl_Resize()
    DataCombo1.Width = UserControl.Width
    DataCombo1.Height = UserControl.Height
End Sub
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
    Enabled = DataCombo1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    DataCombo1.Enabled = New_Enabled
    Mandatory = False
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
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,BoundColumn
Public Property Get BoundColumn() As String
Attribute BoundColumn.VB_Description = "Returns/sets the name of the source field in a Recordset object used to supply a data value to another control."
    BoundColumn = DataCombo1.BoundColumn
End Property
Public Property Let BoundColumn(ByVal New_BoundColumn As String)
    DataCombo1.BoundColumn() = New_BoundColumn
    PropertyChanged "BoundColumn"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,BoundText
Public Property Get BoundText() As String
Attribute BoundText.VB_Description = "Returns/sets the value of the data field named in the BoundColumn property."
    BoundText = DataCombo1.BoundText
End Property
Public Property Let BoundText(ByVal New_BoundText As String)
    DataCombo1.BoundText() = New_BoundText
    PropertyChanged "BoundText"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,DataSource
Public Property Get DataSource() As adodb.Recordset
Attribute DataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set DataSource = DataCombo1.DataSource
End Property
Public Property Set DataSource(ByVal New_DataSource As adodb.Recordset)
    Set DataCombo1.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=16,0,0,0
Public Property Get DataMembers() As DataMembers
Attribute DataMembers.VB_Description = "Returns a collection of data members to show at design time for this data source."
    Set DataMembers = m_DataMembers
End Property
Public Property Set DataMembers(ByVal New_DataMembers As DataMembers)
    Set m_DataMembers = New_DataMembers
    PropertyChanged "DataMembers"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    DataMember = DataCombo1.DataMember
End Property
Public Property Let DataMember(ByVal New_DataMember As String)
    DataCombo1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,ListField
Public Property Get ListField() As String
Attribute ListField.VB_Description = "Returns/sets the name of the field in the Recordset object used to fill a control's list portion."
    ListField = DataCombo1.ListField
End Property
Public Property Let ListField(ByVal New_ListField As String)
    DataCombo1.ListField() = New_ListField
    PropertyChanged "ListField"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,RowSource
Public Property Get RowSource() As adodb.Recordset
Attribute RowSource.VB_Description = "Returns/Sets data source for list items"
    Set RowSource = DataCombo1.RowSource
End Property
Public Property Set RowSource(ByVal New_RowSource As adodb.Recordset)
    Set DataCombo1.RowSource = New_RowSource
    PropertyChanged "RowSource"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get Style() As StyleConstants
Attribute Style.VB_Description = "Returns/sets a value that determines the type of control and the behavior of its list box portion."
    Style = m_Style
End Property
Public Property Let Style(ByVal New_Style As StyleConstants)
    m_Style = New_Style
    PropertyChanged "Style"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=DataCombo1,DataCombo1,-1,SelectedItem
Public Property Get SelectedItem() As Variant
Attribute SelectedItem.VB_Description = "Returns a value containing a bookmark for the selected record in a control."
    Dim sValue As Variant
    sValue = DataCombo1.SelectedItem
    
    If Not IsEmpty(sValue) And Not IsNull(sValue) Then
        If sValue <> COMBO_NONE Then
            SelectedItem = DataCombo1.SelectedItem
        End If
    Else
        SelectedItem = ""
    End If
    
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = m_Text
End Property
Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    DataCombo1.Text = New_Text
    PropertyChanged "Text"
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error GoTo Failed
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Style = m_def_Style
    m_Text = m_def_Text
    m_Mandatory = m_def_Mandatory
    m_ComboDatafield1 = m_def_ComboDatafield1
    m_SelText = m_def_SelText
    Exit Sub
Failed:
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo Failed
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    DataCombo1.BoundColumn = PropBag.ReadProperty("BoundColumn", "")
    DataCombo1.BoundText = PropBag.ReadProperty("BoundText", "")
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set m_DataMembers = PropBag.ReadProperty("DataMembers", Nothing)
    DataCombo1.DataMember = PropBag.ReadProperty("DataMember", "")
    DataCombo1.ListField = PropBag.ReadProperty("ListField", "")
    Set RowSource = PropBag.ReadProperty("RowSource", Nothing)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    m_ComboDatafield1 = PropBag.ReadProperty("ComboDatafield1", m_def_ComboDatafield1)
    m_SelText = PropBag.ReadProperty("SelText", m_def_SelText)
Failed:
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BoundColumn", DataCombo1.BoundColumn, "")
    Call PropBag.WriteProperty("BoundText", DataCombo1.BoundText, "")
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataMembers", m_DataMembers, Nothing)
    Call PropBag.WriteProperty("DataMember", DataCombo1.DataMember, "")
    Call PropBag.WriteProperty("ListField", DataCombo1.ListField, "")
    Call PropBag.WriteProperty("RowSource", RowSource, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("ComboDatafield1", m_ComboDatafield1, m_def_ComboDatafield1)
    Call PropBag.WriteProperty("SelText", m_SelText, m_def_SelText)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property
Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    DataCombo1.BackColor = vbWhite
    PropertyChanged "Mandatory"
End Property
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14
'Public Sub ComboDataField(sField As String)
'    DataCombo1.DataField = sField
'End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ComboDataField(sField As String) As Variant

End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ComboDatafield1() As String
    ComboDatafield1 = m_ComboDatafield1
End Property
Public Property Let ComboDatafield1(ByVal New_ComboDatafield1 As String)
    m_ComboDatafield1 = New_ComboDatafield1
    PropertyChanged "ComboDatafield1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,0
Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    Dim sText As String
    
    sText = DataCombo1.Text
    
    If sText <> COMBO_NONE Then
        SelText = sText
    End If
    
End Property

Public Property Let SelText(ByVal New_SelText As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_SelText = New_SelText
    DataCombo1.Text = New_SelText
    PropertyChanged "SelText"
End Property
' Sets the colour of the control to red if no item has been selected from
' the combo to indicate the control must have an item selected.
Public Function CheckMandatory(Optional bReportError As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim vVal As Variant
        
    bRet = True
    
    If m_Mandatory = True And DataCombo1.Enabled = True Then
        'vVal = DataCombo1.SelectedItem
        vVal = SelectedItem

        If Len(vVal) > 0 Then
            DataCombo1.BackColor = vbWhite
            bRet = True
        Else
            bRet = False
        End If

    End If

    If bRet = False Then
        DataCombo1.BackColor = vbRed
    End If
    
    CheckMandatory = bRet
End Function
' Validate doesn't work on ActiveX controls properly, so have to use lost_focus,
' and even then, only when the control that's active is allowed to cause validation
Private Sub DataCombo1_LostFocus()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    If UserControl.Parent.ActiveControl.CausesValidation = True Then
        bRet = CheckMandatory()
    End If
    
    If bRet = True Then
        ValidateControls
    End If
Failed:
End Sub
Public Sub SetFocus()
    If DataCombo1.Enabled = True Then
        DataCombo1.SetFocus
    End If
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

Private Sub DataCombo1_Click(Area As Integer)
    RaiseEvent Click(Area)
End Sub

