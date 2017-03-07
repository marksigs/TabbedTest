VERSION 5.00
Begin VB.UserControl MSGMulti 
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   ScaleHeight     =   1515
   ScaleWidth      =   2805
   Begin VB.Timer timerCheckLength 
      Enabled         =   0   'False
      Left            =   1140
      Top             =   720
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "MSGMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UserControl   : MSGTextMulti
' Description   : ActiveX control

' Change history
' Prog      Date        Description
' TW        28/11/2007  VR973 - Error in Supervisor when pasting a field which is too long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
'Const m_def_BorderStyle = 0
Const m_def_Mandatory = 0
Const m_def_MaxLength = 200
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
'Dim m_BorderStyle As Integer
Dim m_Mandatory As Boolean
Dim m_MaxLength As Integer

' TW 28/11/2007 VR973
Private m_clsErrorHandling As ErrorHandling
' TW 28/11/2007 VR973 End

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
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
Event Validate(Cancel As Boolean) 'MappingInfo=Text1,Text1,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."

Private Sub Text1_GotFocus()
    On Error GoTo Failed
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
Failed:
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
' TW 28/11/2007 VR973
    On Error GoTo Failed
    If Len(Text1.Text) > m_MaxLength Then
        Text1.Text = Left$(Text1.Text, m_MaxLength)
    End If
    Cancel = Not ValidateData()
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
' TW 28/11/2007 VR973 End
End Sub


Private Sub timerCheckLength_Timer()
    Dim sText As String
    sText = timerCheckLength.Tag
    
    If Len(Text1.Text) > m_MaxLength Then
' TW 28/11/2007 VR973
'        Text1.Text = sText
'        Text1.SelStart = Len(sText)
        Text1.Text = Left$(sText, m_MaxLength)
        Text1.SelStart = m_MaxLength
' TW 28/11/2007 VR973 End
    End If
    timerCheckLength.Enabled = False
End Sub


Private Sub UserControl_Resize()
    UserControl.Text1.Width = UserControl.Width
    UserControl.Text1.Height = UserControl.Height
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
    Enabled = m_Enabled
End Property

Public Property Set Enabled(ByVal New_Enabled As Variant)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    Text1.Enabled = New_Enabled
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
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BorderStyle() As Integer
'    BorderStyle = m_BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    m_BorderStyle = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = Text1.MultiLine
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'Private Sub Text1_Validate(Cancel As Boolean)
'    RaiseEvent Validate(Cancel)
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ValidateControls
Public Sub ValidateControls()
Attribute ValidateControls.VB_Description = "Validate contents of the last control on the form before exiting the form"
    UserControl.ValidateControls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
'    m_BorderStyle = m_def_BorderStyle
    m_Mandatory = m_def_Mandatory
    m_MaxLength = m_def_MaxLength


End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
'    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Text1.Text = PropBag.ReadProperty("Text", "Text1")
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.SelText = PropBag.ReadProperty("SelText", "")
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    m_MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
    Text1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
'    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Text", Text1.Text, "Text1")
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Text1.SelText, "")
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
    Call PropBag.WriteProperty("BorderStyle", Text1.BorderStyle, 1)
End Sub
Public Sub SetFocus()
    Text1.SetFocus
End Sub

Public Function ValidateData(Optional bReportError As Boolean = True) As Boolean
    Dim bRet As Boolean
    
    bRet = CheckMandatory
    
    If bRet = False Then
        HighLight
    End If
    
    ValidateData = bRet

End Function
Public Sub HighLight(Optional bFocus As Boolean = False)
    On Error GoTo Error
    Dim nLen As String
    Dim sTmp As String
    
    
    sTmp = Text1.Text
    
    nLen = Len(sTmp)
    
    If (nLen > 0) Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
    
    If (bFocus) Then
        Text1.SetFocus
    End If
Error:
End Sub
Public Function CheckMandatory(Optional bHighLight As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim sVal As String
    
    bRet = True

    If m_Mandatory = True Then
        sVal = Text1.Text

        If Len(sVal) > 0 Then
            Text1.BackColor = vbWhite
            bRet = True
        Else
            bRet = False
            If bHighLight = True Then
                Text1.BackColor = vbRed '&H80FF&
            End If
        End If
    Else
        ' Incase the type changed from mandatory to not mandatory
        Text1.BackColor = vbWhite
    End If

    CheckMandatory = bRet
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
    Mandatory = m_Mandatory
End Property
Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    CheckMandatory False
End Property
Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Integer)
    m_MaxLength = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim sChar As String
    sChar = Chr(KeyAscii)
    timerCheckLength.Enabled = True
    timerCheckLength.Interval = 1
    timerCheckLength.Tag = Text1.Text
    'MsgBox "Len is " & Len(Text1.Text) '    If Text1.SelStart = 0 Then
'        If sChar = UCase(sChar) Then
'            ' Upper case, make lower
'            KeyAscii = Asc(LCase(sChar))
'        Else
'            ' Lower case,make upper
'            KeyAscii = Asc(UCase(sChar))
'        End If
'    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    Dim sChar As String
'    sChar = Chr(KeyAscii)
'
'    If InStr(1, g_sStandardCharSet, sChar) > 0 Then
'        If Text1.SelStart = 0 Then
'            If sChar = UCase(sChar) Then
'                ' Upper case, make lower
'                KeyAscii = Asc(LCase(sChar))
'            Else
'                ' Lower case,make upper
'                KeyAscii = Asc(UCase(sChar))
'            End If
'        End If
'    Else
'        KeyAscii = 0
'    End If
'End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Text1,Text1,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Text1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

