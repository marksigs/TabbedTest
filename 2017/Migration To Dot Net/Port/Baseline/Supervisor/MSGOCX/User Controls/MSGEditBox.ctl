VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl MSGEditBox 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   315
   ScaleWidth      =   1185
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "MSGEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmMain
' Description   :   ActiveX control to add functionality to the textedit
'                   control, such as mandatory processing, date processing etc.
' Change history
'
' Prog      Date        Description
' AA/DJP    08/12/00    Remove DebugInfo which was unused, and tidy up.
' AA        20/02/01    Add Time Control Type
' STB       05/12/01    SYS1664 - Increased number of allowed decimal places
'                       for numerics (percentages) to six.
' STB   08-May-2002 MSMS0069 The Mask property now works for string or numeric types.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'------------------------------------------------------------------------------
' BMids History
' Prog      Date        Description
' GHun      07/10/2004  BMIDS911 Modified CONTROL_NUMBER_NEGATIVE to allow negative values
' HMA       03/11/2004  BMIDS923 Remove Format for CONTROL_DATE and add extra validation.
'------------------------------------------------------------------------------
' Epsom History
' Prog      Date        Description
' TW        28/11/2007  VR973 - Improved Mandatory checks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Enums

Public Enum ControlTypes
    CONTROL_NUMBER = 0
    CONTROL_DATE = 1
    CONTROL_DOUBLE = 2
    CONTROL_BOOLEAN = 3
    CONTROL_STRING = 4
    CONTROL_NUMBER_NEGATIVE = 5
    CONTROL_SHORT = 6
    CONTROL_LONG = 7
    CONTROL_TIME = 8
End Enum

' Constants (number definitions as defined in Select Case Tool)
Private Const MAX_SHORT = 32767
Private Const MAX_NUMBER = 2147483647
Private Const MIN_NUMBER = -2147483647
Private Const MAX_DOUBLE = 999999999999#
Private Const MIN_DOUBLE = -999999999999#
Private Const DEFAULT_PLACES_BEFORE_POINT = 9
Private Const DEFAULT_PLACES_AFTER_POINT = 6

' Private data
'Private m_bReportError As Boolean
'Private m_sPassword As String
Private m_clsErrorHandling As ErrorHandling
Private m_AllowTab As Boolean
Private m_Mandatory As Boolean
Private m_MaxValue As Variant
Private m_MinValue As Variant
Private m_FirstCharUpper As Boolean
Private m_ValidateOnLoseFocus As Boolean
Private m_CheckValidate As Boolean
Private m_SelStart As Variant
Private m_TextLeft As Variant
Private m_TextRight As Variant
Private m_TextTop As Variant
Private m_TextHeight As Variant
Private m_TextType As ControlTypes
Private m_Text As String
Private m_EditDataField As String
Private m_DataMembers As DataMembers
Private m_SelLength As Variant

'Default Property Values:
Private Const m_def_TabStop = True
Private Const m_def_EditDataField = ""
Private Const m_def_SelLength = 0
Private Const m_def_AllowTab = False
Private Const m_def_Mandatory = 0
Private Const m_def_MaxValue = Null
Private Const m_def_MinValue = Null
Private Const m_def_FirstCharUpper = 1
Private Const m_def_ValidateOnLoseFocus = True
Private Const m_def_CheckValidate = True
Private Const m_def_SelStart = 0
Private Const m_def_TextLeft = 0
Private Const m_def_TextRight = 0
Private Const m_def_TextTop = 0
Private Const m_def_TextHeight = 0
Private Const m_def_TextType = 0
Private Const m_def_Text = ""

' Events
Public Event Change()
Public Event KeyPress(ByRef KeyAscii As Integer)
'Property Variables:
Private m_TabStop As Integer

Private Sub MaskEdBox1_Change()
    RaiseEvent Change
End Sub

Private Sub UserControl_Initialize()
    Set m_clsErrorHandling = New ErrorHandling
End Sub
' Highlights the text held in the control, if there is any
Public Sub HighLight(Optional ByVal bFocus As Boolean = False)
    On Error GoTo Error
    
    ' DJP, use len(cliptext) for highlighting
    If Len(MaskEdBox1.ClipText) > 0 Then
        MaskEdBox1.SelStart = 0
        MaskEdBox1.SelLength = Len(MaskEdBox1)
    End If
    
    If bFocus Then
        MaskEdBox1.SetFocus
    End If

Error:
End Sub
' public method that returns true of the data held in the control is valid, false if not. Also does
' mandatory checking.
Public Function ValidateData(Optional ByVal bReportError As Boolean = True) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    
    ' Do mandatory processing
' TW 28/11/2007 VR973
'    bRet = True
'    CheckMandatory
    bRet = CheckMandatory
    If bRet = True Then
' TW 28/11/2007 VR973 End
        ' Do the validation
        Select Case m_TextType
        Case CONTROL_DATE
            bRet = ValidateDate(bReportError)
        Case CONTROL_TIME
            bRet = ValidateTime(bReportError)
        'BMIDS911 GHun added CONTROL_NUMBER_NEGATIVE to case below
        Case CONTROL_SHORT, CONTROL_LONG, CONTROL_DOUBLE, CONTROL_NUMBER, CONTROL_NUMBER_NEGATIVE
            bRet = ValidateNumber()
        End Select
' TW 28/11/2007 VR973
    End If
' TW 28/11/2007 VR973 End
    If bRet = False Then
        ' Highlight the control if the data is invalid
        HighLight
    End If

    ValidateData = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
' Utility method that removes sChar from string sString, which is not modified, just returned
Private Function StripCharFromString(ByRef sString As String, ByRef sChar As String) As String
    On Error GoTo Failed
    'Dim nStart As Long
    Dim nLen As Long
    Dim nPos As Long
    Dim sThisChar As String
    Dim sTmp As String
    
    nLen = Len(sString)
    
    For nPos = 1 To nLen
        sThisChar = Mid$(sString, nPos, 1)
    
        If (sThisChar <> sChar) Then
            sTmp = sTmp + sThisChar
        End If
    Next nPos
    
    StripCharFromString = sTmp
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
' Called when the type of the textbox is changed - e.g., it may be changed from a date control to a control
' that holds a double.
Private Sub HandleNewType(ByVal nType As ControlTypes)
    On Error GoTo Failed
    ' Clear the mast, reset it if necessary
    Select Case nType
        Case CONTROL_NUMBER, CONTROL_NUMBER_NEGATIVE, CONTROL_SHORT, CONTROL_LONG
            MaskEdBox1.mask = ""
            MaskEdBox1.Format = ""
            PropertyChanged "Format"
            MaskEdBox1.Text = ""
        Case CONTROL_DATE
            
            SetDateMask
            
            'BMIDS923  Remove Format = "c" because this causes automatic conversion of an invalid date.
            '          eg. 12/13/2004 becomes 13/12/2004 instead of reporting an error.
            MaskEdBox1.Format = ""
            
            PropertyChanged "Format"
            MaskEdBox1.PromptChar = "_"
            MaskEdBox1.PromptInclude = True
        Case CONTROL_STRING
            MaskEdBox1.Format = ""
            MaskEdBox1.mask = ""
            MaskEdBox1.Text = ""
            MaskEdBox1.PromptInclude = False
        Case CONTROL_TIME
            MaskEdBox1.Format = ""
            MaskEdBox1.PromptChar = "_"
            PropertyChanged "Format"
            MaskEdBox1.mask = "##:##"
            MaskEdBox1.PromptInclude = True
            
    End Select
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
' Sets the mask to either contain two year digits or four
Private Sub SetDateMask()
    On Error GoTo Failed
    Dim nPos As Integer
    Dim nLastPos As Integer
    Dim nThisPos As Integer
    Dim sMask As String
    Dim sTestDate As String
    Dim sMaskChar As String
    
    sMaskChar = "#"
    
    sTestDate = CStr(#11/12/2000#)
    
    nLastPos = InStr(1, sTestDate, "/")
    
    For nThisPos = 1 To nLastPos - 1
        sMask = sMask + sMaskChar
    Next
    
    sMask = sMask + "/"
    
    nLastPos = nLastPos + 1
    nPos = InStr(nLastPos, sTestDate, "/")
    
    For nThisPos = nLastPos To nPos - 1
        sMask = sMask + sMaskChar
    Next

    sMask = sMask + "/"
    
    nPos = InStr(nLastPos, sTestDate, "/")
    
    For nThisPos = nPos + 1 To Len(sTestDate)
        sMask = sMask + sMaskChar
    Next
    
    MaskEdBox1.mask = sMask
    
    Exit Sub
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Sub
' Checks the validity of the variant passed in. Returns true of the variant is not null and not empty
Private Function ValidateVariant(ByVal vVal As Variant) As Boolean
    ValidateVariant = False
    
    If Not IsNull(vVal) And Not IsNull(vVal) Then
        If Len(CStr(vVal)) > 0 Then
            ValidateVariant = True
        End If
    End If
End Function

' Returns the minimum value the control can take - the user may have set up a minimum value, in which case,
' return that. If the user has not setup a min value, return the value based on the type of the control.
Private Function GetMin() As Double
    If ValidateVariant(m_MinValue) Then
        GetMin = CDbl(m_MinValue)
    Else
        GetMin = GetMinFromType
    End If
End Function

' Returns the maximum value the control can take - the user may have set up a maximum value, in which case,
' return that. If the user has not setup a max value, return the value based on the type of the control.
Private Function GetMax() As Double
    If ValidateVariant(m_MaxValue) Then
        GetMax = CDbl(m_MaxValue)
    Else
        GetMax = GetMaxFromType
    End If
End Function

' Returns the maximum value the contrl can take based on the type of the contrl
Private Function GetMaxFromType() As Double
    GetMaxFromType = MAX_NUMBER
    
    Select Case m_TextType
    Case CONTROL_SHORT
        GetMaxFromType = MAX_SHORT
    Case CONTROL_LONG
        GetMaxFromType = MAX_NUMBER
    Case CONTROL_DOUBLE
        GetMaxFromType = MAX_DOUBLE
    Case CONTROL_NUMBER
        GetMaxFromType = MAX_NUMBER
    'BMIDS911 GHun
    Case CONTROL_NUMBER_NEGATIVE
        GetMaxFromType = MAX_NUMBER
    'BMIDS911 End
    End Select
End Function

' Returns the minimum value the contrl can take based on the type of the contrl
Private Function GetMinFromType() As Double
    GetMinFromType = 0
    
    Select Case m_TextType
    Case CONTROL_SHORT
        GetMinFromType = 0
    Case CONTROL_LONG
        GetMinFromType = 0
    Case CONTROL_DOUBLE
        GetMinFromType = MIN_DOUBLE
    Case CONTROL_NUMBER
        GetMinFromType = 0
    'BMIDS911 GHun
    Case CONTROL_NUMBER_NEGATIVE
        GetMinFromType = MIN_NUMBER
    'BMIDS911 End
    End Select
End Function

' Validates the number the user has entered into the control
Friend Function ValidateNumber(Optional ByVal bReportError As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim dMin As Double
    Dim dMax As Double
    Dim sTmpStr As String
    Dim sError As String

    On Error GoTo Failed
    
    bRet = True
    sTmpStr = MaskEdBox1.Text
    
    ' Get the min and max values for the control
    dMin = GetMin()
    dMax = GetMax()
    
    If Len(sTmpStr) > 0 Then
        If IsNumeric(sTmpStr) Then
            If CDbl(sTmpStr) > dMax Then
                bRet = False
                sError = "Value entered cannot be more than " & dMax
            End If
        
            If bRet = True Then
                If CDbl(sTmpStr) < dMin Then
                    bRet = False
                    sError = "Value entered cannot be less than " & dMin
                End If
            End If
        Else
            bRet = False
            sError = "Please enter a valid number"
        End If
    End If
        
    If bRet = False And bReportError = True Then
        DisplayError sError
        HighLight
    End If
        
    ValidateNumber = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

' Validates the time the user has entered into the control
Public Function ValidateTime(Optional ByVal bReportError As Boolean = True) As Boolean
    
    On Error GoTo Failed
    Dim sTmpStr As String
    Dim bRet As Boolean
    Dim nLen As Long
    
    sTmpStr = MaskEdBox1.Text
    nLen = Len(sTmpStr) - 1
    
    If (nLen > 0) Then
        On Error Resume Next
        
        If IsTime(sTmpStr) Then
            bRet = True
        End If
        
        If (bRet = True) Then
'            If (MaskEdBox1.Text <> sTmpStr) Then
'                MaskEdBox1.Text = sTmpStr
'            End If
            
            If (InStr(1, MaskEdBox1.Text, "_") = 0) Then
                bRet = True
            Else
                bRet = False
            End If
        Else
            bRet = False
        End If
    End If
    
    If (bRet = True And nLen > 0) Then
        MaskEdBox1.BackColor = vbWhite
    End If
    
    If (bRet = False) Then
        If (bReportError = True) Then
            m_clsErrorHandling.DisplayError "Invalid Time Entered"
            HighLight
        End If
    
    End If
    ValidateTime = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

' Validates the date the user has entered into the control
Public Function ValidateDate(Optional ByVal bReportError As Boolean = True) As Boolean
    
    On Error GoTo Failed
    Dim sTmpStr As String
    Dim bRet As Boolean
    Dim nLen As Long
    
    bRet = True
    
    sTmpStr = GetFullYear(MaskEdBox1.Text)
    nLen = Len(sTmpStr) - 2  ' For the // in 12/12/1998
    
    If (nLen > 0) Then
        On Error Resume Next
        
        'BMIDS923  Use new function to perform extra validation
        If (Not Check_IsDate(sTmpStr)) Then
            bRet = False
        End If
        
        If (bRet = True) Then
            If (MaskEdBox1.Text <> sTmpStr) Then
                MaskEdBox1.Text = sTmpStr
            End If
            
            If (InStr(1, MaskEdBox1.Text, "_") = 0) Then
                bRet = True
            Else
                bRet = False
            End If
        Else
            bRet = False
        End If
    End If
    
    If (bRet = True And nLen > 0) Then
        SetDateNormal MaskEdBox1
    End If
    
    If (bRet = False) Then
        If (bReportError = True) Then
            m_clsErrorHandling.DisplayError "Invalid Date Entered"
            HighLight
        End If
    
    End If
    ValidateDate = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
Private Sub SetDateNormal(ByRef txtDate As MaskEdBox)
    txtDate.BackColor = vbWhite
End Sub

Public Property Let TextSelStart(ByVal nStart As Integer)
    MaskEdBox1.SelStart = nStart
End Property

Public Property Get TextSelStart() As Integer
    TextSelStart = MaskEdBox1.SelStart
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Mask
Public Property Get mask() As String
Attribute mask.VB_Description = "Determines the input mask for the control."
    mask = MaskEdBox1.mask
End Property

Public Property Let mask(ByVal New_Mask As String)
    MaskEdBox1.mask() = New_Mask
    PropertyChanged "Mask"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
    Text = MaskEdBox1.Text
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
    CausesValidation = MaskEdBox1.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
    MaskEdBox1.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get TextType() As ControlTypes
    TextType = m_TextType
End Property

Public Property Let TextType(ByVal New_TextType As ControlTypes)
    m_TextType = New_TextType
    PropertyChanged "TextType"
    HandleNewType New_TextType
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,PromptInclude
Public Property Get PromptInclude() As Boolean
Attribute PromptInclude.VB_Description = "Specifies whether prompt characters are contained in the Text property value."
    PromptInclude = MaskEdBox1.PromptInclude
End Property

Public Property Let PromptInclude(ByVal New_PromptInclude As Boolean)
    MaskEdBox1.PromptInclude() = New_PromptInclude
    PropertyChanged "PromptInclude"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,PromptChar
Public Property Get PromptChar() As String
Attribute PromptChar.VB_Description = "Sets/returns the character used to prompt a user for input."
    PromptChar = MaskEdBox1.PromptChar
End Property

Public Property Let PromptChar(ByVal New_PromptChar As String)
    MaskEdBox1.PromptChar() = New_PromptChar
    PropertyChanged "PromptChar"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SelStart() As Variant
    SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Variant)
    m_SelStart = New_SelStart
    MaskEdBox1.SelStart = m_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get CheckValidate() As Boolean
Attribute CheckValidate.VB_Description = "Specifies whether or not the control does validation when it loses focus"
    CheckValidate = m_CheckValidate
End Property

Public Property Let CheckValidate(ByVal New_CheckValidate As Boolean)
    m_CheckValidate = New_CheckValidate
    PropertyChanged "CheckValidate"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ValidateControls
'Public Sub ValidateControls()
'    UserControl.ValidateControls
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get ValidateOnLoseFocus() As Boolean
    ValidateOnLoseFocus = m_ValidateOnLoseFocus
End Property

Public Property Let ValidateOnLoseFocus(ByVal New_ValidateOnLoseFocus As Boolean)
    m_ValidateOnLoseFocus = New_ValidateOnLoseFocus
    PropertyChanged "ValidateOnLoseFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = MaskEdBox1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    MaskEdBox1.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Function ValidateKey(nKey As Integer) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sChar As String
    
    sChar = Chr$(nKey)
    bRet = False
    
    Select Case m_TextType
    'BMIDS911 GHun
    Case CONTROL_SHORT, CONTROL_LONG, CONTROL_NUMBER_NEGATIVE
        bRet = HandleNumberKey(nKey, False)
    Case CONTROL_DOUBLE, CONTROL_NUMBER
        bRet = HandleNumberKey(nKey)
    Case CONTROL_BOOLEAN
        If (sChar = "0" Or sChar = "1") And Len(MaskEdBox1.Text) = 0 Then
            bRet = True
        End If
    Case CONTROL_STRING
        sChar = Chr$(nKey)
        
        bRet = True

    Case Else
        bRet = True
    End Select

    ValidateKey = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Mandatory() As Boolean
Attribute Mandatory.VB_Description = "Defines whether the control is mandatory or not."
    Mandatory = m_Mandatory
End Property
Public Property Let Mandatory(ByVal New_Mandatory As Boolean)
    m_Mandatory = New_Mandatory
    PropertyChanged "Mandatory"
    CheckMandatory False
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MaxValue() As Variant
    MaxValue = m_MaxValue
End Property
Public Property Let MaxValue(ByVal New_MaxValue As Variant)
    m_MaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MinValue() As Variant
    MinValue = m_MinValue
End Property
Public Property Let MinValue(ByVal New_MinValue As Variant)
    m_MinValue = New_MinValue
    PropertyChanged "MinValue"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
' Validates the control to make sure the user has entered data into it if the control is set to
' mandatory - if no data has been entered, the control is highlighted in red
Public Function CheckMandatory(Optional ByVal bHighLight As Boolean = True) As Boolean
Attribute CheckMandatory.VB_Description = "If the control is mandatory, highlights the control in red and returns false."
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sVal As String
    
    bRet = True
    If m_Mandatory = True Then
        Select Case m_TextType
            
        Case CONTROL_DATE
            bRet = ValidateDate(False)
            If bRet = True Then
                sVal = MaskEdBox1.ClipText
            End If
        Case Else
            sVal = MaskEdBox1.Text
        End Select

        If Len(sVal) > 0 Then
            MaskEdBox1.BackColor = vbWhite
            bRet = True
        Else
            bRet = False
            If bHighLight = True Then
                MaskEdBox1.BackColor = vbRed
            End If
        End If
    Else
        ' Incase the type changed from mandatory to not mandatory
        MaskEdBox1.BackColor = vbWhite
    End If

    CheckMandatory = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get AllowTab() As Boolean
Attribute AllowTab.VB_Description = "Allows the control to accept the TAB key."
    AllowTab = m_AllowTab
End Property

Public Property Let AllowTab(ByVal New_AllowTab As Boolean)
    m_AllowTab = New_AllowTab
    PropertyChanged "AllowTab"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = MaskEdBox1.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    MaskEdBox1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
' When a key is pressed into the textbox, validate the key - make sure that the number of characters pressed
' is correct, before the point and after, if appropriate.
Public Function HandleNumberKey(KeyAscii As Integer, Optional ByVal bStopValid As Boolean = True) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nBeforePoint As Long
    Dim nAfterPoint As Long
    Dim sChar As String
    
    bRet = True
    sChar = Chr$(KeyAscii)

    If sChar = "." And bStopValid = False Then
        bRet = False
    End If
    
    If bRet Then
        bRet = IsValidControl(KeyAscii)
        
        If (bRet = False) Then
            bRet = IsDigit(sChar, MaskEdBox1.Text)
        
        
            If bRet = False Then
                Dim dMin As Double
                
                dMin = GetMin()
                
                If dMin < 0 And sChar = "-" And MaskEdBox1.SelStart = 0 Then
                    If InStr(1, MaskEdBox1.Text, "-", vbTextCompare) = 0 Then
                        bRet = True
                    End If
                End If
            End If
        End If
    
        If (bRet = True) Then
            Dim nSel As Long
            
            nSel = MaskEdBox1.SelLength
            
            If (nSel = 0) Then
                Dim sTmp As String
                Dim stmp1 As String
                Dim nStart As Long
                Dim nLen As Long
                
                nLen = Len(MaskEdBox1.Text)
                nStart = MaskEdBox1.SelStart
                sTmp = Left$(MaskEdBox1.Text, nStart)
                stmp1 = sTmp + Chr$(KeyAscii)
                sTmp = stmp1 + Right$(MaskEdBox1.Text, nLen - nStart)
                
                nBeforePoint = MaskEdBox1.MaxLength
                
                If nBeforePoint <= 0 Then
                    ' Default
                    nBeforePoint = DEFAULT_PLACES_BEFORE_POINT
                End If
                
                ' Default
                nAfterPoint = DEFAULT_PLACES_AFTER_POINT
                bRet = CheckDecimalPlaces(sTmp, nBeforePoint, nAfterPoint)
            Else
                bRet = True
            End If
        End If
    End If
    
    HandleNumberKey = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
' Validate the number in sText to make sure that the number of places before and after the decimal point
' are valid
Public Function CheckDecimalPlaces(ByVal sText As String, ByVal nBeforePoint As Long, ByVal nAfterPointValid As Long) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim nPos As Long
    Dim nAfterPoint As Long
    
    bRet = True
    
    ' After Point
    nPos = InStr(1, sText, ".", vbTextCompare)
    
    If (nPos >= 1) Then
        nAfterPoint = Len(sText) - nPos
        
        If (nAfterPoint > nAfterPointValid Or nAfterPointValid = 0) Then
            bRet = False
        End If
    End If

    ' Before Point
    
    If (bRet = True) Then
        If (nPos = 0) Then
          nPos = Len(sText)
         Else
            nPos = nPos - 1
        End If
          
        If (nPos > 0) Then
            'Dim sTmp As String
            'Dim nLen As Long
            
            'sTmp = Left$(sText, nPos - 1)
            
            If (nPos > nBeforePoint) Then
                bRet = False
            End If
        End If
    End If

    CheckDecimalPlaces = bRet
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get SelLength() As Variant
    SelLength = MaskEdBox1.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Variant)
    m_SelLength = New_SelLength
    MaskEdBox1.SelLength = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,DataSource
Public Property Get DataSource() As Object
Attribute DataSource.VB_Description = "Sets a value that specifies the Data control through which the current control is bound to a database. "
    Set DataSource = MaskEdBox1.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As Object)
    Set MaskEdBox1.DataSource = New_DataSource
    PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,DataFormat
Public Property Get DataFormat() As Object
Attribute DataFormat.VB_Description = "Returns a DataFormat object for use against a bindable property of this component."
    Set DataFormat = MaskEdBox1.DataFormat
End Property

Public Property Set DataFormat(ByVal New_DataFormat As Object)
    Set MaskEdBox1.DataFormat = New_DataFormat
    PropertyChanged "DataFormat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns/sets a value that describes the DataMember for a data connection."
    DataMember = MaskEdBox1.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
    MaskEdBox1.DataMember() = New_DataMember
    PropertyChanged "DataMember"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DataMemberChanged
Public Sub DataMemberChanged(ByVal DataMember As String)
Attribute DataMemberChanged.VB_Description = "Notify data consumers that a data member of this data source has changed."
    UserControl.DataMemberChanged DataMember
End Sub

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
'MemberInfo=13,0,2,
Public Property Get EditDataField() As String
Attribute EditDataField.VB_MemberFlags = "400"
    'EditDataField = m_EditDataField
    EditDataField = MaskEdBox1.DataField
End Property

Public Property Let EditDataField(ByVal New_EditDataField As String)
    If Ambient.UserMode = False Then
        Err.Raise 387
    End If
    m_EditDataField = New_EditDataField
    MaskEdBox1.DataField = New_EditDataField
    PropertyChanged "EditDataField"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,MaxLength
Public Property Get MaxLength() As Integer
Attribute MaxLength.VB_Description = "Sets/returns the maximum length of the masked edit control."
    MaxLength = MaskEdBox1.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Integer)
    MaskEdBox1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property
Public Property Get FirstCharUpper() As Boolean
    FirstCharUpper = m_FirstCharUpper
End Property
Public Property Let FirstCharUpper(ByVal New_Upper As Boolean)
    m_FirstCharUpper = New_Upper
    PropertyChanged "FirstCharUpper"
End Property
Public Sub SetTabStop(ByVal bStop As Boolean)
    MaskEdBox1.TabStop = bStop
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=MaskEdBox1,MaskEdBox1,0,TabStop
'Public Property Get TabStop() As Integer
'    TabStop = MaskEdBox1.TabStop
'End Property
'Public Property Let TabStop(ByVal New_TabStop As Integer)
'    MaskEdBox1.TabStop = New_TabStop
'    PropertyChanged "TabStop"
'End Property
Public Property Let Visible(ByVal New_Visible As Boolean)
    MaskEdBox1.Visible = New_Visible
    PropertyChanged "Visible"
End Property
Public Property Get Visible() As Boolean
    Visible = MaskEdBox1.Visible
End Property
Private Sub MaskEdBox1_GotFocus()
    ' To make sure the cursor is in position 1
    MaskEdBox1.Text = MaskEdBox1.Text
    HighLight
End Sub
Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
    On Error GoTo Failed
    Dim bRet As Boolean
    
    bRet = True
    
    If Not IsControl(KeyAscii) Then
        bRet = ValidateKey(KeyAscii)
    End If
    
    If bRet = False Then
        KeyAscii = 0
    Else
        ' DJP 2
        RaiseEvent KeyPress(KeyAscii)
    End If
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    On Error GoTo Failed
    
    Cancel = Not ValidateData()
    
    Exit Sub
Failed:
    m_clsErrorHandling.DisplayError
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Text = m_def_Text
    m_TextType = m_def_TextType
    m_SelStart = m_def_SelStart
    m_TextLeft = m_def_TextLeft
    m_TextRight = m_def_TextRight
    m_TextTop = m_def_TextTop
    m_TextHeight = m_def_TextHeight
    m_CheckValidate = m_def_CheckValidate
    m_ValidateOnLoseFocus = m_def_ValidateOnLoseFocus
    m_Mandatory = m_def_Mandatory
    m_MinValue = m_def_MinValue
    m_MaxValue = m_def_MaxValue
    m_FirstCharUpper = m_def_FirstCharUpper
    m_AllowTab = m_def_AllowTab
    m_SelLength = m_def_SelLength
    m_EditDataField = m_def_EditDataField
    m_TabStop = m_def_TabStop
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    MaskEdBox1.mask = PropBag.ReadProperty("Mask", "")
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    MaskEdBox1.Format = PropBag.ReadProperty("Format", "")
    PropertyChanged "Format"
    MaskEdBox1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
    m_TextType = PropBag.ReadProperty("TextType", m_def_TextType)
    
    MaskEdBox1.PromptInclude = PropBag.ReadProperty("PromptInclude", True)
    MaskEdBox1.PromptChar = PropBag.ReadProperty("PromptChar", "_")
    
    Dim nProp As Integer
    Dim sProp As String
    nProp = PropBag.ReadProperty("FontSize", 0)
    
    If (nProp > 0) Then
        UserControl.FontSize = nProp
    End If
    
    sProp = PropBag.ReadProperty("FontName", "")
    
    If (Len(sProp) > 0) Then
        UserControl.FontName = sProp
    End If
    m_SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
    m_TextLeft = PropBag.ReadProperty("TextLeft", m_def_TextLeft)
    m_TextRight = PropBag.ReadProperty("TextRight", m_def_TextRight)
    m_TextTop = PropBag.ReadProperty("TextTop", m_def_TextTop)
    m_TextHeight = PropBag.ReadProperty("TextHeight", m_def_TextHeight)
    m_CheckValidate = PropBag.ReadProperty("CheckValidate", m_def_CheckValidate)
    m_ValidateOnLoseFocus = PropBag.ReadProperty("ValidateOnLoseFocus", m_def_ValidateOnLoseFocus)
    MaskEdBox1.Enabled = PropBag.ReadProperty("Enabled", True)
    m_MinValue = PropBag.ReadProperty("MinValue", m_def_MinValue)
    m_MaxValue = PropBag.ReadProperty("MaxValue", m_def_MaxValue)
    m_Mandatory = PropBag.ReadProperty("Mandatory", m_def_Mandatory)
    m_FirstCharUpper = PropBag.ReadProperty("FirstCharUpper", m_def_FirstCharUpper)
    m_AllowTab = PropBag.ReadProperty("AllowTab", m_def_AllowTab)
    MaskEdBox1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    m_SelLength = PropBag.ReadProperty("SelLength", m_def_SelLength)
    Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
    Set DataFormat = PropBag.ReadProperty("DataFormat", Nothing)
    MaskEdBox1.DataMember = PropBag.ReadProperty("DataMember", "")
    Set m_DataMembers = PropBag.ReadProperty("DataMembers", Nothing)
    m_EditDataField = PropBag.ReadProperty("EditDataField", m_def_EditDataField)
    MaskEdBox1.MaxLength = PropBag.ReadProperty("MaxLength", 64)

    'MSMS0069 - If the format isn't a Date or time then firing this event will
    'clear our mask property if there is one set.
    If m_TextType = CONTROL_DATE Or m_TextType = CONTROL_TIME Then
        HandleNewType m_TextType
    End If
    
    MaskEdBox1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_TabStop = PropBag.ReadProperty("TabStop", m_def_TabStop)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Mask", MaskEdBox1.mask, "")
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("Format", MaskEdBox1.Format, "")
    Call PropBag.WriteProperty("CausesValidation", MaskEdBox1.CausesValidation, True)
    Call PropBag.WriteProperty("TextType", m_TextType, m_def_TextType)
    Call PropBag.WriteProperty("PromptInclude", MaskEdBox1.PromptInclude, True)
    Call PropBag.WriteProperty("PromptChar", MaskEdBox1.PromptChar, "_")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("SelStart", m_SelStart, m_def_SelStart)
    Call PropBag.WriteProperty("TextLeft", m_TextLeft, m_def_TextLeft)
    Call PropBag.WriteProperty("TextRight", m_TextRight, m_def_TextRight)
    Call PropBag.WriteProperty("TextTop", m_TextTop, m_def_TextTop)
    Call PropBag.WriteProperty("TextHeight", m_TextHeight, m_def_TextHeight)
    Call PropBag.WriteProperty("CheckValidate", m_CheckValidate, m_def_CheckValidate)
    Call PropBag.WriteProperty("ValidateOnLoseFocus", m_ValidateOnLoseFocus, m_def_ValidateOnLoseFocus)
    Call PropBag.WriteProperty("Enabled", MaskEdBox1.Enabled, True)
    Call PropBag.WriteProperty("FirstCharUpper", m_FirstCharUpper, m_def_FirstCharUpper)
    Call PropBag.WriteProperty("MinValue", m_MinValue, m_def_MinValue)
    Call PropBag.WriteProperty("MaxValue", m_MaxValue, m_def_MaxValue)
    Call PropBag.WriteProperty("Mandatory", m_Mandatory, m_def_Mandatory)
    Call PropBag.WriteProperty("AllowTab", m_AllowTab, m_def_AllowTab)
    Call PropBag.WriteProperty("BackColor", MaskEdBox1.BackColor, &H80000005)
    Call PropBag.WriteProperty("SelLength", m_SelLength, m_def_SelLength)
    Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
    Call PropBag.WriteProperty("DataFormat", DataFormat, Nothing)
    Call PropBag.WriteProperty("DataMember", MaskEdBox1.DataMember, "")
    Call PropBag.WriteProperty("DataMembers", m_DataMembers, Nothing)
    Call PropBag.WriteProperty("EditDataField", m_EditDataField, m_def_EditDataField)
    Call PropBag.WriteProperty("MaxLength", MaskEdBox1.MaxLength, 64)
    Call PropBag.WriteProperty("BorderStyle", MaskEdBox1.BorderStyle, 1)
    Call PropBag.WriteProperty("TabStop", m_TabStop, m_def_TabStop)
End Sub
' Resize the textbox when the activeX control is resized so it'll act
' like a normal textbox at design time.
Private Sub UserControl_Resize()
    UserControl.MaskEdBox1.Width = UserControl.Width
    UserControl.MaskEdBox1.Height = UserControl.Height
End Sub
Public Property Let Text(ByVal New_Text As String)
    On Error GoTo Failed
    Dim sMask As String
    
    sMask = MaskEdBox1.mask
    MaskEdBox1.mask = ""

    MaskEdBox1.Text = New_Text

    If Len(sMask) > 0 Then
        MaskEdBox1.mask = sMask
    End If
    PropertyChanged "Text"
Failed:
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,Format
Public Property Get Format() As String
    Format = MaskEdBox1.Format
End Property

Public Property Let Format(ByVal New_Format As String)
    MaskEdBox1.Format() = New_Format
    PropertyChanged "Format"
End Property
Private Sub DisplayError(ByVal sError As String)
    'MsgBox sError, vbCritical, "Validation Error"
End Sub
Public Function ClipText() As String
    ClipText = MaskEdBox1.ClipText
End Function
Private Function IsTime(ByVal sTime As String) As Boolean
    On Error GoTo Failed
    
    Dim bRet As Boolean
    Dim nHours As Integer
    Dim nMins As Integer
            
    If IsNumeric(Left$(sTime, 2)) Then
        nHours = CInt(Left$(sTime, 2))
        
        If IsNumeric(Right$(sTime, 2)) Then
            bRet = True
            nMins = CInt(Right$(sTime, 2))
        Else
            bRet = True
            MaskEdBox1.Text = Left$(sTime, 2) & ":00"
            nMins = 0
        End If
    Else
        If sTime = MaskEdBox1.Text Then
            MaskEdBox1.Text = "00:00"
            bRet = True
        Else
            bRet = False
        End If
        
    End If
    
    If bRet Then
        If nHours >= 0 And nHours <= 23 Then
            'Hours are valid
            If nMins >= 0 And nMins <= 59 Then
                bRet = True
            Else
                bRet = False
            End If
        Else
            bRet = False
        End If
    End If
    
    IsTime = bRet
    
    Exit Function
Failed:
    m_clsErrorHandling.RaiseError Err.Number, Err.Description
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MaskEdBox1,MaskEdBox1,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = MaskEdBox1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    MaskEdBox1.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,
Public Property Get TabStop() As Integer
    TabStop = m_TabStop
End Property

Public Property Let TabStop(ByVal New_TabStop As Integer)
    m_TabStop = New_TabStop
    PropertyChanged "TabStop"
End Property


