VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      GeneralAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   General helper object
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'PSC    16/08/99 Created
'RF     05/10/99 Added IsDigits, HasDuplicatedChars
'RF     27/10/99 Fix to HasDuplicatedChars
'DP     31/08/00 Added debugging functionality.
'CL     01/12/00 CORE000005 Correct duplicate charecters
'------------------------------------------------------------------------------------------
Option Explicit
' Private data
Private Const sPath As String = "C:\\"
Private Const TEXT_EXTENSION As String = ".txt"
Private Const XML_EXTENSION As String = ".xml"
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, _
                 ByVal lpszOp As String, _
                 ByVal lpszFile As String, _
                 ByVal lpszParams As String, _
                 ByVal lpszDir As String, _
                 ByVal FsShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Public Enum RoundingDirection
    enumRoundUp = 1
    enumRoundDown = 2
End Enum
Private Enum DEBUG_TYPE
    enumXML_DEBUG = 1
    enumTEXT_DEBUG
End Enum
Const SW_SHOWNORMAL = 1
' header ----------------------------------------------------------------------------------
' description:
'   Get the version info for a component in format XX.YY.ZZ
' pass:     n/a
' return:
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'INR    06/11/02 Added ConvertToMixedCase
'------------------------------------------------------------------------------------------
Public Function GetVersion() As String
    GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function
Public Function GetComments() As String
    GetComments = App.Comments
End Function
Public Function IsAlpha(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if all characters in vstrText are between A and Z or a and z
' pass:
'   vstrText
'       String to check
' return:
'   True  If all characters are between A and Z or a and z
'   False If any character is not between A and Z or a and z
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnIsAlpha As Boolean
    Dim intTextLen As Integer
    blnIsAlpha = True
    intIndex = 1
    intTextLen = Len(vstrText)
    If intTextLen = 0 Then
        blnIsAlpha = False
    End If
    While intIndex <= intTextLen And blnIsAlpha = True
        
        strChar = Mid$(vstrText, intIndex, 1)
        'If Asc(strChar) < 65 Or Asc(strChar) > 122 Or _
        '   (Asc(strChar) > 90 And Asc(strChar) < 97) Then
        If IsUpperAlphaChar(strChar) = False Then
            If IsLowerAlphaChar(strChar) = False Then
                blnIsAlpha = False
            End If
        End If
        intIndex = intIndex + 1
    Wend
    IsAlpha = blnIsAlpha
End Function
Public Function IsUpperAlphaChar(ByVal vstrChar As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if vstrChar is between A and Z
' pass:
'   vstrChar
'       Character to check
' return:
'------------------------------------------------------------------------------------------
    If Len(vstrChar) > 0 Then
        If Asc(vstrChar) >= 65 And Asc(vstrChar) <= 90 Then
            IsUpperAlphaChar = True
        Else
            IsUpperAlphaChar = False
        End If
    Else
         IsUpperAlphaChar = False
    End If
End Function
Public Function IsLowerAlphaChar(ByVal vstrChar As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if vstrChar is between a and z
' pass:
'   vstrChar
'       Character to check
' return:
'------------------------------------------------------------------------------------------
    If Len(vstrChar) > 0 Then
        If Asc(vstrChar) >= 97 And Asc(vstrChar) <= 122 Then
            IsLowerAlphaChar = True
        Else
            IsLowerAlphaChar = False
        End If
    Else
         IsLowerAlphaChar = False
    End If
End Function
Public Function IsDigits(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if all characters in vstrText are digits
' pass:
'   vstrText
'       String to check
' return:
'   True  If all characters are between 0 and 9 inclusive
'   False If any character is not between 0 and 9 inclusive
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnIsDigits As Boolean
    Dim intTextLen As Integer
    blnIsDigits = True
    intIndex = 1
    intTextLen = Len(vstrText)
    If intTextLen = 0 Then
        blnIsDigits = False
    End If
    While intIndex <= intTextLen And blnIsDigits = True
        
        strChar = Mid$(vstrText, intIndex, 1)
        If Asc(strChar) < 48 Or Asc(strChar) > 57 Then
             blnIsDigits = False
        End If
        intIndex = intIndex + 1
    Wend
    IsDigits = blnIsDigits
End Function
Public Function ContainsUpperAlpha(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if any character in vstrText is upper case alpha
' pass:
'   vstrText
'       String to check
' return:
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnContainsUpperAlpha As Boolean
    Dim intTextLen As Integer
    blnContainsUpperAlpha = False
    intIndex = 1
    intTextLen = Len(vstrText)
    While intIndex <= intTextLen And blnContainsUpperAlpha = False
        
        strChar = Mid$(vstrText, intIndex, 1)
        If Me.IsUpperAlphaChar(strChar) Then
             blnContainsUpperAlpha = True
        End If
        intIndex = intIndex + 1
    Wend
    ContainsUpperAlpha = blnContainsUpperAlpha
End Function
Public Function ContainsLowerAlpha(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if any character in vstrText is lower case alpha
' pass:
'   vstrText
'       String to check
' return:
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnContainsLowerAlpha As Boolean
    Dim intTextLen As Integer
    blnContainsLowerAlpha = False
    intIndex = 1
    intTextLen = Len(vstrText)
    While intIndex <= intTextLen And blnContainsLowerAlpha = False
        
        strChar = Mid$(vstrText, intIndex, 1)
        If Me.IsLowerAlphaChar(strChar) Then
             blnContainsLowerAlpha = True
        End If
        intIndex = intIndex + 1
    Wend
    ContainsLowerAlpha = blnContainsLowerAlpha
End Function
Public Function ContainsDigits(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if any character in vstrText is a digit
' pass:
'   vstrText
'       String to check
' return:
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnContainsDigits As Boolean
    Dim intTextLen As Integer
    blnContainsDigits = False
    intIndex = 1
    intTextLen = Len(vstrText)
    While intIndex <= intTextLen And blnContainsDigits = False
        
        strChar = Mid$(vstrText, intIndex, 1)
        If Me.IsDigits(strChar) Then
             blnContainsDigits = True
        End If
        intIndex = intIndex + 1
    Wend
    ContainsDigits = blnContainsDigits
End Function
Public Function ContainsSpecialChars(ByVal vstrText As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks if any character in vstrText is a special char (i.e. non-alpha and
'   non-numeric)
' pass:
'   vstrText
'       String to check
' return:
'------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strChar As String
    Dim blnContainsSpecialChars As Boolean
    Dim intTextLen As Integer
    blnContainsSpecialChars = False
    intIndex = 1
    intTextLen = Len(vstrText)
    While intIndex <= intTextLen And blnContainsSpecialChars = False
        
        strChar = Mid$(vstrText, intIndex, 1)
        If Me.IsAlpha(strChar) = False Then
            If Me.IsDigits(strChar) = False Then
                blnContainsSpecialChars = True
            End If
        End If
        intIndex = intIndex + 1
    Wend
    ContainsSpecialChars = blnContainsSpecialChars
End Function
Public Function HasDuplicatedChars(ByVal vstrText As String, _
    ByVal vblnCaseSensitive As Boolean) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Checks all the alphas or digits in a string are the same:
'   If >1 digit, check if digits all the same
'   If >1 alpha, check if alphas all the same (case sensitivity according to
'   vblnCaseSensitive)
' pass:
'   vstrText
'       String to check
' return:
'   True    If there is character duplication.
'           E.g. "AAAa123", "111AAA111", "Q2q345", "Fred11"
'   False   If there is no character duplication
'           E.g. "AAA3BBB", "11111112"
'------------------------------------------------------------------------------------------
    Dim intNumAlphas, intNumDigits As Integer
    Dim intIndex As Integer
    Dim strChar, strLastAlphaChar, strLastDigit As String
    Dim blnHasDuplicatedAlphas, blnHasDuplicatedDigits, _
        blnHasDuplicatedChars As Boolean
    Dim intTextLen As Integer
    blnHasDuplicatedChars = False
    blnHasDuplicatedAlphas = False
    blnHasDuplicatedDigits = False
    intIndex = 1
    intTextLen = Len(vstrText)
    While intIndex <= intTextLen And _
          blnHasDuplicatedAlphas = False And blnHasDuplicatedDigits = False
        strChar = Mid$(vstrText, intIndex, 1)
        If IsAlpha(strChar) = True Then
            intNumAlphas = intNumAlphas + 1
            If vblnCaseSensitive = False Then
                strChar = UCase(strChar)
            End If
            If intNumAlphas > 1 Then
                If strChar = strLastAlphaChar Then
                    blnHasDuplicatedAlphas = True
                End If
            End If
            strLastAlphaChar = strChar
        Else
            If IsDigits(strChar) = True Then
                intNumDigits = intNumDigits + 1
                If intNumDigits > 1 Then
                    If strChar = strLastDigit Then
                        blnHasDuplicatedDigits = True
                    End If
                End If
            End If
            strLastDigit = strChar
        End If
        intIndex = intIndex + 1
    Wend
    If (blnHasDuplicatedAlphas = True And intNumAlphas > 1) Or _
       (blnHasDuplicatedDigits = True And intNumDigits > 1) Then
        blnHasDuplicatedChars = True
    End If
    HasDuplicatedChars = blnHasDuplicatedChars
End Function
Function StartDoc(sDocName As String) As Long
    On Error GoTo Failed
    Dim Scr_hDC As Long
    Dim strFunctionName As String
    strFunctionName = "StartDoc"
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, _
                                "Open", _
                                sDocName, _
                                "", _
                                "C:\", _
                                SW_SHOWNORMAL)
    Exit Function
Failed:
    StartDoc = False
End Function
Public Sub DisplayDebug(sDescription As String, Optional objDom As Object)
    On Error GoTo Failed
    Dim sExt As String
    Dim sFile As String
    Dim sText As String
    Dim strFunctionName As String
    Dim enumDebugType As DEBUG_TYPE
    strFunctionName = "DisplayDebug"
    If Not objDom Is Nothing Then
        If TypeOf objDom Is FreeThreadedDOMDocument40 Then
            sText = FreeThreadedDOMDocument40(objDom).xml
            enumDebugType = enumXML_DEBUG
        ElseIf TypeOf objDom Is IXMLDOMElement Or TypeOf objDom Is IXMLDOMNode Then
            sText = DomNode(objDom).xml
            enumDebugType = enumXML_DEBUG
        End If
    Else
        enumDebugType = enumTEXT_DEBUG
    End If
    Select Case enumDebugType
    Case enumXML_DEBUG
        sFile = sPath & sDescription & XML_EXTENSION
    Case enumTEXT_DEBUG
        sText = sDescription
    End Select
    If WriteDebug(sText, sFile, False) Then
        StartDoc sFile
    End If
    Exit Sub
Failed:
    ' Do nothing
End Sub
Public Function WriteDebug(sDebug As String, Optional sFile As String, Optional bAppend As Boolean = True) As Boolean
    On Error GoTo Failed
    Dim fd As Integer
    Dim strFunctionName As String
    strFunctionName = "WriteDebug"
    fd = FreeFile
    If Len(sFile) = 0 Then
        sFile = sPath & "Omiga4Debug.txt"
    End If
    If bAppend Then
        sDebug = Now & ": " & sDebug
        Open sFile For Append Access Write As fd
    Else
        Open sFile For Output Access Write As fd
    End If
    Print #fd, sDebug
    Close #fd
    WriteDebug = True
    Exit Function
Failed:
    WriteDebug = False
End Function
Private Function DomNode(objDom As IXMLDOMNode) As IXMLDOMNode
    Set DomNode = objDom
End Function
Public Function RoundToNearest(dblAmt As Double, dblRoundAmt As Double, _
                               Optional enumDirection As RoundingDirection = enumRoundUp) As Double
        On Error Resume Next
        Dim dblTemp As Double
        dblTemp = dblAmt / dblRoundAmt
        If Int(dblTemp) = dblTemp Then
           RoundToNearest = dblAmt
        Else
           If enumDirection = enumRoundDown Then
              dblTemp = Int(dblTemp)
           Else
              dblTemp = Int(dblTemp) + 1
           End If
           RoundToNearest = dblTemp * dblRoundAmt
        End If
End Function
' header ----------------------------------------------------------------------------------
' description:
'   Converts a string to Mixed Case.
'   The first Char of each word/s will be returned in uppercase.
'   The remainder of each word/s will be returned in lower case.
'
' pass:
'   strInToConvert
'       String to convert
' return:
'   strMixedCase
'
'   e.g. strInToConvert "strInToConvert sssSSSSS"
'        strMixedCase   "Strintoconvert Ssssssss"
'------------------------------------------------------------------------------------------
Public Function ConvertToMixedCase(ByVal strInToConvert As String) As String
    Dim nCharFinishPos, nCharStartPos, nPosSpace As Integer
    Dim strMixedCase, strLastDigit As String
    Dim bHasSpaces, bFirstTime As Boolean
    'make everthing consistent
    strInToConvert = LCase(strInToConvert)
    nPosSpace = InStr(1, strInToConvert, " ", vbTextCompare)
    nCharStartPos = 1
    'Do we have any spaces
    If nPosSpace = 0 Then
        nCharFinishPos = Len(strInToConvert) + 1
    Else
        nCharFinishPos = nPosSpace + 1
    End If
    'While Have Spaces,
    While nPosSpace > 0
        'Capitalise the first char after the space
        strMixedCase = strMixedCase & UCase(Mid(strInToConvert, nCharStartPos, 1))
        nCharStartPos = nCharStartPos + 1
        'everything else to the next space remains lowercase
        strMixedCase = strMixedCase & Mid(strInToConvert, nCharStartPos, nCharFinishPos - nCharStartPos)
        nCharStartPos = nPosSpace + 1
        nPosSpace = InStr(nCharStartPos, strInToConvert, " ", vbTextCompare)
        'Any more spaces to deal with
        If nPosSpace = 0 Then
            nCharFinishPos = Len(strInToConvert) + 1
        Else
            nCharFinishPos = nPosSpace + 1
        End If
    Wend
    'Make sure we actually have something
    If (nCharFinishPos > nCharStartPos) Then
        strMixedCase = strMixedCase & UCase(Mid(strInToConvert, nCharStartPos, 1))
        nCharStartPos = nCharStartPos + 1
        strMixedCase = strMixedCase & Mid(strInToConvert, nCharStartPos, nCharFinishPos - nCharStartPos)
    End If
    ConvertToMixedCase = strMixedCase
End Function
Private Function FreeThreadedDOMDocument40(objDom As FreeThreadedDOMDocument40) As FreeThreadedDOMDocument40
    Set FreeThreadedDOMDocument40 = objDom
End Function
