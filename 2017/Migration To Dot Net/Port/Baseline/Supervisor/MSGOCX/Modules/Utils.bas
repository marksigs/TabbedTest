Attribute VB_Name = "Utils"
Option Explicit

Private Const cstrAPOSTROPHE As String = "'"
Private Const cstrWILDCARD As String = "*"
Private Const cstrWILDCARDSQL = "%"
Public Const g_sStandardCharSet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz',.-/&0123456789"
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
Public Function StripCharFromString(ByRef sString As String, ByRef sChar As String) As String
    Dim nStart As Long
    Dim nLen As Long
    Dim nPos As Long
    Dim sThisChar As String
    Dim sTmp As String
    
    nLen = Len(sString)
    
    For nPos = 1 To nLen
        sThisChar = Mid(sString, nPos, 1)
    
        If (sThisChar <> sChar) Then
            sTmp = sTmp + sThisChar
        End If
    Next nPos
    
    StripCharFromString = sTmp
End Function

Public Function GetKeyFromString(sName As String) As String
    GetKeyFromString = StripCharFromString(sName, " ")
End Function

Public Function StripUnderscoreAndSlash(ByRef sTxt As String) As String
    Dim sTmp As String
    
    sTmp = StripCharFromString(sTxt, "_")
    StripUnderscoreAndSlash = StripCharFromString(sTmp, "/")
End Function
Public Function StripUnderscore(ByRef sTxt As String) As String
    StripUnderscore = StripCharFromString(sTxt, "_")
End Function

Public Function IsDateValid(sDate As String) As Boolean
    Dim sTmp As String
    
    sTmp = StripUnderscoreAndSlash(sDate)

    If Len(sTmp) > 0 Then
        IsDateValid = True
    Else
        IsDateValid = False
    End If

End Function
Public Function IsControl(KeyCode As Integer) As Boolean
    If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
       KeyCode = vbKeyHome Or KeyCode = vbKeyEnd Or KeyCode = vbKeyShift Or KeyCode = vbKeyControl Or _
       KeyCode = vbKeyBack Or KeyCode = vbKeyTab Or KeyCode = 18 Then
        
        IsControl = True
    Else
        IsControl = False
    End If
End Function
Public Function IsValidControl(ByRef nVal As Integer)
    If (nVal < 30) Then
        IsValidControl = True
    Else
        IsValidControl = False
    End If
End Function

Public Function IsDigit(ByRef sVal As String, ByRef sText As String)
    Dim nPos As Long
    Dim bRet As Boolean
    
    bRet = False
    
    If (sVal = ".") Then
        nPos = InStr(sText, sVal)
        
        If (nPos = 0) Then
            bRet = True
        End If
    Else
        If (IsNumeric(sVal)) Then
            bRet = True
        End If
    End If
    IsDigit = bRet
End Function

Public Function FormatString(sStr As String) As String
    FormatString = CheckApostrophes(sStr)
End Function
Private Function CheckApostrophes(ByVal strIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Doubles each ' in a string. This is needed for names with
'   an apostrophe in them. E.g. "O'Conner" converts to "O''Conner"
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo errhandler
    
    Dim nPos As Integer
    Dim strOut As String
    Dim nLen As Integer
    Dim thisChar As String
    Dim i As Integer
    
    nLen = Len(strIn)
    
    For nPos = 1 To nLen Step 1
        thisChar = Mid$(strIn, nPos, 1)
        strOut = strOut + thisChar
        If thisChar = cstrAPOSTROPHE Then
            ' add an extra apostrophe
            strOut = strOut + cstrAPOSTROPHE
        End If
    Next nPos
    
    CheckApostrophes = strOut
    
    Exit Function

errhandler:

End Function
Public Sub BeginWaitCursor()
    Screen.MousePointer = vbHourglass
End Sub
Public Sub EndWaitCursor()
    Screen.MousePointer = vbDefault
End Sub
Public Function GetConnectionKey(sService As String, sUserID As String) As String
    GetConnectionKey = sService & ", " & sUserID
End Function
Public Function GetFullYear(sDate As String) As String
    On Error GoTo Failed
    Dim sTmp As String
    Dim sYear As String
    Dim tmpDate As Date
    Dim nYearStart As Long
    Dim sFullDate As String
    Dim nFirstDelim As Long
    sTmp = StripUnderscore(sDate)
    
    If (IsDate(sTmp)) Then
        tmpDate = CDate(sTmp)
        
        nYearStart = InStr(1, sDate, "/")
        
        If (nYearStart >= 1) Then
            If (nYearStart = 2) Then
                sFullDate = "0" + Left(sTmp, 1)
            Else
                sFullDate = Left(sTmp, 2)
            End If
            nFirstDelim = nYearStart
            nYearStart = InStr(nYearStart + 1, sTmp, "/")
        End If
        
        If (nYearStart > 1) Then
            sYear = Year(tmpDate)
                        
            If (nYearStart - nFirstDelim = 2) Then
                sFullDate = sFullDate + "/0" + Mid(sTmp, nFirstDelim + 1, 1)
            Else
                sFullDate = sFullDate + "/" + Mid(sTmp, nFirstDelim + 1, 2)
            End If
            
            sTmp = sFullDate + "/" + sYear
        End If
    End If
    
    GetFullYear = sTmp
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Description
End Function
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
         0, flags)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, flags)
      SetTopMostWindow = False
   End If
End Function
Public Function GetBooleanFromNumber(nNum As Long) As String
    If nNum = 0 Then
        GetBooleanFromNumber = "False"
    Else
        GetBooleanFromNumber = "True"
    End If
End Function
Public Function GetNumberFromBoolean(bBool As Boolean) As Byte
    
    If bBool Then
        GetNumberFromBoolean = 1
    ElseIf Not bBool Then
        GetNumberFromBoolean = 0
    End If

End Function

Public Function GetBooleanFromNumberAsBoolean(nNum As Byte) As Boolean

    If nNum = 0 Then
        GetBooleanFromNumberAsBoolean = False
    Else
        GetBooleanFromNumberAsBoolean = True
    End If
    
End Function

'BMIDS923  New function

Public Function Check_IsDate(sDate As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check that value is a date of the correct format
'------------------------------------------------------------------------------------------
    On Error GoTo Failed

    Dim aDaysInMonth(12)
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim iYear As Integer
    Dim iDaysinMonth As Integer
    
    aDaysInMonth(0) = 31
    aDaysInMonth(1) = 28
    aDaysInMonth(2) = 31
    aDaysInMonth(3) = 30
    aDaysInMonth(4) = 31
    aDaysInMonth(5) = 30
    aDaysInMonth(6) = 31
    aDaysInMonth(7) = 31
    aDaysInMonth(8) = 30
    aDaysInMonth(9) = 31
    aDaysInMonth(10) = 30
    aDaysInMonth(11) = 31

    'Use IsDate first to perform a rudimentary validation
    If (IsDate(sDate)) Then
    
        'Now perform a more detailed validation of the Day and Month
    
        'sDate = ##/##/####
        
        iYear = CInt(Mid(sDate, 7, 4))
        iMonth = CInt(Mid(sDate, 4, 2))
        iDay = CInt(Mid(sDate, 1, 2))

        If (iMonth < 13) And (iMonth > 0) Then
        
            'Set up days in month
            If (iMonth <> 2) Then
                iDaysinMonth = aDaysInMonth(iMonth - 1)
            Else
                'Check for leap year
                If (iYear Mod 4 = 0 And iYear Mod 100 <> 0) Or (iYear Mod 400 = 0) Then
                    iDaysinMonth = 29
                Else
                    iDaysinMonth = 28
                End If
            End If
            
            If (iDay <= iDaysinMonth) And (iDay > 0) Then
                Check_IsDate = True
            Else
                ' Day is invalid
                Check_IsDate = False
            End If
        Else
            ' Month is invalid
            Check_IsDate = False
        End If
             
    Else
        Check_IsDate = False
    End If

    Exit Function
Failed:
    Err.Raise Err.Number, Err.Description

End Function


