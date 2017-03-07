VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : Validation.cls
' Description   : Contains generic routines to aid validation of form data
' Change history
' Prog      Date        Description
' DJP       21/06/01    Removed GetDateString as it was no longer needed (as part of SQL Server port)
' INR       18/12/06    EP2_517 error message spelling error.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bReportError As Boolean
Private m_bHighLight As Boolean
Private m_bMandFocus As Boolean

Public Function ValidateDate(ByRef txtDate As MSGEditBox, bNull As Boolean, Optional bNullIsValid = True) As Boolean
    Dim vTmp As Variant
    Dim tmpDate As Date
    Dim bRet As Boolean
    Dim nLen As Long
    Dim sYear As String
    bRet = False

    nLen = Len(txtDate.Text) - 2  ' For the // in 12/12/1998
    
    bNull = False
    
    If (nLen > 0) Then
        On Error Resume Next
        
        If (InStr(1, txtDate.Text, "_") = 0) Then
            bRet = True
        Else
            bRet = False
        End If
    Else
        If (m_bHighLight = True And bNullIsValid = False) Then
            HighLightdate txtDate
            bRet = False
        ElseIf (bNullIsValid) Then
            bRet = True
        Else
            bRet = False
        End If
        bNull = True
    End If
    
    If (bRet = True And nLen > 0) Then
        SetDateNormal txtDate
    End If
    
    If (bRet = False) Then
        If (m_bReportError = True) Then
            MsgBox "Invalid Date Entered", vbCritical
            SetFocus txtDate
        End If
    End If
    ValidateDate = bRet
    
End Function

Public Function ValidateCombo(ByRef cboField As ComboBox, Optional bNullIsValid As Boolean = True) As Boolean
    Dim bRet As Boolean
    bRet = False
    
    If (Len(cboField.Text) > 0) Then
        bRet = True
    End If
    
    If (bRet = False) Then
        If (m_bReportError = True) Then
            MsgBox "Value must be entered", vbCritical
            SetFocusCombo cboField
        End If
            
        If (m_bHighLight = True) Then
            HighLightCombo cboField
        End If
    Else
        SetComboNormal cboField
    End If
    ValidateCombo = bRet
    
End Function

Public Function IsValidControl(ByRef nVal As Integer)
    If (nVal < 30) Then
        IsValidControl = True
    Else
        IsValidControl = False
    End If
End Function

Public Function IsDigit(ByRef vVal As Variant, ByRef txtControl As MSGEditBox)
    Dim nPos As Long
    Dim bRet As Boolean
    
    bRet = False
    
    
    If (vVal = ".") Then
        nPos = InStr(txtControl.Text, vVal)
        
        If (nPos = 0) Then
            bRet = True
        End If
    Else
        If (IsNumeric(vVal)) Then
            bRet = True
        End If
    End If
    IsDigit = bRet
End Function

Public Function ValidateNumber(ByRef txtNumber As MaskEdBox, Optional bZeroIsValid = False, Optional bNullIsValid = True)
    Dim nLen As Long
    Dim bRet As Boolean
    Dim sError As String
    
    bRet = False
    nLen = Len(txtNumber.Text)
    
    If (nLen = 0) Then
        ' Nothing entered
        If (bNullIsValid = True) Then
            bRet = True
        Else
            sError = "Mandatory field - number must be entered"
        End If
    Else
        Dim nVal As Double
        nVal = Val(txtNumber.Text)
        ' Something entered, is it valid?
        If (nVal > 0) Then
            bRet = True
        ElseIf (bZeroIsValid = True And nVal = 0) Then
            bRet = True
        Else
            sError = "Value cannot be less than or equal to 0"
        End If
    End If

    If (bRet = False) Then
        If (nLen = 0 And m_bHighLight = True) Then
            HighLightTextField txtNumber
        End If
        
        If (m_bReportError = True) Then
            If (Len(sError) > 0) Then
                MsgBox sError, vbCritical
            End If
            SetFocusText txtNumber
        End If
    Else
        If (nLen > 0) Then
            SetTextNormal txtNumber
        End If
    End If
    ValidateNumber = bRet

End Function

Public Sub SetTextNormal(txtNumber As MaskEdBox)
    txtNumber.BackColor = vbWhite
End Sub

Public Sub SetDateNormal(txtDate As MaskEdBox)
    txtDate.BackColor = vbWhite
End Sub

Public Sub HighLightTextField(txtNumber As MaskEdBox)
    txtNumber.BackColor = vbRed '&H80FF&
    
    If (m_bReportError = True Or m_bMandFocus = True) Then
        txtNumber.SetFocus
        m_bMandFocus = False
    End If
End Sub

Public Sub HighLightdate(txtDate As MaskEdBox)
    On Error GoTo Failed
    txtDate.BackColor = vbRed '&H80FF&
    If (m_bReportError = True Or m_bMandFocus = True) Then
        txtDate.SetFocus
        m_bMandFocus = False
    End If
    Exit Sub
    
Failed:
    MsgBox ("HighLightDate: Unable to highlight - Error is: " + Err.DESCRIPTION)
End Sub

Public Sub HighLightCombo(cboField As ComboBox)
    cboField.BackColor = vbRed '&H80FF&
    If (m_bReportError = True Or m_bMandFocus = True) Then
        cboField.SetFocus
        m_bMandFocus = False
    End If
End Sub

Public Sub SetComboNormal(cboField As ComboBox)
    cboField.BackColor = vbWhite
End Sub

Private Function GetQuarterlyDate(ByRef sYear As String, dateQuarter As Date, Optional bGetYear As Boolean = True) As Variant
    Dim sDayAndMonth As String
    
    sDayAndMonth = GetDayAndMonth(CStr(dateQuarter))
    
    If (bGetYear = True) Then
        GetQuarterlyDate = sDayAndMonth + "/" + sYear
    Else
        GetQuarterlyDate = sDayAndMonth
    End If
End Function

Private Function GetYear(ByRef thisDate As Date) As String
    Dim nLen As Long
    Dim nIndex As Long
    Dim nRet As Long
    Dim sYear As Variant
    Dim sDate As String
    sYear = ""
    
    sDate = thisDate
    nLen = Len(sDate)
    nRet = InStr(1, sDate, "/", 1)
    
    If (nRet > 0) Then
        nRet = InStr(nRet + 1, sDate, "/", 1)
    
        If (nRet > 0) Then
            sYear = Right(sDate, nLen - nRet)
        End If
    End If
    GetYear = sYear
End Function

Public Sub SetFocus(ByRef txtDate As MSGEditBox)
    On Error GoTo Error
    Dim nLen As String
    Dim sTmp As String
    
    txtDate.SetFocus
    sTmp = StripUnderscoreAndSlash(txtDate.Text)
    
    nLen = Len(sTmp)
    
    If (nLen > 0) Then
                
        txtDate.SelStart = 0
        txtDate.SelLength = Len(txtDate.Text)
    End If
    
Error:
End Sub

Public Sub SetFocusText(ByRef txtDate As MaskEdBox)
    On Error GoTo Error
    txtDate.SetFocus
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate)

Error:
End Sub

Public Sub SetFocusCombo(ByRef cboField As ComboBox)
    On Error GoTo Error
    cboField.SetFocus
Error:
End Sub

Public Sub SetReportError(Optional bReport As Boolean = True)
    m_bReportError = bReport
    If (bReport = True) Then
        m_bMandFocus = False
    Else
        m_bMandFocus = True
    End If
End Sub

Public Sub HighLight(Optional bHighLight As Boolean = True)
    m_bHighLight = bHighLight
End Sub

Private Sub Class_Initialize()
    m_bReportError = True
    m_bHighLight = False
End Sub

Public Function CheckDecimalPlaces(sText As String, nBeforePoint, nAfterPointValid As Long)
    Dim bRet As Boolean
    Dim nPos As Long
    Dim nSel As Long
    Dim nAfterPoint As Long
    
    bRet = True
    ' Need to work out how many places are after the point. > 1 and return false, else true
    
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
            Dim sTmp As String
            Dim nLen As Long
            
            sTmp = Left(sText, nPos - 1)
            
            If (nPos > nBeforePoint) Then
                bRet = False
            End If
        End If
    End If

    CheckDecimalPlaces = bRet
End Function

Public Sub ClearDate(dateEdit As MaskEdBox)
    Dim sTmp As String
    sTmp = dateEdit.Mask
    dateEdit.Mask = ""
    dateEdit.Text = ""
    dateEdit.Mask = sTmp
End Sub

Public Sub EnableText(txtEdit As MaskEdBox, Optional bEnable As Boolean = True, Optional bClear As Boolean = True)
    If (bClear = True) Then
        txtEdit = ""
    End If
    txtEdit.Enabled = bEnable
End Sub

Public Sub HandleNumber(txtField As Object, nIndex As Integer, KeyAscii As Integer)
'    Dim bRet As Boolean
'
'    bRet = m_clsvalidation.IsValidControl(KeyAscii)
'
'    If (bRet = False) Then
'        bRet = m_clsvalidation.IsDigit(Chr(KeyAscii), txtField(nIndex))
'
'        If (bRet = True) Then
'            If (bRet = True) Then
'                ' More than nNumDigits?
'                Dim nBeforePoint As Long
'                Dim nAfterPoint As Long
'
'                Select Case nIndex
'                    Case TOTAL_TRANSFER_VALUE, GMP_PROTECTED_RIGHTS, AVC_TRANSFER_VALUE, _
'                            GMP_CURRENT_FUND, CURRENT_FUND_VALUE, ORIGINAL_VAL_EXCLUDING_AVC, _
'                            PRE_88_GMP, POST_88_GMP, NGMP_REVALUED, NGMP_NON_REVALUED, TAX_FREE_SUM
'
'                        nBeforePoint = 7
'                        nAfterPoint = 2
'                    Case RETIREMENT_AGE_MONTHS, RETIREMENT_AGE_YEARS
'                        nBeforePoint = 2
'                        nAfterPoint = 0
'                    Case MANAGEMENT_CHARGE
'                        nBeforePoint = 2
'                        nAfterPoint = 2
'                    Case Else
'                        nBeforePoint = 7
'                        nAfterPoint = 2
'                End Select
'
'                Dim nSel As Long
'                nSel = txtField(nIndex).SelLength
'
'                If (nSel = 0) Then
'                    Dim sTmp As String
'                    Dim stmp1 As String
'                    Dim nStart As Long
'                    Dim nLen As Long
'
'                    nLen = Len(txtField(nIndex).Text)
'                    nStart = txtField(nIndex).SelStart
'                    sTmp = Left(txtField(nIndex).Text, nStart)
'                    stmp1 = sTmp + Chr(KeyAscii)
'                    sTmp = stmp1 + Right(txtField(nIndex).Text, nLen - nStart)
'
'                    bRet = m_clsvalidation.CheckDecimalPlaces(sTmp, nBeforePoint, nAfterPoint)
'                Else
'                    bRet = True
'                End If
'            End If
'        End If
'    End If
'
'    If (bRet = False) Then
'        KeyAscii = 0
'    End If
End Sub

Public Function GetFullYear(sDate As String) As String
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
End Function

Public Function GetDayAndMonth(sDate As String) As String
    Dim nYearStart As Long
    
    nYearStart = InStr(1, sDate, "/")
    
    If (nYearStart >= 1) Then
        nYearStart = InStr(nYearStart + 1, sDate, "/")
    End If

    If (nYearStart > 1) Then
        GetDayAndMonth = Left(sDate, nYearStart - 1)
    End If

End Function

Friend Sub ValidateProductDates(txtStartDate As MSGEditBox, txtEndDate As MSGEditBox, Optional txtWithdrawnDate As MSGEditBox = Nothing)
    On Error GoTo Failed
    Dim vStartDate As Variant
    Dim sStartDate As Date
    Dim vWithdrawnDate As Variant
    Dim sEndDate As Date
    Dim vEndDate As Variant
    Dim sSystemDate As Date
    
    sSystemDate = Format(Now, "Short Date")

    g_clsFormProcessing.HandleDate txtStartDate, vStartDate, GET_CONTROL_VALUE
    g_clsFormProcessing.HandleDate txtEndDate, vEndDate, GET_CONTROL_VALUE
    
    If Not txtWithdrawnDate Is Nothing Then
        g_clsFormProcessing.HandleDate txtWithdrawnDate, vWithdrawnDate, GET_CONTROL_VALUE
    End If
    
    If Not IsNull(vStartDate) Then
        sStartDate = CDate(vStartDate)
    Else
        g_clsFormProcessing.SetControlFocus txtStartDate
        
        g_clsErrorHandling.RaiseError errGeneralError, "Mortgage Product - Start Date must be valid"
    End If

    If IsNull(vEndDate) Then
        g_clsFormProcessing.SetControlFocus txtEndDate
        g_clsErrorHandling.RaiseError errGeneralError, "Mortgage Product - End Date must be valid"
    Else
        sEndDate = CDate(vEndDate)
    End If
    
    If (Not IsNull(vWithdrawnDate)) And (Not IsEmpty(vWithdrawnDate)) Then
        ' Must be greater than Start Date
        If CDate(vWithdrawnDate) <= CDate(vStartDate) Then
            g_clsFormProcessing.SetControlFocus txtWithdrawnDate
            g_clsErrorHandling.RaiseError errGeneralError, "Withdrawn Date must be greater than Start Date"
        End If
    End If
    
    If sEndDate <= sSystemDate Then
        g_clsFormProcessing.SetControlFocus txtEndDate
        g_clsErrorHandling.RaiseError errGeneralError, "End Date must be greater than today's date"
    End If
    
    If sEndDate <= sStartDate Then
        g_clsFormProcessing.SetControlFocus txtEndDate
        g_clsErrorHandling.RaiseError errGeneralError, "End Date must be greater than Start Date"
    End If
    
    If Not IsNull(vWithdrawnDate) Then
        If CDate(sEndDate) < CDate(vWithdrawnDate) Then
            g_clsFormProcessing.SetControlFocus txtEndDate
            g_clsErrorHandling.RaiseError errGeneralError, "End Date cannot be less than Withdrawn Date"
        End If
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Function ValidateActiveFromTo(txtStartDate As MSGEditBox, txtEndDate As MSGEditBox) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim vStartDate As Variant
    Dim sStartDate As Date
    Dim sEndDate As Date
    Dim vEndDate As Variant
    bRet = True
    
    g_clsFormProcessing.HandleDate txtStartDate, vStartDate, GET_CONTROL_VALUE
    g_clsFormProcessing.HandleDate txtEndDate, vEndDate, GET_CONTROL_VALUE
    
    If Not IsNull(vStartDate) Then
        sStartDate = CDate(vStartDate)
    Else
        MsgBox "Active From Date must be valid", vbCritical
        If txtStartDate.Visible = True And txtStartDate.Enabled = True Then
            g_clsFormProcessing.SetControlFocus txtStartDate
        End If
        bRet = False
    End If

    If bRet = True Then
        If Not IsNull(vEndDate) Then
            sEndDate = CDate(vEndDate)
        
            If sEndDate <= sStartDate Then
                MsgBox "Active To Date must be greater than Active From Date", vbCritical
                txtEndDate.SetFocus
                bRet = False
            End If
        Else
            bRet = True
        End If
    End If
    
    ValidateActiveFromTo = bRet
    Exit Function
Failed:
    MsgBox "ValidateActiveFromTo: Error is " + Err.DESCRIPTION
    ValidateActiveFromTo = False

End Function

Public Function GetDate(sDate As String) As String
    GetDate = CStr(Format(CDate(sDate), "dd/mm/yyyy"))
End Function

Public Function GetFullDate(sDate As String) As String
    Dim sTestDate As String
    Dim nPos As Integer
    Dim nLastPos As Integer
    Dim sMask As String
    Dim nThisPos
    Dim sNewDate As String
    
    nLastPos = InStr(1, sDate, "/")
    
    If nLastPos > 0 Then
        If nLastPos = 2 Then
            sNewDate = "0" + Left(sDate, 1)
            sNewDate = sNewDate + "/"
        Else
            sNewDate = Left(sDate, 2) + "/"
        End If
        
        nLastPos = nLastPos + 1
        nPos = InStr(nLastPos + 1, sDate, "/")
        
        If nPos - nLastPos = 1 Then
            sNewDate = sNewDate + "0" + Mid(sDate, nLastPos, 1)
        Else
            sNewDate = sNewDate + Mid(sDate, nLastPos, 2)
        End If
    
        sNewDate = sNewDate + "/" + Right(sDate, Len(sDate) - nPos)
    End If
    
    GetFullDate = sNewDate

End Function

Public Function ValidateTime(sTime As String) As Boolean
    ' Validate the product start time
    Dim nTimePart As Long
    Dim bTimeValid As Boolean
    Dim nTimeLen As String
    Dim sTimePart As String
    nTimeLen = Len(sTime)
    
    bTimeValid = False
    
    If nTimeLen > 0 And nTimeLen < 6 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Time must be in format of hh:mm:ss"
    End If
    
    ' 6 for hh:mm:ss
    
    ' Hour portion
    If nTimeLen > 0 Then
        sTimePart = Left(sTime, 2)
        
        If IsNumeric(sTimePart) Then
            If sTimePart >= 0 And sTimePart <= 23 Then
                ' Minute
                sTimePart = Mid(sTime, 3, 2)
                
                If IsNumeric(sTimePart) Then
                    If sTimePart >= 0 And sTimePart <= 59 Then
                        sTimePart = Mid(sTime, 5, 2)
                        
                        If IsNumeric(sTimePart) Then
                            If sTimePart >= 0 And sTimePart <= 59 Then
                                bTimeValid = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        bTimeValid = True
    End If
    
    ValidateTime = bTimeValid
End Function

Public Function ValidatePostCode(sText As String) As Boolean
    
    Dim colPostCodes(3 To 8) As Collection, bValidPostCode As Boolean, iPostCodePos As Integer, iCollNo As Integer, iCollCnt As Integer
    
    For iCollCnt = 3 To 8
        Set colPostCodes(iCollCnt) = New Collection
    Next
    
    colPostCodes(3).Add "@@#"
    colPostCodes(3).Add "@#@"
    colPostCodes(3).Add "@##"
    colPostCodes(3).Add "IOM"
    
    colPostCodes(4).Add "@@#@"
    colPostCodes(4).Add "@# #@"
    colPostCodes(4).Add "BT74"
    
    colPostCodes(5).Add "@@# #"
    colPostCodes(5).Add "@#@ #"
    colPostCodes(5).Add "@## #"
    
    colPostCodes(6).Add "@@#@ #"
    colPostCodes(6).Add "@@## #"
    colPostCodes(6).Add "@# #@@"

    colPostCodes(7).Add "@@# #@@"
    colPostCodes(7).Add "@#@ #@@"
    colPostCodes(7).Add "@## #@@"
    
    colPostCodes(8).Add "@@#@ #@@"
    colPostCodes(8).Add "@@## #@@"
    
    On Error GoTo Failed
       
    If Len(sText) < 3 Or Len(sText) > 9 Then
        If sText = Space(Len(sText)) Then
            bValidPostCode = True
        Else
            bValidPostCode = False
        End If
        ValidatePostCode = bValidPostCode
        Exit Function
    End If
    
    iCollNo = Len(sText)
    For iCollCnt = 1 To colPostCodes(iCollNo).Count
        If Not bValidPostCode Then
            For iPostCodePos = 1 To iCollNo
                Select Case Mid(colPostCodes(iCollNo).Item(iCollCnt), iPostCodePos, 1)
                    Case "@"
                       ' Is the postcode character between A and Z?
                        If Mid(sText, iPostCodePos, 1) >= "A" And Mid(sText, iPostCodePos, 1) <= "Z" Then
                            bValidPostCode = True
                        Else
                            bValidPostCode = False
                            Exit For
                        End If
                    Case "#"
                       ' Is the postcode character between 0 and 9?
                        If Mid(sText, iPostCodePos, 1) >= "0" And Mid(sText, iPostCodePos, 1) <= "9" Then
                            bValidPostCode = True
                        Else
                            bValidPostCode = False
                            Exit For
                        End If
                    Case Else
                       'If the postcode character isn't @(String) and #(Numeric) then the inputted postcode character must be equal the mask character
                        If Mid(colPostCodes(iCollNo).Item(iCollCnt), iPostCodePos, 1) = Mid(sText, iPostCodePos, 1) Then
                            bValidPostCode = True
                        Else
                            bValidPostCode = False
                            Exit For
                        End If
                    End Select
            Next iPostCodePos
        Else
            Exit For
        End If
    Next iCollCnt
    
    ValidatePostCode = bValidPostCode

    Exit Function

Failed:
    Err.Raise Err.Number, Err.DESCRIPTION
End Function

Public Function StripCharFromString(ByRef sString As String, ByRef sChar As String)
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

Public Function StripUnderscoreAndSlash(ByRef sTxt As String) As String
    Dim sTmp As String
    
    sTmp = StripCharFromString(sTxt, "_")
    StripUnderscoreAndSlash = StripCharFromString(sTmp, "/")
End Function

Public Function StripUnderscore(ByRef sTxt As String) As String
    StripUnderscore = StripCharFromString(sTxt, "_")
End Function

'BMIDS00987 New funtion to validate sort code
Friend Function ValidateSortCode(txtSortCode As MSGEditBox) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim strSortCode As String
    
    bRet = True
    
    strSortCode = g_clsFormProcessing.GetText(txtSortCode)
    If Len(strSortCode) > 0 Then
        If Mid(strSortCode, 3, 1) <> "-" Or Mid(strSortCode, 6, 1) <> "-" Then
            'EP2_517 message spelling error
            MsgBox "Bank Sort Code must be input in the format 99-99-99", vbCritical
            g_clsFormProcessing.SetControlFocus txtSortCode
            bRet = False
        End If
    End If
    
    ValidateSortCode = bRet
    Exit Function
Failed:
    MsgBox "ValidateSortCode: Error is " + Err.DESCRIPTION
    ValidateSortCode = False

End Function

