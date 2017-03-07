VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "MapRulesAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      C:\Projects\dev\Omiga4ToOmiga3Download\om4to3\CMapRulesAssist.cls
' Copyright:    Copyright © 1999 Marlborough Stirling
'
'Description:
'   Used by the Data Mapper to implement 'generic' rules for
'   processing field data.
'    All rules have a standard parameter list. Each rule will interpret
'    the parameters according to their local requirements.
'
' Dependencies:
'
' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MH    23/02/00 Created
'MC    29/09/00 Add Subtract to Calculate Rule
'MC    02/10/00 SYS1599 - Add greater than comparison to IF rule.
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

Private m_objErrAssist As New ErrAssist
Private m_colComboValues As ComboValues
Private m_colCounters As Collection
Private m_objContext As ObjectContext


Public Sub Create(ByVal vconDB As Connection)
' description:
'    Creates the class and creates the lookup tables required by Lookup,
'    OM3Lookup, LookupValue, LookupID and LookupDescription.

' pass:
'   vconDB                 The source of lookup values'
'
    Const strMethodName = "Create"
    On Error GoTo CreateVBErr
    
    Set m_colComboValues = New ComboValues
    m_colComboValues.Create vconDB
    
    Set m_colCounters = New Collection
    Exit Sub
CreateVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function EvaluateRule(ByVal vstrRuleName As Variant, _
                            ByVal vvarValue As Variant, _
                            ByVal vcolDataCache As DataCache) As String
    Const strMethodName As String = "EvaluateRule"
    On Error GoTo EvaluateRuleVBErr
    Dim strRule As String
    Dim lngParam As Long
    Dim arrParam() As String
    Dim lngHigh As Long
    Dim lngLow As Long
    Dim lngIdx As Long
    
    ' See if rule parameters have been  given
    strRule = Trim$(vstrRuleName)
    lngParam = InStr(strRule, "(")
    
    'If parameters are given then process them, otherwise use a null array
    If lngParam > 0 And Right$(strRule, 1) = ")" Then
        arrParam = Split(Mid$(strRule, lngParam + 1, Len(strRule) - lngParam - 1), ",")
        strRule = Left$(strRule, lngParam - 1)
        
        'Once the parameters are identified the data can be substituted safely
        lngHigh = UBound(arrParam)
        lngLow = LBound(arrParam)
        lngIdx = 0
        Do While lngIdx <= lngHigh
            arrParam(lngIdx) = vcolDataCache.ReplaceFieldsWithData(arrParam(lngIdx))
            lngIdx = lngIdx + 1
        Loop
        
    Else
        arrParam = Split(vbNullString)
    End If
    
    Select Case UCase$(strRule)
    Case "AGE"
        EvaluateRule = RuleAge(arrParam, vvarValue)
    Case "BOOLEAN"
        EvaluateRule = RuleBoolean(arrParam, vvarValue)
    Case "CALCULATE"
        EvaluateRule = RuleCalculate(arrParam, vvarValue)
    Case "COUNT"
        EvaluateRule = RuleCount(arrParam, vvarValue, vcolDataCache)
    Case "COUNTER"
        EvaluateRule = RuleCounter(arrParam, vvarValue)
    Case "CONCATENATE"
        EvaluateRule = RuleConcatenate(arrParam, vvarValue)
    Case "HEADSTRING"
        EvaluateRule = RuleHeadstring(arrParam, vvarValue)
    Case "IF"
        EvaluateRule = RuleIF(arrParam, vvarValue)
    Case "ISNULL"
        EvaluateRule = RuleIsNull(arrParam, vvarValue)
    Case "OM3LOOKUP"
        EvaluateRule = RuleOM3Lookup(arrParam, vvarValue)
    Case "LOOKUPTEXT"
        EvaluateRule = RuleLookupText(arrParam, vvarValue)
    Case "ROWEXISTS"
        EvaluateRule = RuleRowExists(arrParam, vvarValue, vcolDataCache)
    Case "TODAY"
        EvaluateRule = RuleToday()
    Case "TAILSTRING"
        EvaluateRule = RuleTailstring(arrParam, vvarValue)
    Case "TELEPHONESTD"
        EvaluateRule = RuleTelephoneSTD(arrParam, vvarValue)
    Case "TELEPHONENUMBER"
        EvaluateRule = RuleTelephoneNumber(arrParam, vvarValue)
    Case Else
         Err.Raise omiga4UnspecifiedError, TypeName(Me), "Unknown RULE - """ & vstrRuleName & """"
    End Select

    Exit Function

EvaluateRuleVBErr:
    If Err.Number <> 94 And Err.Number <> omiga4UnspecifiedError Then
        'It is not an "Invalid use of null"
        Err.Description = Err.Description & vbCrLf & "Rule=""" & vstrRuleName & """"
        
        m_objErrAssist.AddToErrSource strMethodName
        
        If m_objErrAssist.IsSystemError = True Then
            m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
        End If
    End If
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
                            

' header ----------------------------------------------------------------------------------
' description: Calculates Age in years. If an offset is given it will be added on
'                Egs: Age,  Age(+1)
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleAge(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleAge"
    On Error GoTo RuleAgeVBErr
    
    Dim lngDiff As Long
        
    CheckParameters rarrParameters(), , 1
    
    If Not IsNull(vvarValue) Then
        If Not Trim$(vvarValue) = vbNullString Then
            lngDiff = DateDiff("yyyy", vvarValue, Now)
            
            If UBound(rarrParameters()) >= 0 Then
                RuleAge = lngDiff + CLng(rarrParameters(0))
            Else
                RuleAge = lngDiff
            End If
        Else
            RuleAge = vbNullString
        End If
    Else
        RuleAge = vbNullString
    End If
    
    Exit Function

RuleAgeVBErr:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Perform simple sums on the source field
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleCalculate(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleCalculate"
    On Error GoTo RuleCalculateVBErr
    
    CheckParameters rarrParameters(), 1, 1
    
    Select Case Left$(rarrParameters(0), 1)
    Case "+"
        RuleCalculate = vvarValue + rarrParameters(0)
    Case "-"
        RuleCalculate = vvarValue + rarrParameters(0)
    Case Else
        Err.Raise oeNotImplemented, TypeName(Me), "Evaluation rule is too complicated"
    End Select
    
    Exit Function

RuleCalculateVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Concatenate a set of fields together
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleConcatenate(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleConcatenate"
    On Error GoTo RuleConcatenateVBErr
    Dim lngLow As Long
    Dim lngHigh As Long
    Dim lngIndex As Long
    Dim strOut As String

    
    lngLow = LBound(rarrParameters)
    lngHigh = UBound(rarrParameters)
    strOut = vbNullString
    
    For lngIndex = lngLow To lngHigh
        strOut = strOut & rarrParameters(lngIndex)
    Next lngIndex
    
    RuleConcatenate = strOut
    
    Exit Function

RuleConcatenateVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Returns the number of rows in the table given in the parameter
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleCount(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant, ByVal vcolDataCache As DataCache) As String
    Const strMethodName As String = "RuleCount"
    On Error GoTo RuleCountVBErr
                            
    CheckParameters rarrParameters(), 1, 1
    
    RuleCount = vcolDataCache.Item(rarrParameters(0)).Count

    Exit Function

RuleCountVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Provides a sequence number / counter for when one cannot be used on the database
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleCounter(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleCounter"
    On Error GoTo RuleCounterVBErr
    
    Dim objCounter As Counter
    Dim objContext As ObjectContext
    
    CheckParameters rarrParameters(), 1, 1
    
    On Error Resume Next
    
    Set objCounter = m_colCounters.Item(rarrParameters(0))
    
    If Not Err.Number = 0 Then
        On Error GoTo RuleCounterVBErr
        If m_objContext Is Nothing Then
            Set objCounter = New Counter
        Else
            Set objCounter = m_objContext.CreateInstance(App.Title & ".Counter")
        End If
        objCounter.Name = rarrParameters(0)
        m_colCounters.Add objCounter, objCounter.Name
    End If
    
    RuleCounter = objCounter.NextValue
    
    Exit Function

RuleCounterVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description:  Returns True if the field is null
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleIsNull(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleIsNull"
    On Error GoTo RuleIsNullVBErr
    
    CheckParameters rarrParameters(), 2, 2
    
    If IsNull(vvarValue) Then
        RuleIsNull = rarrParameters(0)
    Else
        If vvarValue = vbNullString Then
            RuleIsNull = rarrParameters(0)
        Else
            RuleIsNull = rarrParameters(1)
        End If
    End If
    
    Exit Function

RuleIsNullVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description:  Translates a boolean for humans
'               Boolean(YesValue,NoValue,DefaultValue)
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleBoolean(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleBoolean"
    On Error GoTo RuleBooleanVBErr
    
    Const intYes = 0
    Const intNo = 1
    Const intPass = 2
    
    CheckParameters rarrParameters(), 3, 3
    
    If IsNull(vvarValue) Then
        RuleBoolean = rarrParameters(intPass)
    Else
        Select Case vvarValue
        Case True
            RuleBoolean = rarrParameters(intYes)
        Case False
            RuleBoolean = rarrParameters(intNo)
        Case 1, -1
            RuleBoolean = rarrParameters(intYes)
        Case 0
            RuleBoolean = rarrParameters(intNo)
        Case Else
            RuleBoolean = rarrParameters(intPass)
        End Select
    End If
    
    Exit Function

RuleBooleanVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Returns the headstring based on a ' ' delim.
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleHeadstring(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleHeadstring"
    On Error GoTo RuleHeadstringVBErr
    Dim strValue As String
       
    CheckParameters rarrParameters(), 0, 0
    
    strValue = vvarValue & " "
    
    RuleHeadstring = Left$(strValue, InStr(strValue, " ") - 1)
    
    Exit Function

RuleHeadstringVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description:  Translates a value based on the name of a lookup table
'    If an OM3ValueID exists it will be used. If it does not exist then the Omiga4 value will
'    be used since in this case it is the same as the om3 value.
'
'    Null values lookup will be attempted.
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleOM3Lookup(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    'for now return the basic value
    Dim objComboValue As ComboValue
        
    On Error GoTo RuleOM3LookupVBErr1
    CheckParameters rarrParameters(), 1, 1
    
   On Error GoTo RuleOM3LookupVBErr2

    If Not IsNull(vvarValue) Then
        If Not vvarValue = vbNullString Then
            Set objComboValue = m_colComboValues.GetByValueID(rarrParameters(0), vvarValue)
            If objComboValue Is Nothing Then
               RuleOM3Lookup = vbNullString
            Else
               If objComboValue.OM3ValueID = vbNullString Then
                    RuleOM3Lookup = vvarValue
               Else
                    RuleOM3Lookup = objComboValue.OM3ValueID
                End If
            End If
        Else
            RuleOM3Lookup = vbNullString
        End If
    Else
        RuleOM3Lookup = vbNullString
    End If
    
    'tot = tot + 1
    Exit Function
        
RuleOM3LookupVBErr1:
        Err.Raise Err.Number, TypeName(Me), Err.Description

RuleOM3LookupVBErr2:
    #If IGNORECOMBOLOOKUPERRORS = 0 Then
        Err.Raise Err.Number, TypeName(Me), Err.Description & vbCrLf & _
            "OM3Lookup failed for " & rarrParameters(0) & "(" & vvarValue & ")" & vbCrLf & _
             m_colComboValues.GroupInfo(rarrParameters(0))
    #End If
End Function

' header ----------------------------------------------------------------------------------
' description:  Translates a value based on the name of a lookup table
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleLookupText(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleLookupText"
    'for now return the basic value
    
    On Error GoTo RuleLookupTextVBErr
    
        CheckParameters rarrParameters(), 1, 1
    
        If Not IsNull(vvarValue) Then
            If Not vvarValue = vbNullString Then
                RuleLookupText = m_colComboValues.GetByValueID(rarrParameters(0), vvarValue).ValueName
            Else
                RuleLookupText = vbNullString
            End If
        Else
            RuleLookupText = vbNullString
        End If
        
        Exit Function
        
RuleLookupTextVBErr:
    #If IGNORECOMBOLOOKUPERRORS = 0 Then
        If Err.Number <> oeInvalidParameter Then _
            Err.Raise Err.Number, TypeName(Me), Err.Description & vbCrLf & "Lookup failed for " & rarrParameters(0) & "(" & vvarValue & ")"
    #End If

End Function

' header ----------------------------------------------------------------------------------
' description: Returns true if a table contains data
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleRowExists(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant, _
                            ByVal vcolDataCache As DataCache) As String
    Const strMethodName As String = "RuleRowExists"
    On Error GoTo RuleRowExistsVBErr
       
    CheckParameters rarrParameters(), 3, 3
    
    If vcolDataCache.Item(rarrParameters(0)).HasData Then
        'Return the True value
        RuleRowExists = rarrParameters(1)
    Else
        'Return the False value
        RuleRowExists = rarrParameters(2)
    End If
    
    Exit Function

RuleRowExistsVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description:
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleTailstring(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleTailstring"
    On Error GoTo RuleTailstringVBErr
    Dim lngSpace As String
    Dim strValue As String
    
    CheckParameters rarrParameters(), 0, 0
    
    strValue = vvarValue
    lngSpace = InStr(strValue, " ")
    
    If lngSpace > 0 Then
        RuleTailstring = Mid$(strValue, lngSpace + 1)
    Else
        RuleTailstring = vbNullString
    End If
    
    Exit Function

RuleTailstringVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Returns the STD part of a telephone number assuming that it is trailed
'               by a single white space.
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleTelephoneSTD(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleTelephoneSTD"
    On Error GoTo RuleTelephoneSTDVBErr
    'Only keep the first 6 digits
    RuleTelephoneSTD = Left$(RuleHeadstring(rarrParameters, vvarValue), 6)
    
    Exit Function

RuleTelephoneSTDVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Return the local part of the telephone or ignore the 1st 6 digits
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleTelephoneNumber(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleTelephoneNumber"
    On Error GoTo RuleTelephoneNumberVBErr
    Dim lngSpace As String
    Dim strValue As String
    
    strValue = vvarValue
    lngSpace = InStr(strValue, " ")
    
    If lngSpace > 7 Then
        RuleTelephoneNumber = Mid$(strValue, 7)
    ElseIf lngSpace > 0 Then
        RuleTelephoneNumber = Mid$(strValue, lngSpace + 1)
    Else
        RuleTelephoneNumber = Mid$(strValue, 7)
    End If

    Exit Function

RuleTelephoneNumberVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

' header ----------------------------------------------------------------------------------
' description: Returns the current date/time
'
' return:     The current date teim
'
'------------------------------------------------------------------------------------------
Private Function RuleToday() As Date
    RuleToday = Now
End Function



' header ----------------------------------------------------------------------------------
' description: Performs If-then-else processing
'
' pass:         rarrParameters  Parameters used by this rule (if any)
'               vvarValue       The value that is being processed
'
' return:                       Variant (usually a string), the result of the manipulation
'
'------------------------------------------------------------------------------------------
Private Function RuleIF(ByRef rarrParameters() As String, _
                            ByVal vvarValue As Variant) As String
    Const strMethodName As String = "RuleIF"
    On Error GoTo RuleIFVBErr
    Dim lngEquals As Long
    Dim lngGt As Long
    Dim strExpression As String
    
    CheckParameters rarrParameters(), 3, 3
    
    strExpression = rarrParameters(0)
    
    lngEquals = InStr(strExpression, "=")
    lngGt = InStr(strExpression, ">")
    
    If lngEquals > 0 Then
        If (Left$(strExpression, lngEquals - 1) = Mid$(strExpression, lngEquals + 1)) Then
            RuleIF = rarrParameters(1)
        Else
            RuleIF = rarrParameters(2)
        End If
    ElseIf lngGt > 0 Then
        If IsNumeric(Left$(strExpression, lngGt - 1)) And IsNumeric(Mid$(strExpression, lngGt + 1)) Then
            If (Val(Left$(strExpression, lngGt - 1)) > Val(Mid$(strExpression, lngGt + 1))) Then
                RuleIF = rarrParameters(1)
            Else
                RuleIF = rarrParameters(2)
            End If
        Else
            m_objErrAssist.RaiseError TypeName(Me), strMethodName, oeNotImplemented, "IF expression """ & strExpression & """ is too complicated. Non numeric greater than comparison."
        End If
    Else
        m_objErrAssist.RaiseError TypeName(Me), strMethodName, oeNotImplemented, "IF expression """ & strExpression & """ is too complicated"
    End If
    

    Exit Function

RuleIFVBErr:
          
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub CheckParameters(ByRef rarrParameters() As String, Optional ByVal intMin As Integer = 0, Optional ByVal intMax As Integer = 32767)
    Dim intTot As Integer
    
    intTot = UBound(rarrParameters) - LBound(rarrParameters) + 1
    If intTot < intMin Then
        Err.Raise oeInvalidParameter, TypeName(Me), "Too few parameters, expected at least " & CStr(intMin)
    End If
    
    If intTot > intMax Then
        Err.Raise oeInvalidParameter, TypeName(Me), "Too many parameters, expected up to " & CStr(intMax)
    End If
    
End Sub

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_colComboValues = Nothing
    Set m_colCounters = Nothing
End Sub
