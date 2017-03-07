Attribute VB_Name = "CheckDigit"
'Workfile:      CheckDigit.bas
'Copyright:     Copyright © 2006 Vertex Financial Services
'
'Description:   Check Digit utilities
'Dependencies:  none
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     17/05/2006  EP561 add IsCheckDigitRequired
'------------------------------------------------------------------------------------------

Option Explicit

Public Function GenerateCheckDigit(ByVal vstrNumber As String) As String
' header ----------------------------------------------------------------------------------
' description:  Generates the Check Digit for the number that is passed in
' pass:
'   vstrNumber  Number for which the check digit is required
' return:       Check Digit
'------------------------------------------------------------------------------------------
On Error GoTo GenerateCheckDigitError

    Const strFunctionName As String = "GenerateCheckDigit"
    
    Dim objErrAssist As New ErrAssist
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim intTotal      As Integer
    Dim intIndex      As Integer
    Dim intLength     As Integer
    Dim intDigit      As Integer
    Dim intRemainder  As Integer
    Dim intCheckDigit As Integer
    Dim strCheckDigit As String
    
    intLength = Len(vstrNumber)
    
    If intLength <= 11 And IsNumeric(vstrNumber) Then
        
        ' Calculate the total of all digits multiplied by their
        ' respective mutiplier
        For intIndex = intLength To 1 Step -1
             intDigit = CInt(Mid$(vstrNumber, intIndex, 1))
             intTotal = intTotal + (intDigit * (intLength - intIndex + 2))
        Next intIndex
        
        ' Determine the check digit. A valid check digit is anything
        ' other than 10.
        intRemainder = intTotal Mod 11
        
        If intRemainder = 0 Then
            intCheckDigit = 0
        Else
            intCheckDigit = 11 - intRemainder
        End If
        
        If intCheckDigit <> 10 Then
            strCheckDigit = CStr(intCheckDigit)
        End If
        
        GenerateCheckDigit = strCheckDigit
    Else
        objErrAssist.RaiseError "StdData", strFunctionName, omiga4Err001, "Invalid Number: " & vstrNumber
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

GenerateCheckDigitError:

    App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
           
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'IK_17/05/2006_EP561
Public Function IsCheckDigitRequired() As Boolean

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objCRUD As Object
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("GLOBALPARAMETER")
    xmlElem.setAttribute "NAME", "NoCheckDigitsRequired"
    xmlNode.appendChild xmlElem
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc.loadXML objCRUD.omRequest(xmlDoc.xml)
    Set objCRUD = Nothing
    
    If Not xmlDoc.selectSingleNode("RESPONSE/GLOBALPARAMETER[@BOOLEAN='1']") Is Nothing Then
        IsCheckDigitRequired = False
    Else
        IsCheckDigitRequired = True
    End If
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing

End Function

