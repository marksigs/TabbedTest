Attribute VB_Name = "omCRUDGlobals"
Option Explicit
Public Sub Main()
    ' adoAssist
    adoCRUDBuildDbConnectionString
    adoCRUDLoadSchema
    adoCRUDInitComboFile
End Sub

'------------------------------------------------------------------------
' Procedure DoRequest
' Author:           IK
' Date:             24/02/2004
' Purpose: The main operation request handler for the class
' Input parameters: vxmlRequestNode holding XML request
'                   vxmlResponseNode holding blank response
' Output parameters: vxmlResponseNode which will hold the XML response
'------------------------------------------------------------------------
Public Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    Dim xmlOperationNode As IXMLDOMNode
    
    Dim strOperation As String
    
    If xmlAttributeValueExists(vxmlRequestNode, "CRUD_OP") Then
        ' back door, for simple operations only?
        adoCRUD vxmlRequestNode, vxmlRequestNode, vxmlResponseNode
    
    Else
    
        If vxmlRequestNode.nodeName = "REQUEST" Then
            If xmlAttributeValueExists(vxmlRequestNode, "OPERATION") Then
                strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
            End If
        Else
            If xmlAttributeValueExists(vxmlRequestNode, "NAME") Then
                strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
            End If
        End If
        
        If Len(strOperation) > 0 Then
        
            strOperation = UCase(strOperation)
            
            ' do request data validation
            validateRequestData strOperation, vxmlRequestNode, vxmlResponseNode
            
            ' do we have a CRUD OPERATION node
            Set xmlOperationNode = adoCRUDGetOperationNode(strOperation)
            
        Else
            
            errThrowError _
                "OmRequest", _
                oeNotImplemented, _
                "invalid REQUEST, no CRUD_OP or OPERATION - " & vxmlRequestNode.xml
            
        End If
        
        If Not xmlOperationNode Is Nothing Then
        
            If xmlAttributeValueExists(xmlOperationNode, "CRUD_OP") Then
                
                adoCRUD xmlOperationNode, vxmlRequestNode, vxmlResponseNode
                
            Else
            
                ' possible enhancement:
                ' replacement object/method call specified in OPERATION element
                ' now:
                ' raise error
                errThrowError _
                    "adoCRUD", _
                    oeXMLMissingAttribute, _
                    "OM_SCHEMA OPERATION missing CRUD_OP: " & xmlOperationNode.xml
                
            End If
        
        Else
        
            ' forward to private (named) function
        
            Select Case strOperation
            
                Case Else
                    errThrowError _
                        "OmRequest", _
                        oeNotImplemented, _
                        vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
            
            End Select
            
        End If
        
    End If
    
DoRequestExit:
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Sub

Public Sub validateRequestData( _
    ByVal vstrOperation As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
   
    ' is operation valid for this object?
    Select Case vstrOperation
    
        ' add Case for each operation to be supported by this module,
        ' e.g.
        
        Case Else
            errThrowError _
                "OmRequest", _
                oeNotImplemented, _
                vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
                
    End Select
               
End Sub

Public Function PreProcInterface( _
    ByVal vxmlRequestDoc As DOMDocument40) _
    As IXMLDOMNode

    Dim objPreProcComponent As Object

    Dim xmlRequestNode As IXMLDOMNode
    
    Dim strRef As String, _
        strProgId As String
    
    Set xmlRequestNode = vxmlRequestDoc.selectSingleNode("REQUEST")
    
    If xmlAttributeValueExists(xmlRequestNode, "preProcRef") Then
        
        strRef = UCase(xmlGetAttributeText(xmlRequestNode, "preProcRef"))
        
        If Right(strRef, 5) = ".XSLT" Or Right(strRef, 4) = ".XSL" Then
        
            xmlTransform vxmlRequestDoc, strRef
            
        Else
    
            If xmlAttributeValueExists(xmlRequestNode, "preProcProgId") Then
            
                strProgId = xmlGetAttributeText(xmlRequestNode, "preProcProgId")
            
                Set objPreProcComponent = GetObjectContext.CreateInstance(strProgId)
                
                If objPreProcComponent Is Nothing Then
                
                    App.LogEvent "missing pre-proc component: " & strProgId, vbLogEventTypeError
                
                End If
                        
            Else
            
                Set objPreProcComponent = GetObjectContext.CreateInstance("omCRUD.omCRUDPreProc")
            
            End If
            
            If Not objPreProcComponent Is Nothing Then
            
                vxmlRequestDoc.loadXML objPreProcComponent.OmRequest(vxmlRequestDoc.xml)
                
            End If
            
        End If
    
    End If
            
    Set PreProcInterface = vxmlRequestDoc.selectSingleNode("REQUEST")

End Function

Public Function PostProcInterface( _
    ByVal vxmlRequestDoc As DOMDocument40, _
    ByVal vxmlResponseDoc As DOMDocument40) _
    As IXMLDOMNode

    Dim objPostProcComponent As Object

    Dim xmlRequestNode As IXMLDOMNode
    
    Dim strRef As String, _
        strProgId As String
    
    Set xmlRequestNode = vxmlRequestDoc.selectSingleNode("REQUEST")
    
    If xmlAttributeValueExists(xmlRequestNode, "postProcRef") Then
        
        strRef = UCase(xmlGetAttributeText(xmlRequestNode, "postProcRef"))
        
        If Right(strRef, 5) = ".XSLT" Or Right(strRef, 4) = ".XSL" Then
        
            xmlTransform vxmlResponseDoc, strRef
            
        Else
    
            If xmlAttributeValueExists(xmlRequestNode, "postProcProgId") Then
            
                strProgId = xmlGetAttributeText(xmlRequestNode, "postProcProgId")
            
                Set objPostProcComponent = GetObjectContext.CreateInstance(strProgId)
                
                If objPostProcComponent Is Nothing Then
                
                    App.LogEvent "missing post-proc component: " & strProgId, vbLogEventTypeError
                
                End If
                        
            Else
            
                Set objPostProcComponent = GetObjectContext.CreateInstance("omCRUD.omCRUDPostProc")
            
            End If
            
            If Not objPostProcComponent Is Nothing Then
            
                vxmlResponseDoc.loadXML _
                    objPostProcComponent.OmRequest(vxmlRequestDoc.xml, vxmlResponseDoc.xml)
                
            End If
            
        End If
    
    End If
            
    Set PostProcInterface = vxmlResponseDoc.selectSingleNode("RESPONSE")

End Function

Private Sub xmlTransform(ByVal xmlDoc As DOMDocument40, ByVal vstrFileName)

    Dim xslDoc As DOMDocument40
    
    Dim strFileSpec As String
    
    strFileSpec = App.Path & "\" & vstrFileName
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set xslDoc = New DOMDocument40
    xslDoc.async = False
    xslDoc.Load strFileSpec
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeUnspecifiedError, _
            "omCRUD", _
                "error loading xslt document: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If
    
    xmlDoc.loadXML xmlDoc.transformNode(xslDoc)
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeUnspecifiedError, _
            "omCRUD", _
                "error in transformation: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If

End Sub

Public Sub NamedInterface( _
    ByVal vxmlOperationNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    Dim objNamedComponent As Object
    
    Dim xmlResponseDoc As DOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim strProgId As String, _
        strMethod As String, _
        strResponse As String, _
        strErrText As String
        
    On Error GoTo NamedInterfaceExit:
    
    strMethod = xmlGetAttributeText(vxmlOperationNode, "MethodName")
    strErrText = "no MethodName specified: " & vxmlOperationNode.xml
    If vxmlOperationNode.Attributes.getNamedItem("MethodName") Is Nothing Then
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
    
    strErrText = "no valid REQUEST for method: " & strProgId & "." & strMethod & " - " & vxmlOperationNode.xml
    If vxmlOperationNode.selectSingleNode("REQUEST") Is Nothing Then
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
                
    strProgId = xmlGetAttributeText(vxmlOperationNode, "ProgId")
    strErrText = "unable to create component: " & strProgId
    Set objNamedComponent = GetObjectContext.CreateInstance(strProgId)
    If objNamedComponent Is Nothing Then
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
    
    strErrText = "error calling method: " & strProgId & "." & strMethod
    strResponse = CallByName(objNamedComponent, strMethod, VbMethod, vxmlOperationNode.firstChild.xml)
    Set objNamedComponent = Nothing

    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    
    strErrText = "invalid RESPONSE calling method: " & strProgId & "." & strMethod & " - " & strResponse
    xmlResponseDoc.loadXML strResponse
    
    If xmlResponseDoc.parseError.errorCode <> 0 Then
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
    
    Set xmlThisResponseNode = xmlResponseDoc.selectSingleNode("RESPONSE")
    If xmlThisResponseNode Is Nothing Then
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
    
    If xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        vxmlResponseNode.ownerDocument.loadXML strResponse
        Err.Raise oeUnspecifiedError, "NamedInterface", strErrText
    End If
    
    For Each xmlAttrib In xmlThisResponseNode.Attributes
        vxmlResponseNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    Next
    
    For Each xmlNode In xmlThisResponseNode.childNodes
        vxmlResponseNode.appendChild xmlNode.cloneNode(True)
    Next

NamedInterfaceExit:
    
    Set objNamedComponent = Nothing

    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlThisResponseNode = Nothing
    Set xmlResponseDoc = Nothing
    
    If Err.Number <> 0 Then
        If Err.Number = oeUnspecifiedError Then
            Err.Raise Err.Number, Err.Source, strErrText
        Else
            Err.Raise Err.Number, "NamedInterface " & Err.Source, strErrText & ", " & Err.Description
        End If
    End If
    
End Sub

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
        If IsUpperAlphaChar(strChar) Then
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
        If IsLowerAlphaChar(strChar) Then
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
        If IsDigits(strChar) Then
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
        If IsAlpha(strChar) = False Then
            If IsDigits(strChar) = False Then
                blnContainsSpecialChars = True
            End If
        End If
        intIndex = intIndex + 1
    Wend
    ContainsSpecialChars = blnContainsSpecialChars

End Function


