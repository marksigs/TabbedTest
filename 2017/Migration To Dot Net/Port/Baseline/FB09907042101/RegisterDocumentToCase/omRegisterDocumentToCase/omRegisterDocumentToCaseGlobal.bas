Attribute VB_Name = "omRegisterDocumentToCaseGlobal"
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     24/10/2005  created for Project MARS (MAR232)

Option Explicit

Public gstrComponentId As String
Public gstrComponentResponse As String

Public Sub xmlParseError(ByVal objParseError As IXMLDOMParseError)
    
    Dim strErrDesc As String    ' formatted parser error
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str$(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str$(objParseError.Line) & vbCr & _
        "Character: " & Str$(objParseError.linepos) & vbCr & _
        "Source text: " & objParseError.srcText
        
        Err.Raise oeXMLParserError, "xmlParseError", strErrDesc
            
End Sub

Public Sub CheckError(ByVal vstrMethodName As String)
    If Err.Number = 0 Then
        Exit Sub
    End If
    If Err.Source <> vstrMethodName Then
        If InStr(Err.Source, "." & vstrMethodName) = 0 Then
            If Err.Source = App.EXEName Then
                Err.Source = vstrMethodName
            Else
                Err.Source = vstrMethodName & "." & Err.Source
            End If
        End If
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function FormatError(ByVal vstrXmlIn As String) As String
    
    Dim xmlErrDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlErrDoc = New DOMDocument40
    xmlErrDoc.setProperty "NewParser", True
    xmlErrDoc.async = False
    
    Set xmlElem = xmlErrDoc.createElement("RESPONSE")
    
    If Err.Source = "om4Wrapper" Then
        xmlElem.setAttribute "TYPE", "APPERR"
    Else
        xmlElem.setAttribute "TYPE", "SYSERR"
    End If
    
    Set xmlNode = xmlErrDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlErrDoc.createElement("ERROR")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlErrDoc.createElement("NUMBER")
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("SOURCE")
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("DESCRIPTION")
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("REQUEST")
    xmlElem.Text = vstrXmlIn
    xmlNode.appendChild xmlElem
    
    If Len(gstrComponentResponse) <> 0 Then
        Set xmlElem = xmlErrDoc.createElement("COMPONENT_ID")
        xmlElem.Text = gstrComponentId
        xmlNode.appendChild xmlElem
        Set xmlElem = xmlErrDoc.createElement("COMPONENT_RESPONSE")
        xmlElem.Text = gstrComponentResponse
        xmlNode.appendChild xmlElem
    End If
    
    FormatError = xmlErrDoc.xml
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlErrDoc = Nothing

End Function


