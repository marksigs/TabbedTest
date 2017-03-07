Attribute VB_Name = "xmlFunctions"
Option Explicit

'------------------------------------------------------------------------
' Procedure: xmlGetMandatoryNode
' Author: MO
' Date: 27/07/2002
' Purpose: Attempts to retreive a node from its parent and raises an error
'           if it cant be found
' Input parameters: xmlParentNode, strXPath
' Output parameters: Node
'------------------------------------------------------------------------
Public Function xmlGetMandatoryNode(xmlParentNode As IXMLDOMNode, strXPath As String) As IXMLDOMNode
    
    On Error GoTo xmlGetMandatoryNodeErrorOccured
    
    Dim objNode As IXMLDOMNode
    
    Set objNode = xmlParentNode.selectSingleNode(strXPath)
    
    If objNode Is Nothing Then
        Err.Raise direMandatoryXMLNodeMissing, "xmlGetMandatoryNode", "Missing Mandatory Node, xmlParentNode.nodeName = " & xmlParentNode.nodeName & " : " & " strXPath = " & strXPath
    End If
    
    Set xmlGetMandatoryNode = objNode
    
xmlGetMandatoryNodeExit:
    
    Set objNode = Nothing
    
Exit Function

xmlGetMandatoryNodeErrorOccured:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume xmlGetMandatoryNodeExit

End Function

'------------------------------------------------------------------------
' Procedure: xmlGetMandatoryAttribute
' Author: MO
' Date: 27/07/2002
' Purpose: Attempts to retreive an attribute from its parent and raises an error
'           if it cant be found
' Input parameters: xmlParentNode, strAttributeName
' Output parameters: Node
'------------------------------------------------------------------------
Public Function xmlGetMandatoryAttribute(xmlParentNode As IXMLDOMNode, strAttributeName As String) As IXMLDOMNode
    
    On Error GoTo xmlGetMandatoryAttributeErrorOccured
    
    Dim objAttribute As IXMLDOMNode
    
    Set objAttribute = xmlParentNode.Attributes.getNamedItem(strAttributeName)
    
    If objAttribute Is Nothing Then
        Err.Raise direMandatoryXMLAttributeMissing, "xmlGetMandatoryAttribute", "Missing Mandatory Attribute, xmlParentNode.nodeName = " & xmlParentNode.nodeName & " : " & " strAttributeName = " & strAttributeName
    End If
    
    Set xmlGetMandatoryAttribute = objAttribute
    
xmlGetMandatoryAttributeExit:
    
    Set objAttribute = Nothing
    
Exit Function

xmlGetMandatoryAttributeErrorOccured:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume xmlGetMandatoryAttributeExit

End Function

'------------------------------------------------------------------------
' Procedure: xmlSetAttributeValue
' Author: MO
' Date: 14/08/2002
' Purpose: Creates a xml attribute on vxmlDestNode will overwrite existing attribute value
' Input parameters: vxmlDestNode, vstrAttribName, vstrAttribValue
' Output parameters: None
'------------------------------------------------------------------------
Public Sub xmlSetAttributeValue(ByVal vxmlDestNode As IXMLDOMNode, ByVal vstrAttribName As String, ByVal vstrAttribValue As String)

On Error GoTo xmlSetAttributeValueErrorOccured

    Dim xmlAttrib As IXMLDOMAttribute
    
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName)
    xmlAttrib.Value = vstrAttribValue
        
    vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)

xmlSetAttributeValueExit:
    
    Set xmlAttrib = Nothing
    
Exit Sub

xmlSetAttributeValueErrorOccured:
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume xmlSetAttributeValueExit

End Sub

'------------------------------------------------------------------------
' Procedure: xmlLoad
' Author: MO
' Date: 29/08/2002
' Purpose: Loads a string into a XML DOM and returns relevants error
' Input parameters: strXMLData
' Output parameters: FreeThreadedDOMDocument40
'------------------------------------------------------------------------
Public Function xmlLoad(ByVal vstrXMLData As String) As FreeThreadedDOMDocument40
    
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    
    objXmlDoc.async = False         ' wait until XML is fully loaded
    
    objXmlDoc.loadXML vstrXMLData
    
    If objXmlDoc.parseError.errorCode <> 0 Then
        With objXmlDoc.parseError
            Err.Raise direFailedToLoadXML, "xmlLoad", "Failed to load XML data: " & vbCrLf & "errorcode - " & .errorCode & vbCrLf & "reason - " & .reason & vbCrLf & "scrtext - " & .srcText
        End With
    End If
    
    Set xmlLoad = objXmlDoc
    
End Function


