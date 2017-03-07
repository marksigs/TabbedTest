Attribute VB_Name = "xmlAssistEx"
'header ----------------------------------------------------------------------------------
'Workfile:      xmlAssist.bas
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Helper functions for XML handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog    Date        Description
'IK      01/11/00    Initial creation
'ASm     15/01/01    SYS1824: Rationalisation of methods following meeting with PC and IK
'LD      30/01/01    Add optional default value to xmlGetAttribute*
'PSC     22/02/01    Add xmlGetRequestNode
'SR      05/03/01    New methods 'xmlSetSysDateToNodeAttrib' and 'xmlSetSysDateToNodeListAttribs'
'PSC     10/04/01    Add extra parameter to the Check/Get mandatory attribut/node methods to
'                    enable the default message to be overridden
'SR      13/06/01    SYS2362 - modified method 'xmlMakeNodeElementBased'. Create attrib for
'                    combo descriptions
'AL      17/01/02    Add variable strParserError when calling errThrowError in xmlParserError function
'SF      08/02/02    FormatParserError() has been declared as Public.
'                    IXMLDOMParseError.url has been added to the FormatParserError() output
'------------------------------------------------------------------------------------------
Option Explicit
Private Const NO_MESSAGE_NUMBER As Long = -1
Public Sub xmlSetSysDateToNodeAttrib(ByVal vxmlNode As IXMLDOMNode, _
                                   ByVal strDateAttribName As String, _
                                   Optional blnOverWrite As Boolean = False)
'----------------------------------------------------------------------------------
'Description :
'Set system date to the attrib. Overwrite the existing value based on the optional
'parameter passed in
'Pass:
'vxmlNode           : Node to which the required attrib is attached
'strDateAttribName  : Name of the attribute to which the value is to be set
'blnOverWrite       : if Yes, OverWrite the existing value(if any), else no.
'----------------------------------------------------------------------------------
    If Len(xmlGetAttributeText(vxmlNode, strDateAttribName)) > 0 Then
        If blnOverWrite Then
            xmlSetAttributeValue vxmlNode, strDateAttribName, Format(Now, "dd/mm/yyyy hh:mm:ss")
        End If
    Else
        xmlSetAttributeValue vxmlNode, strDateAttribName, Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If
End Sub
Public Sub xmlSetSysDateToNodeListAttribs(ByVal vxmlNodeList As IXMLDOMNodeList, _
                                        ByVal strDateAttribName As String, _
                                        Optional blnOverWrite As Boolean = False)
'----------------------------------------------------------------------------------
'Description :
'Set system date to the attrib. Overwrite the existing value based on the optional
'parameter passed in
'Pass:
'vxmlNode           : Node to which the required attrib is attached
'strDateAttribName  : Name of the attribute to which the value is to be set
'blnOverWrite       : if Yes, OverWrite the existing value(if any), else no.
'----------------------------------------------------------------------------------
   
    Dim lngNodeCount As Long
    For lngNodeCount = 0 To vxmlNodeList.length - 1
        xmlSetSysDateToNodeAttrib vxmlNodeList.Item(lngNodeCount), strDateAttribName, False
    Next lngNodeCount
End Sub
Public Sub xmlParserError(ByVal vobjParseError As IXMLDOMParseError, ByVal vstrCallingFunction As String)
' header ----------------------------------------------------------------------------------
' description:  Raise XML parser error
' pass:         vobjParseError parser error
'               vstrCallingFunction name of the original calling function
' return:       n/a
'------------------------------------------------------------------------------------------
    
    Dim strParserError As String
    strParserError = FormatParserError(vobjParseError)
    errThrowError vstrCallingFunction, oeXMLParserError, strParserError
        
End Sub
Public Function FormatParserError(ByVal objParseError As IXMLDOMParseError) As String
' header ----------------------------------------------------------------------------------
' description:  Format parser error into user friendly error information
' pass:         vobjParseError parser error
' return:       Parser error converted to a string
'------------------------------------------------------------------------------------------
    Dim strErrDesc As String    ' formatted parser error
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str(objParseError.Line) & vbCr & _
        "Character: " & Str(objParseError.linepos) & vbCr & _
        "Source text: " & objParseError.srcText & vbCr & _
        "url: " & objParseError.url
    FormatParserError = strErrDesc
End Function
Public Function xmlLoad( _
    ByVal vstrXMLData As String, _
    ByVal vstrCallingFunction As String) _
    As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:  Creates and loads an XML Document from a string
' pass:         vstrXMLData as xml data stream to be loaded
'               vstrCallingFunction which is the name of the calling function
' return:       XML DOM Document of the vstrXMLData string
'------------------------------------------------------------------------------------------
    
    Dim strFunctionName As String
    strFunctionName = "xmlLoad"
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.async = False         ' wait until XML is fully loaded
    objXmlDoc.setProperty "NewParser", True
    objXmlDoc.validateOnParse = False
    objXmlDoc.loadXML vstrXMLData
    If objXmlDoc.parseError.errorCode <> 0 Then
        xmlParserError objXmlDoc.parseError, vstrCallingFunction
    End If
    Set xmlLoad = objXmlDoc
End Function
Private Function GetNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vblnMandatoryNode As Boolean = False, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vblnMandatoryNode   Optional -
'                                   if true, an error is raised if the pattern is not met.
'                                   If false, nothing is returned.
' return:       IXMLDOMNode         The node which matches the search pattern
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "xmlGetNode"
    Dim xmlNode As IXMLDOMNode
    If vxmlParentNode Is Nothing Then
        errThrowError cstrFunctionName, oeMissingElement, "Missing parent node"
    End If
    Set xmlNode = vxmlParentNode.selectSingleNode(vstrMatchPattern)
    If vblnMandatoryNode = True And xmlNode Is Nothing Then
        If vlngMessageNo = NO_MESSAGE_NUMBER Then
            errThrowError cstrFunctionName, oeMissingElement, "Match pattern:- " & vstrMatchPattern
        Else
            errThrowError cstrFunctionName, vlngMessageNo
        End If
    End If
    Set GetNode = xmlNode
    Set xmlNode = Nothing
End Function
Public Function xmlGetNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:      Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:             vxmlParentNode      Node to be searched
' return:           The node which matches the search pattern
'                   Nothing if the node is not found
'------------------------------------------------------------------------------------------
    
    Set xmlGetNode = GetNode(vxmlParentNode, vstrMatchPattern)
End Function
Public Function xmlGetNodeText( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As String
' header ----------------------------------------------------------------------------------
' description:      Gets the node text
' pass:             vxmlParentNode      xml Node to search for child node
'                   vstrMatchPattern    XSL pattern match
' return:           Node text of child node
'                   empty string if node not found
'------------------------------------------------------------------------------------------
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = xmlGetNode(vxmlParentNode, vstrMatchPattern)
    If Not xmlNode Is Nothing Then
        xmlGetNodeText = xmlNode.Text
        Set xmlNode = Nothing
    End If
End Function
Public Function xmlGetNodeAsBoolean( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       True if node text = "1" or "Y" or "YES" else False
'               False if node not found
'------------------------------------------------------------------------------------------
    Dim strValue As String
    strValue = xmlGetNodeText(vxmlParentNode, vstrMatchPattern)
    strValue = UCase(strValue)
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetNodeAsBoolean = True
    End If
End Function
Public Function xmlGetNodeAsDate( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As Date
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       The text value of the found node as a Date
'               empty Date if node not found
'------------------------------------------------------------------------------------------
    
    xmlGetNodeAsDate = _
        CSafeDate(xmlGetNodeText(vxmlParentNode, vstrMatchPattern))
End Function
Public Function xmlGetNodeAsDouble( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As Double
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       The text value of the found node as a Double
'               0 if node not found
'------------------------------------------------------------------------------------------
    
    xmlGetNodeAsDouble = _
        CSafeDbl(xmlGetNodeText(vxmlParentNode, vstrMatchPattern))
End Function
Public Function xmlGetNodeAsInteger( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As Integer
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       The text value of the found node as a Integer
'               0 if node not found
'------------------------------------------------------------------------------------------
    
    xmlGetNodeAsInteger = _
        CSafeInt(xmlGetNodeText(vxmlParentNode, vstrMatchPattern))
End Function
Public Function xmlGetNodeAsLong( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As Long
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       The text value of the found node as a Long
'               0 if node not found
'------------------------------------------------------------------------------------------
    
    xmlGetNodeAsLong = _
        CSafeLng(xmlGetNodeText(vxmlParentNode, vstrMatchPattern))
End Function
Public Sub xmlCheckMandatoryNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:  calls GetNode to check that specified node exists
'               GetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
'------------------------------------------------------------------------------------------
    GetNode vxmlParentNode, vstrMatchPattern, True, vlngMessageNo
End Sub
Public Function xmlGetMandatoryNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       IXMLDOMNode         The node which matches the search pattern or
'------------------------------------------------------------------------------------------
    Set xmlGetMandatoryNode = GetNode(vxmlParentNode, vstrMatchPattern, True, vlngMessageNo)
End Function
Public Function xmlGetMandatoryNodeList( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As IXMLDOMNodeList
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       IXMLDOMNode         The node which matches the search pattern or
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName = "xmlGetMandatoryNodeList"
    Set xmlGetMandatoryNodeList = vxmlParentNode.selectNodes(vstrMatchPattern)
    If xmlGetMandatoryNodeList.length = 0 Then
        If vlngMessageNo = NO_MESSAGE_NUMBER Then
            errThrowError cstrFunctionName, oeMissingElement, "Match pattern:- " & vstrMatchPattern
        Else
            errThrowError cstrFunctionName, vlngMessageNo
        End If
    End If
End Function
Public Function xmlGetMandatoryNodeText( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As String
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       String              Node text of child node
'------------------------------------------------------------------------------------------
    Dim xmlNode As IXMLDOMNode
    Dim strValue As String
    Set xmlNode = xmlGetMandatoryNode(vxmlParentNode, vstrMatchPattern, vlngMessageNo)
    strValue = xmlNode.Text
    Set xmlNode = Nothing
    If Len(strValue) = 0 Then
        If vlngMessageNo = NO_MESSAGE_NUMBER Then
            errThrowError _
                "xmlGetMandatoryNodeText", _
                oeMissingElementValue, _
                "Match pattern:- " & vstrMatchPattern
        Else
            errThrowError "xmlGetMandatoryNodeText", vlngMessageNo
        End If
    End If
    xmlGetMandatoryNodeText = strValue
End Function
Public Function xmlGetMandatoryNodeAsBoolean( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       True if node text = "1" or "Y" or "YES" else False
'------------------------------------------------------------------------------------------
    Dim strValue As String
    strValue = xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo)
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetMandatoryNodeAsBoolean = True
    End If
End Function
Public Function xmlGetMandatoryNodeAsDate( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Date
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       The text value of the found node as a Date
'------------------------------------------------------------------------------------------
    
    xmlGetMandatoryNodeAsDate = _
        CSafeDate(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo))
End Function
Public Function xmlGetMandatoryNodeAsDouble( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Double
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       The text value of the found node as a Double
'------------------------------------------------------------------------------------------
    xmlGetMandatoryNodeAsDouble = _
        CSafeDbl(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo))
End Function
Public Function xmlGetMandatoryNodeAsInteger( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Integer
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       The text value of the found node as a Integer
'------------------------------------------------------------------------------------------
    xmlGetMandatoryNodeAsInteger = _
        CSafeInt(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo))
End Function
Public Function xmlGetMandatoryNodeAsLong( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Long
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetNode to check that specified node exists and return node
'               xmlGetNode will raise an error if this condition is not met
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       The text value of the found node as a Long
'------------------------------------------------------------------------------------------
    xmlGetMandatoryNodeAsLong = _
        CSafeLng(xmlGetMandatoryNodeText(vxmlParentNode, vstrMatchPattern, vlngMessageNo))
End Function
Public Function xmlGetAttributeNode( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttributeName As String) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds the attribute node specified by vstrAttributeName
'               in the XML node vxmlNode
' pass:         vxmlNode            Node to be searched
'               vstrAttributeName   Attribute to be found in the node
' return:       IXMLDOMNode         The attribute node
'------------------------------------------------------------------------------------------
    
    Set xmlGetAttributeNode = vxmlNode.Attributes.getNamedItem(vstrAttributeName)
End Function
Public Function xmlAttributeValueExists( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:  Finds the attribute node specified by vstrAttributeName
'               in the XML node vxmlNode
' pass:         vxmlNode            Node to be searched
'               vstrAttributeName   Attribute to be found in the node
' return:       True if attribute (with value) exists on node
'------------------------------------------------------------------------------------------
    
    Dim xmlNode As IXMLDOMNode
    If Not vobjNode Is Nothing Then
        Set xmlNode = xmlGetAttributeNode(vobjNode, vstrAttribName)
        If Not xmlNode Is Nothing Then
            
            If Len(xmlNode.Text) > 0 Then
                xmlAttributeValueExists = True
            End If
        End If
    End If
End Function
Public Function xmlGetAttributeText( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As String
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node text
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
' return:           Node text of attribute
'                   empty string if attribute does not exist on Node
'------------------------------------------------------------------------------------------
    
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = xmlGetAttributeNode(vobjNode, vstrAttribName)
    If Not xmlNode Is Nothing Then
        xmlGetAttributeText = xmlNode.Text
        Set xmlNode = Nothing
    Else
        xmlGetAttributeText = vstrDefault
    End If
End Function
Public Function xmlGetAttributeAsBoolean( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
' return:       True if attribute value is "1" or "Y" or "YES"
'               otherwise False
'------------------------------------------------------------------------------------------
    Dim strValue As String
    strValue = UCase(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault))
    ' AL added strValue = TRUE comparison
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Or strValue = "TRUE" Then
        xmlGetAttributeAsBoolean = True
    End If
End Function
Public Function xmlGetAttributeAsDate( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Date
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node as a date
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
' return:           Date - Node text as a date
'                   empty date if attribute value not found
'------------------------------------------------------------------------------------------
    
    xmlGetAttributeAsDate = CSafeDate(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function xmlGetAttributeAsDouble( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Double
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node as a double
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
' return:           Double - Node text as a double
'                   0 if attribute value not found
'------------------------------------------------------------------------------------------
    
    xmlGetAttributeAsDouble = CSafeDbl(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function xmlGetAttributeAsInteger( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Integer
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node as a integer
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
' return:           Integer - Node text as a integer
'                   0 if attribute value not found
'------------------------------------------------------------------------------------------
    
    xmlGetAttributeAsInteger = CSafeInt(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function xmlGetAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Long
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node as a long
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
' return:           Long - Node text as a long
'                   0 if attribute value not found
'------------------------------------------------------------------------------------------
    
    xmlGetAttributeAsLong = CSafeLng(xmlGetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function xmlGetMandatoryAttribute( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Finds the attribute node specified by vstrAttributeName
'               in the XML node vxmlNode
' pass:         vxmlNode            Node to be searched
'               vstrAttributeName   Attribute to be found in the node
'               vlngMessageNo       Optional message number to override default message
' return:       IXMLDOMNode         The attribute node
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'------------------------------------------------------------------------------------------
    
    Dim xmlAttrib As IXMLDOMNode
    Set xmlAttrib = vobjNode.Attributes.getNamedItem(vstrAttribName)
    If xmlAttrib Is Nothing Then
        If vlngMessageNo = NO_MESSAGE_NUMBER Then
            errThrowError _
                "xmlGetMandatoryAttribute", _
                oeXMLMissingAttribute, _
                "[@" & vstrAttribName & "]"
        Else
            errThrowError "xmlGetMandatoryAttribute", vlngMessageNo
        End If
    Else
        
        Set xmlGetMandatoryAttribute = xmlAttrib
        Set xmlAttrib = Nothing
    End If
End Function
Public Function xmlGetMandatoryAttributeText( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As String
' header ----------------------------------------------------------------------------------
' description:      Gets the attribute node text
' pass:             vobjNode - xml Node to search for attribute
'                   vstrAttribName - attribute name to search for
'                   vlngMessageNo - Optional message number to override default message
' return:           String - Node text of attribute
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    Dim xmlAttrib As IXMLDOMNode
    Set xmlAttrib = xmlGetMandatoryAttribute(vobjNode, vstrAttribName, vlngMessageNo)
    If Len(xmlAttrib.Text) = 0 Then
        If vlngMessageNo = NO_MESSAGE_NUMBER Then
            errThrowError _
                "xmlGetMandatoryAttributeText", _
                oeXMLInvalidAttributeValue, _
                "[@" & vstrAttribName & "]"
        Else
            errThrowError "xmlGetMandatoryAttributeText", vlngMessageNo
        End If
    Else
        
        xmlGetMandatoryAttributeText = xmlAttrib.Text
    End If
End Function
Public Function xmlGetMandatoryAttributeAsBoolean( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Boolean
' header ----------------------------------------------------------------------------------
' description:  Finds the node specified by vstrMatchPattern in the XML node vxmlParentNode
' pass:         vxmlParentNode      Node to be searched
'               vstrMatchPattern    XSL search pattern
'               vlngMessageNo       Optional message number to override default message
' return:       True if attribute value is "1" or "Y" or "YES"
'               otherwise False
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    Dim strValue As String
    strValue = UCase(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo))
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetMandatoryAttributeAsBoolean = True
    End If
End Function
Public Function xmlGetMandatoryAttributeAsDate( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Date
' header ----------------------------------------------------------------------------------
' description:  Gets the attribute node as a date
' pass:         vobjNode - xml Node to search for attribute
'               vstrAttribName - attribute name to search for
'               vlngMessageNo - Optional message number to override default message
' return:       Date - Node text as a date
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlGetMandatoryAttributeAsDate = _
        CSafeDate(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo))
End Function
Public Function xmlGetMandatoryAttributeAsDouble( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Double
' header ----------------------------------------------------------------------------------
' description:  Gets the attribute node as a double
' pass:         vobjNode - xml Node to search for attribute
'               vstrAttribName - attribute name to search for
'               vlngMessageNo - Optional message number to override default message
' return:       Double - Node text as a double
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlGetMandatoryAttributeAsDouble = _
        CSafeDbl(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo))
End Function
Public Function xmlMandatoryGetAttributeAsInteger( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Integer
' header ----------------------------------------------------------------------------------
' description:  Gets the attribute node as a integer
' pass:         vobjNode - xml Node to search for attribute
'               vstrAttribName - attribute name to search for
'               vlngMessageNo - Optional message number to override default message
' return:       Integer - Node text as a integer
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlMandatoryGetAttributeAsInteger = _
        CSafeInt(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo))
End Function
Public Function xmlGetMandatoryAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER) _
    As Long
' header ----------------------------------------------------------------------------------
' description:  Gets the attribute node as a long
' pass:         vobjNode - xml Node to search for attribute
'               vstrAttribName - attribute name to search for
'               vlngMessageNo - Optional message number to override default message
' return:       Long - Node text as a long
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlGetMandatoryAttributeAsLong = _
        CSafeLng(xmlGetMandatoryAttributeText(vobjNode, vstrAttribName, vlngMessageNo))
End Function
Public Sub xmlCheckMandatoryAttribute( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttributeName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:  calls xmlGetMandatoryAttributeText to check that specified attribute exists
'               and has value,
'               xmlGetMandatoryAttributeText will raise an error if these conditions
'               are not met
' pass:         vxmlNode          Node to be searched
'               vstrAttributeName Attribute to be found in the node
'               vlngMessageNo     Optional message number to override default message
' exceptions:
'               raises error oeXMLMissingAttribute if attribute not found
'               raises error oeXMLInvalidAttributeValue if attribute has no value
'------------------------------------------------------------------------------------------
    xmlGetMandatoryAttributeText vxmlNode, vstrAttributeName, vlngMessageNo
End Sub
Public Sub xmlCopyAttribute( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String)
' header ----------------------------------------------------------------------------------
' description:
'   copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
'   will overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'------------------------------------------------------------------------------------------
    
    If xmlAttributeValueExists(vxmlSrceNode, vstrSrceAttribName) = True Then
        vxmlDestNode.Attributes.setNamedItem _
            vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).cloneNode(True)
    End If
End Sub
Public Sub xmlCopyMandatoryAttribute( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:
'   copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
'   will overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'   vlngMessageNo: Optional message number to override default message
' exceptions:
'   raises error oeXMLMissingAttribute if source attribute not found
'   raises error oeXMLInvalidAttributeValue if source attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlCheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName, vlngMessageNo
    xmlCopyAttribute vxmlSrceNode, vxmlDestNode, vstrSrceAttribName
End Sub
Public Sub xmlCopyAttributeValue( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String)
' header ----------------------------------------------------------------------------------
' description:
'   copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
'   to attribute named as vstrDestAttribName on vxmlDestNode
'   will overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'------------------------------------------------------------------------------------------
    
    If xmlAttributeValueExists(vxmlSrceNode, vstrSrceAttribName) = True Then
            
        Dim xmlAttrib As IXMLDOMAttribute
        Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName)
        xmlAttrib.Text = vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text
        vxmlDestNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
End Sub
Public Sub xmlCopyMandatoryAttributeValue( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:
'   copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
'   to attribute named as vstrDestAttribName on vxmlDestNode
'   will overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'   vlngMessageNo: Optional message number to override default message
' exceptions:
'   raises error oeXMLMissingAttribute if source attribute not found
'   raises error oeXMLInvalidAttributeValue if source attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlCheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName, vlngMessageNo
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName)
            
    xmlAttrib.Text = vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text
    vxmlDestNode.Attributes.setNamedItem xmlAttrib
    Set xmlAttrib = Nothing
End Sub
Public Sub xmlCopyAttribIfMissingFromDest( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String)
' header ----------------------------------------------------------------------------------
' description:
'   copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
'   will not overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'------------------------------------------------------------------------------------------
    
    If xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) = False Then
        xmlCopyAttribute vxmlSrceNode, vxmlDestNode, vstrSrceAttribName
    End If
End Sub
Public Sub xmlCopyMandatoryAttribIfMissingFromDest( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:
'   copies attribute named as vstrSrceAttribName, from vxmlSrceNode to vxmlDestNode
'   will not overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'   vlngMessageNo: Optional message number to override default message
' exceptions:
'   raises error oeXMLMissingAttribute if source attribute not found
'   raises error oeXMLInvalidAttributeValue if source attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlCheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName, vlngMessageNo
    xmlCopyAttribIfMissingFromDest vxmlSrceNode, vxmlDestNode, vstrSrceAttribName
End Sub
Public Sub xmlCopyAttribValueIfMissingFromDest( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String)
' header ----------------------------------------------------------------------------------
' description:
'   copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
'   to attribute named as vstrDestAttribName on vxmlDestNode
'   will not overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'------------------------------------------------------------------------------------------
    
    If xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) = False Then
        xmlCopyAttributeValue _
            vxmlSrceNode, vxmlDestNode, vstrSrceAttribName, vstrDestAttribName
    End If
End Sub
Public Sub xmlCopyMandatoryAttribValueIfMissingFromDest( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String, _
    Optional ByVal vlngMessageNo As Long = NO_MESSAGE_NUMBER)
' header ----------------------------------------------------------------------------------
' description:
'   copies value from attribute named as vstrSrceAttribName on vxmlSrceNode
'   to attribute named as vstrDestAttribName on vxmlDestNode
'   will not overwrite existing attribute value on vxmlDestNode
'   will do nothing if source attribute not present or has no value
' pass:
'   vxmlSrceNode: source node for attribute
'   vxmlDestNode: destination node for attribute
'   vstrSrceAttribName: name of attribute to be copied
'   vlngMessageNo: Optional message number to override default message
' exceptions:
'   raises error oeXMLMissingAttribute if source attribute not found
'   raises error oeXMLInvalidAttributeValue if source attribute has no value
'------------------------------------------------------------------------------------------
    
    xmlCheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName, vlngMessageNo
    If xmlAttributeValueExists(vxmlDestNode, vstrSrceAttribName) = False Then
        xmlCopyAttribValueIfMissingFromDest _
            vxmlSrceNode, vxmlDestNode, vstrSrceAttribName, vstrDestAttribName
    End If
End Sub
    
Public Sub xmlSetAttributeValue(ByVal vxmlDestNode As IXMLDOMNode, _
                                ByVal vstrAttribName As String, _
                                ByVal vstrAttribValue As String)
' header ----------------------------------------------------------------------------------
' description:
'   Creates a xml attribute on vxmlDestNode,
'   will overwrite existing attribute value
' pass:
'   vxmlDestNode - Destination Node for new attribute to be placed on
'   vstrAttribName - Attrbiute name of the new attribute
'   vstrAttribValue - Attribute value for new attribute on vxmlDestNode
'------------------------------------------------------------------------------------------
    
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName)
    xmlAttrib.Value = vstrAttribValue
        
    vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    Set xmlAttrib = Nothing
End Sub
Public Sub xmlSetAttributeValueIfMissingFromDest( _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    ByVal vstrAttribValue As String)
' header ----------------------------------------------------------------------------------
' description:
'   Creates a xml attribute on vxmlDestNode,
'   will not overwrite existing attribute value
' pass:
'   vxmlDestNode - Destination Node for new attribute to be placed on
'   vstrAttribName - Attrbiute name of the new attribute
'   vstrAttribValue - Attribute value for new attribute on vxmlDestNode
'------------------------------------------------------------------------------------------
    
    If xmlAttributeValueExists(vxmlDestNode, vstrAttribName) = False Then
        Dim xmlAttrib As IXMLDOMAttribute
        Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName)
        xmlAttrib.Value = vstrAttribValue
            
        vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
        Set xmlAttrib = Nothing
    End If
End Sub
Public Function xmlCreateElementRequestFromNode( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrMasterTagName As String, _
    ByVal blnConvertChildren As Boolean, _
    Optional ByVal vstrNewTagName As String = "") _
    As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   Convert a Request node from a attribute based to element based.
' pass:
'   vxmlNode            xml Request node to be converted
'   vstrMasterTagName   name of parent tag (can be pattern match)
'   bInConvertChildren  True/False - Convert all children of "MasterTagName" node
'   vstrNewTagName      New name for "Master Tag" (optional)
' return:
'   FreeThreadedDOMDocument40         Converted Request
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim xmlDOMDocument As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDestParentNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim strFunctionName As String
strFunctionName = "xmlCreateElementRequestFromNode"
    If Not vxmlNode Is Nothing Then
        Set xmlDOMDocument = New FreeThreadedDOMDocument40
        ' create a new request node
        Set xmlNode = xmlDOMDocument.createElement("REQUEST")
        If Not xmlNode Is Nothing Then
            Set xmlRequestNode = xmlDOMDocument.appendChild(xmlNode)
        End If
        ' clone the request attributes of the passed in request node
        Set xmlNode = vxmlNode.ownerDocument.selectSingleNode("REQUEST")
        For Each xmlNode In vxmlNode.Attributes
            xmlCopyAttribute vxmlNode, xmlRequestNode, xmlNode.nodeName
        Next
        'Extract the node to convert
        Set xmlNodeList = vxmlNode.selectNodes(vstrMasterTagName)
        If xmlNodeList.length = 0 Then
            errThrowError strFunctionName, oeXMLMissingElementText, "No matching nodes found for: " & vstrMasterTagName
        End If
        For Each xmlNode In xmlNodeList
            Set xmlNode = xmlMakeNodeElementBased(xmlNode, blnConvertChildren, vstrNewTagName)
            Set xmlDestParentNode = xmlRequestNode.appendChild(xmlNode)
        Next
        Set xmlCreateElementRequestFromNode = xmlDOMDocument
    End If
xmlCreateElementRequestFromNodeExit:
    Set xmlDOMDocument = Nothing
    Set xmlRequestNode = Nothing
    Set xmlDestParentNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
End Function
Public Function xmlMakeNodeElementBased( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal blnConvertChildren As Boolean, _
    ByVal vstrNewTagName As String, _
    ParamArray vstrAttributes()) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Convert a single node from a attribute based to element based.
' pass:
'   vxmlNode            xml Request node to be converted
'   bInConvertChildren  True/False - Convert all children of "MasterTagName" node
'   vstrNewTagName      New name for "Master Tag" (optional)
'   vstrAttributes      ParamArray can be used to name individual attributes to be
'                       converted. If not used all attributes are converted.
' return:
'   Node                Converted Node
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim xmlDOMDocument As FreeThreadedDOMDocument40
Dim xmlDestParentNode As IXMLDOMNode
Dim xmlAttribNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode, xmlNode2 As IXMLDOMNode
Dim strNodeName As String, strComboValueNodeName As String, strAttribName As String
Dim intIndex As Integer
Dim strNodeText As String
Dim intNoOfSpecifiedTags As Integer
Dim strFunctionName As String
strFunctionName = "xmlMakeNodeElementBased"
    If Not vxmlNode Is Nothing Then
        Set xmlDOMDocument = New FreeThreadedDOMDocument40
        'Determine the name of the new node
        If Len(Trim$(vstrNewTagName)) > 0 Then
            strNodeName = vstrNewTagName
        Else
            strNodeName = vxmlNode.nodeName
        End If
        ' create the new node
        Set xmlNode = xmlDOMDocument.createElement(strNodeName)
        If Not xmlNode Is Nothing Then
            Set xmlDestParentNode = xmlDOMDocument.appendChild(xmlNode)
        End If
        'Process each attribute
        ' have specified a subset of the attributes to clone?
        ' if we have not specified and specific tags the UBound = -1
        intNoOfSpecifiedTags = UBound(vstrAttributes)
        If intNoOfSpecifiedTags >= 0 Then
            For intIndex = 0 To intIndex > intNoOfSpecifiedTags
                strNodeText = xmlGetAttributeText(vxmlNode, vstrAttributes(intIndex))
                If Len(strNodeText) > 0 Then
                    'SR 12/06/01 : if node name is description of any combo value, create a new attribute
                    '              to the respective node
                    strAttribName = vstrAttributes(intIndex)
                    If Right(strAttribName, 5) = "_TEXT" Then
                        strComboValueNodeName = Left(strAttribName, Len(strAttribName) - 5)
                        Set xmlNode2 = xmlDestParentNode.selectSingleNode("./" & strComboValueNodeName)
                        If Not xmlNode2 Is Nothing Then
                            xmlSetAttributeValue xmlNode2, "TEXT", strNodeText
                        Else
                            Set xmlNode = xmlDOMDocument.createElement(vstrAttributes(intIndex))
                            xmlNode.Text = strNodeText
                            xmlDestParentNode.appendChild xmlNode
                        End If
                    Else
                        Set xmlNode = xmlDOMDocument.createElement(vstrAttributes(intIndex))
                        xmlNode.Text = strNodeText
                        xmlDestParentNode.appendChild xmlNode
                    End If
                Else
                    errThrowError strFunctionName, oeXMLMissingAttribute, ": " & vstrAttributes(intIndex)
                End If
            Next
        Else
            For Each xmlAttribNode In vxmlNode.Attributes
                'SR 12/06/01 : if node name is description of any combo value, create a new attribute
                '              to the respective node
                strAttribName = xmlAttribNode.nodeName
                If Right(strAttribName, 5) = "_TEXT" Then
                    strComboValueNodeName = Left(strAttribName, Len(strAttribName) - 5)
                    Set xmlNode2 = xmlDestParentNode.selectSingleNode("./" & strComboValueNodeName)
                    If Not xmlNode2 Is Nothing Then
                        xmlSetAttributeValue xmlNode2, "TEXT", xmlAttribNode.Text
                    Else
                        Set xmlNode = xmlDOMDocument.createElement(xmlAttribNode.nodeName)
                        If Not xmlNode Is Nothing Then
                            xmlNode.Text = xmlAttribNode.Text
                            xmlDestParentNode.appendChild xmlNode
                        End If
                    End If
                Else
                    Set xmlNode = xmlDOMDocument.createElement(xmlAttribNode.nodeName)
                    If Not xmlNode Is Nothing Then
                        xmlNode.Text = xmlAttribNode.Text
                        xmlDestParentNode.appendChild xmlNode
                    End If
                End If
            Next
        End If
        'If required, process each child element recursively (converting all attributes)
        If blnConvertChildren And vxmlNode.hasChildNodes Then
            For Each xmlNode In vxmlNode.childNodes
                Set xmlNode = xmlMakeNodeElementBased(xmlNode, blnConvertChildren, "")
                xmlDOMDocument.createElement xmlNode.nodeName
                xmlDestParentNode.appendChild xmlNode
            Next
        End If
        Set xmlMakeNodeElementBased = xmlDestParentNode
    End If
xmlMakeNodeElementBasedExit:
    Set xmlDOMDocument = Nothing
    Set xmlDestParentNode = Nothing
    Set xmlAttribNode = Nothing
    Set xmlNode = Nothing
    Set xmlNode2 = Nothing
End Function
Public Sub xmlCreateAttribute( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrName As String, _
    ByVal vstrValue As String)
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlAttrib = vxmlNode.ownerDocument.createAttribute(vstrName)
    xmlAttrib.Value = vstrValue
    vxmlNode.Attributes.setNamedItem xmlAttrib
    Set xmlAttrib = Nothing
End Sub
Public Function xmlCreateAttributeBasedResponse( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal blnConvertChild As Boolean) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Convert a Response node from element based to attribute based.
' pass:
'   vxmlNode            xml Response node to be converted
'   bInConvertChildren  True/False - Convert all children of parent node also
' return:
'   Node                Converted Request
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim xmlParent As IXMLDOMElement
Dim xmlNode As IXMLDOMNode
Dim blnNoChildren As Boolean
Dim xmlAttrib As IXMLDOMAttribute
    If Not vxmlNode Is Nothing Then
        'Clone the parent node (without child nodes) as a basis for the returned node
        Set xmlParent = vxmlNode.cloneNode(False)
        'Iterate through each child node
        For Each xmlNode In vxmlNode.childNodes
            'Check if it is a 'text' node or genuinely has child elements
            If xmlNode.hasChildNodes Then
                If xmlNode.childNodes(0).nodeName = "#text" Then
                    blnNoChildren = True
                End If
            Else
                blnNoChildren = True
            End If
            If blnNoChildren And Len(Trim$(xmlNode.Text)) > 0 Then
                'Create as an attribute of parent
                xmlParent.setAttribute xmlNode.nodeName, xmlNode.Text
                'Check if combo item with TEXT attribute
                For Each xmlAttrib In xmlNode.Attributes
                    If UCase$(xmlAttrib.nodeName) = "TEXT" Then
                        'Create a text attribute also for the combo text
                        xmlParent.setAttribute xmlNode.nodeName & "_TEXT", xmlAttrib.Text
                    End If
                Next
            ElseIf blnConvertChild And Not blnNoChildren Then
                'Process child node recursively
                Set xmlNode = xmlCreateAttributeBasedResponse(xmlNode, blnConvertChild)
                xmlParent.appendChild xmlNode
            End If
            blnNoChildren = False
        Next
        Set xmlCreateAttributeBasedResponse = xmlParent
    End If
xmlCreateAttributeBasedResponseExit:
    Set xmlAttrib = Nothing
    Set xmlParent = Nothing
    Set xmlNode = Nothing
End Function
Public Function xmlCreateChildRequest( _
    ByVal vxmlDataNode As IXMLDOMElement, _
    ByVal vxmlSchemaNode As IXMLDOMNode, _
    ByVal vstrNewName As String, _
    Optional ByVal blnComboLookup As Boolean = False) _
    As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create a Request node that can be used to find all matching records in a child table
' pass:
'   vxmlDataNode        xml element to be converted
'   vxmlSchemaNode      Schema entry for the parent table
'   vstrNewName         Name of the child node to be created
' return:
'   Node                Converted Request
' Raise Errors:
'------------------------------------------------------------------------------------------
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNewNode As IXMLDOMElement
Dim xmlSchemaElement As IXMLDOMNode
Dim xmlAttrib As IXMLDOMAttribute
Dim strValue As String
Dim strFunctionName As String
strFunctionName = "xmlCreateChildRequest"
    If (Not vxmlDataNode Is Nothing) And (Not vxmlSchemaNode Is Nothing) And Len(Trim$(vstrNewName)) > 0 Then
        'Create a new root node
        Set xmlDoc = New FreeThreadedDOMDocument40
        Set xmlNewNode = xmlDoc.createElement(vstrNewName)
        If blnComboLookup Then
            xmlNewNode.setAttribute "_COMBOLOOKUP_", "1"
        End If
        'Iterate through the schema node to get the primary key items
        For Each xmlSchemaElement In vxmlSchemaNode.childNodes
            If Not xmlSchemaElement.Attributes.getNamedItem("KEYTYPE") Is Nothing Then
                strValue = xmlSchemaElement.Attributes.getNamedItem("KEYTYPE").Text
                If strValue = "PRIMARY" Then
                    'Get this attribute from the data element and add it to the new node
                    Set xmlAttrib = vxmlDataNode.Attributes.getNamedItem(xmlSchemaElement.nodeName)
                    If Not xmlAttrib Is Nothing Then
                        xmlNewNode.setAttribute xmlSchemaElement.nodeName, xmlAttrib.Text
                    End If
                End If
            End If
        Next
        Set xmlCreateChildRequest = xmlNewNode
    Else
        errThrowError strFunctionName, oeInvalidParameter
    End If
xmlCreateChildRequestExit:
    Set xmlNewNode = Nothing
    Set xmlSchemaElement = Nothing
    Set xmlAttrib = Nothing
    Set xmlDoc = Nothing
End Function
Public Sub xmlElemFromAttrib( _
    ByVal vxmlDestParentNode As IXMLDOMNode, _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vstrAttribName As String)
' header ----------------------------------------------------------------------------------
' description:
'   Create a Element node from an attribute
'   with node name = attribute name and node value = attribute value
' pass:
'   vxmlDestParentNode  parent node for element to be created
'   vxmlSrceNode        node that contains attribute that is source of new element
'   vstrAttribName      Name of attribute to be located on vxmlSrceNode
'------------------------------------------------------------------------------------------
    
    Dim xmlElem As IXMLDOMElement
    Set xmlElem = vxmlDestParentNode.ownerDocument.createElement(vstrAttribName)
    If Not vxmlSrceNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
        xmlElem.Text = vxmlSrceNode.Attributes.getNamedItem(vstrAttribName).Text
    End If
    vxmlDestParentNode.appendChild xmlElem
    Set xmlElem = Nothing
End Sub
Public Function xmlGetRequestNode(ByVal vxmlElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the <REQUEST> node from an xml element, without any children. T
' pass:
'   vxmlElement The element from which to get the request node
' return:
'   Request node
' Raise Errors:
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName As String = "xmlGetRequestNode"
    Dim xmlDocument As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    If vxmlElement Is Nothing Then
        errThrowError cstrFunctionName, oeMissingElement
    End If
    ' If the top most tag of the node is 'REQUEST', just return the node, else
    ' search for 'REQUEST'
    If vxmlElement.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlElement.cloneNode(False)
    Else
        Set xmlNode = vxmlElement.selectSingleNode("//REQUEST")
        If xmlNode Is Nothing Then
            errThrowError cstrFunctionName, oeMissingPrimaryTag, "Expected REQUEST tag"
        End If
        Set xmlRequestNode = xmlNode.cloneNode(False)
    End If
    ' Attach the cloned node to a new dom document to ensure safety of using selectsinglenode
    ' with something like this: "/REQUEST/SOMETHING"
    Set xmlDocument = New FreeThreadedDOMDocument40
    Set xmlDocument.documentElement = xmlRequestNode
    Set xmlGetRequestNode = xmlRequestNode
    Exit Function
xmlGetRequestNodeExit:
    Set xmlDocument = Nothing
    Set xmlRequestNode = Nothing
    Set xmlNode = Nothing
End Function
Public Sub AttachResponseData(ByVal vxmlNodeToAttachTo As IXMLDOMNode, _
                              ByVal vxmlResponse As IXMLDOMElement)
' Header ----------------------------------------------------------------------------------
' Description:
'   Gets the data nodes out of vxmlResponse and appends them to vxmlNodeToAttachTo
' Pass:
'       vxmlNodeToAttachTo  xml node to attach the data to
'       vxmlResponse        xml element containing the response whos data is to be
'                            extracted
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "AttachResponseData"
    Dim xmlChildList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    If vxmlNodeToAttachTo Is Nothing Or vxmlResponse Is Nothing Then
        errThrowError strFunctionName, _
                        oeInvalidParameter, _
                        "Node to attach to or Response missing"
     End If
    If vxmlResponse.nodeName <> "RESPONSE" Then
        errThrowError strFunctionName, _
                        oeInvalidParameter, _
                        "RESPONSE must be top level tag"
    End If
         
     Set xmlChildList = vxmlResponse.childNodes
     For Each xmlNode In xmlChildList
        If xmlNode.nodeName <> "MESSAGE" Then
            vxmlNodeToAttachTo.appendChild xmlNode.cloneNode(True)
        End If
     Next
    Set xmlChildList = Nothing
    Set xmlNode = Nothing
     
End Sub
