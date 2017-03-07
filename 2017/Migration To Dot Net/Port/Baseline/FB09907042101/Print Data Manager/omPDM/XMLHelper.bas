Attribute VB_Name = "XMLHelper"
Option Explicit
'TW     18/10/2005  MAR223

Public Function GetMandatoryNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As IXMLDOMNode
    Set GetMandatoryNode = GetXMLNode(vxmlParentNode, vstrMatchPattern, True)
End Function

Public Sub GetChangeNodeName(ByRef rxmlNode As IXMLDOMNode, _
                          ByVal vstrOldName As String, _
                          ByVal vstrNewName As String)
' header ----------------------------------------------------------------------------------
' description:
'   Changes all nodes present in rxmlNode that have a node name
'   of vstrOldName to vstrNewName
'
' pass:
'       rxmlNode    Xml Node
'       vstrOldName Name of nodes to be changed
'       vstrNewName Name nodes are to be changed to
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
' BBG1707 Changed to name so it doesn't conflict with xmlAssist
'------------------------------------------------------------------------------------------
On Error GoTo GetChangeNodeNameVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetChangeNodeName"
           
    Dim objNewNode As IXMLDOMNode
    Dim intListLength As Integer
    Dim intChildNodesLength As Integer
    Dim intChildIndex As Integer
    Dim intAttribIndex As Integer
                                     
    ' If this node is to be renamed create a new node with the new name
    ' and copy all the attributes across
    If rxmlNode.nodeName = vstrOldName Then
        Set objNewNode = rxmlNode.ownerDocument.createNode(rxmlNode.nodeType, vstrNewName, rxmlNode.namespaceURI)
        ' Copy attributes if this is an element node
        If objNewNode.nodeType = NODE_ELEMENT Then
            
            Dim objElement As IXMLDOMElement
            Set objElement = objNewNode
            intListLength = rxmlNode.Attributes.length - 1
            For intAttribIndex = 0 To intListLength
                objElement.setAttributeNode rxmlNode.Attributes(intAttribIndex).cloneNode(True)
            Next
        End If
    End If
             
    ' For all children of this node change their name if it is vstrOldName and
    ' append them to the new node
    intChildNodesLength = rxmlNode.childNodes.length - 1
    For intChildIndex = 0 To intChildNodesLength
        GetChangeNodeName rxmlNode.childNodes(intChildIndex), vstrOldName, vstrNewName
        If Not objNewNode Is Nothing Then
            objNewNode.appendChild rxmlNode.childNodes(intChildIndex).cloneNode(True)
        End If
    Next
    ' Replace the original node with the new node
    If Not objNewNode Is Nothing Then
        If Not (rxmlNode.parentNode Is Nothing) Then
            rxmlNode.parentNode.replaceChild objNewNode, rxmlNode
        End If
        Set rxmlNode = objNewNode
    End If
    Set objNewNode = Nothing
    Exit Sub
GetChangeNodeNameVbErr:
   
    Set objNewNode = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CreateErrorResponse() As String
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlReponseElem As IXMLDOMElement
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlDescriptionElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Set xmlReponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlReponseElem
    xmlReponseElem.setAttribute "TYPE", "APPERR"
    Set xmlErrorElem = xmlDoc.createElement("ERROR")
    xmlReponseElem.appendChild xmlErrorElem
    Set xmlElement = xmlDoc.createElement("NUMBER")
    xmlElement.Text = Err.Number
    xmlErrorElem.appendChild xmlElement
    Set xmlElement = xmlDoc.createElement("SOURCE")
    xmlElement.Text = Err.Source
    xmlErrorElem.appendChild xmlElement
    Set xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION")
    xmlDescriptionElem.Text = Err.Description
    xmlErrorElem.appendChild xmlDescriptionElem
    Set xmlElement = xmlDoc.createElement("VERSION")
    xmlElement.Text = App.Comments
    xmlErrorElem.appendChild xmlElement
            
    CreateErrorResponse = xmlReponseElem.xml
        
    Set xmlDoc = Nothing
    Set xmlReponseElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlDescriptionElem = Nothing
    Set xmlElement = Nothing
End Function
Public Sub CheckError( _
    ByVal vstrFunctionName As String, _
    Optional ByVal vstrObjectName As String)
    Dim strErrSource As String
        
    If Err.Number <> 0 Then
        strErrSource = Err.Source
        If strErrSource <> vstrFunctionName Then
            strErrSource = vstrFunctionName & "." & strErrSource
        End If
            
        If Len(vstrObjectName) <> 0 Then
            strErrSource = vstrObjectName & "." & strErrSource
        End If
            
        Err.Raise Err.Number, strErrSource, Err.Description
            
    End If
End Sub
Public Function GetMandatoryAttributeText( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As String
    Dim xmlAttrib As IXMLDOMNode
    Set xmlAttrib = GetMandatoryAttribute(vobjNode, vstrAttribName)
    If Len(xmlAttrib.Text) = 0 Then
        Err.Raise _
            eXMLINVALIDATTTRIBUTEVALUE, _
            "GetMandatoryAttributeText", _
            "[@" & vstrAttribName & "] has invalid attribute value"
    Else
        GetMandatoryAttributeText = xmlAttrib.Text
    End If
End Function
Public Function GetMandatoryAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Long
        
    GetMandatoryAttributeAsLong = CSafeLng(GetMandatoryAttributeText(vobjNode, vstrAttribName))
End Function
Public Function GetAttributeAsInteger( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Integer
    GetAttributeAsInteger = CSafeInt(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Sub SetAttributeValue(ByVal vxmlDestNode As IXMLDOMNode, _
                                ByVal vstrAttribName As String, _
                                ByVal vstrAttribValue As String)
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrAttribName)
    xmlAttrib.Value = vstrAttribValue
        
    vxmlDestNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
    Set xmlAttrib = Nothing
End Sub
Public Sub CopyMandatoryAttribute( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String)
    CheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName
    CopyAttribute vxmlSrceNode, vxmlDestNode, vstrSrceAttribName
End Sub
Public Sub CheckMandatoryAttribute( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttributeName As String)
    GetMandatoryAttributeText vxmlNode, vstrAttributeName
End Sub
Public Sub CopyAttributeValue( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String)
    If AttributeValueExists(vxmlSrceNode, vstrSrceAttribName) = True Then
            
        Dim xmlAttrib As IXMLDOMAttribute
        Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName)
        xmlAttrib.Text = vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text
        vxmlDestNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
End Sub
Public Function GetAttributeText( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As String
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = GetAttributeNode(vobjNode, vstrAttribName)
    If Not xmlNode Is Nothing Then
        GetAttributeText = xmlNode.Text
        Set xmlNode = Nothing
    Else
        GetAttributeText = vstrDefault
    End If
End Function
Public Function GetAttributeNode( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttributeName As String) _
    As IXMLDOMNode
    Set GetAttributeNode = vxmlNode.Attributes.getNamedItem(vstrAttributeName)
End Function
Public Sub CopyMandatoryAttributeValue( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String)
    CheckMandatoryAttribute vxmlSrceNode, vstrSrceAttribName
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName)
            
    xmlAttrib.Text = vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text
    vxmlDestNode.Attributes.setNamedItem xmlAttrib
    Set xmlAttrib = Nothing
End Sub
Public Function GetAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Long
    GetAttributeAsLong = CSafeLng(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function GetNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As IXMLDOMNode
    Set GetNode = GetXMLNode(vxmlParentNode, vstrMatchPattern)
End Function
Private Function GetXMLNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String, _
    Optional ByVal vblnMandatoryNode As Boolean = False) As IXMLDOMNode
    Const cstrFunctionName As String = "GetXMLNode"
    Dim xmlNode As IXMLDOMNode
    If vxmlParentNode Is Nothing Then
        Err.Raise eXMLMISSINGELEMENT, cstrFunctionName, "Missing parent node"
    End If
    Set xmlNode = vxmlParentNode.selectSingleNode(vstrMatchPattern)
    If vblnMandatoryNode = True And xmlNode Is Nothing Then
        Err.Raise eXMLMISSINGELEMENT, cstrFunctionName, "Missing element: Match pattern:- " & vstrMatchPattern
    End If
    Set GetXMLNode = xmlNode
    Set xmlNode = Nothing
End Function
Public Function GetMandatoryAttribute( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMNode
    Set xmlAttrib = vobjNode.Attributes.getNamedItem(vstrAttribName)
    If xmlAttrib Is Nothing Then
        Err.Raise _
            eXMLINVALIDATTTRIBUTE, _
            "GetMandatoryAttribute", _
            "[@" & vstrAttribName & "] attribute missing"
    Else
        
        Set GetMandatoryAttribute = xmlAttrib
        Set xmlAttrib = Nothing
    End If
End Function
Public Function AttributeValueExists( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = GetAttributeNode(vobjNode, vstrAttribName)
    If Not xmlNode Is Nothing Then
        If Len(xmlNode.Text) > 0 Then
            AttributeValueExists = True
        End If
    End If
End Function
Public Function GetAttributeAsBoolean( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Boolean
    Dim strValue As String
    strValue = UCase(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        GetAttributeAsBoolean = True
    End If
End Function
Public Function GetAttributeAsDate( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Date
        
    GetAttributeAsDate = CSafeDate(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Function GetAttributeAsDouble( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    Optional ByVal vstrDefault As String = "") _
    As Double
        
    GetAttributeAsDouble = CSafeDbl(GetAttributeText(vobjNode, vstrAttribName, vstrDefault))
End Function
Public Sub CopyAttribute( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String)
    If AttributeValueExists(vxmlSrceNode, vstrSrceAttribName) = True Then
        vxmlDestNode.Attributes.setNamedItem _
            vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).cloneNode(True)
    End If
End Sub
Public Function LoadXML( _
    ByVal vstrXMLData As String, _
    ByVal vstrCallingFunction As String) _
    As FreeThreadedDOMDocument40
    Const cstrFunctionName As String = "LoadXML"
    Dim strErrDesc As String
    Dim xmlParseError As IXMLDOMParseError
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.async = False         ' wait until XML is fully loaded
    objXmlDoc.setProperty "NewParser", True
    objXmlDoc.validateOnParse = False
    objXmlDoc.LoadXML vstrXMLData
    Set xmlParseError = objXmlDoc.parseError
    If xmlParseError.errorCode <> 0 Then
        strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & xmlParseError.reason & vbCr & _
        "Error code: " & Str(xmlParseError.errorCode) & vbCr & _
        "Line no.: " & Str(xmlParseError.Line) & vbCr & _
        "Character: " & Str(xmlParseError.linepos) & vbCr & _
        "Source text: " & xmlParseError.srcText
    End If
    Set LoadXML = objXmlDoc
    Set xmlParseError = Nothing
End Function
Public Function MakeNodeElementBased( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal blnConvertChildren As Boolean, _
    ByVal vstrNewTagName As String, _
    ParamArray vstrAttributes()) _
    As IXMLDOMNode
Const cstrFunctionName As String = "MakeNodeElementBased"
Dim xmlDOMDocument As FreeThreadedDOMDocument40
Dim xmlDestParentNode As IXMLDOMNode
Dim xmlAttribNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode, xmlNode2 As IXMLDOMNode
Dim strNodeName As String, strComboValueNodeName As String, strAttribName As String
Dim intIndex As Integer
Dim strNodeText As String
Dim intNoOfSpecifiedTags As Integer
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
                strNodeText = GetAttributeText(vxmlNode, vstrAttributes(intIndex))
                If Len(strNodeText) > 0 Then
                    'SR 12/06/01 : if node name is description of any combo value, create a new attribute
                    '              to the respective node
                    strAttribName = vstrAttributes(intIndex)
                    If Right(strAttribName, 5) = "_TEXT" Then
                        strComboValueNodeName = Left(strAttribName, Len(strAttribName) - 5)
                        Set xmlNode2 = xmlDestParentNode.selectSingleNode("./" & strComboValueNodeName)
                        If Not xmlNode2 Is Nothing Then
                            SetAttributeValue xmlNode2, "TEXT", strNodeText
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
                    Err.Raise eXMLMISSINGATTRIBUTE, cstrFunctionName, vstrAttributes(intIndex) & " attribute missing"
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
                        SetAttributeValue xmlNode2, "TEXT", xmlAttribNode.Text
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
                Set xmlNode = MakeNodeElementBased(xmlNode, blnConvertChildren, "")
                xmlDOMDocument.createElement xmlNode.nodeName
                xmlDestParentNode.appendChild xmlNode
            Next
        End If
        Set MakeNodeElementBased = xmlDestParentNode
    End If
MakeNodeElementBasedExit:
    Set xmlDOMDocument = Nothing
    Set xmlDestParentNode = Nothing
    Set xmlAttribNode = Nothing
    Set xmlNode = Nothing
    Set xmlNode2 = Nothing
End Function
Public Function CheckXMLResponse( _
    ByVal vstrXmlReponse As String, _
    Optional ByVal vblnRaiseError As Boolean = False, _
    Optional ByVal vxmlInResponseElement As IXMLDOMElement = Nothing) _
    As Long
    Const cstrFunctionName As String = "CheckXMLResponse"
    Dim xmlXmlIn As FreeThreadedDOMDocument40
    Dim xmlResponseElement As IXMLDOMElement
    Set xmlXmlIn = LoadXML(vstrXmlReponse, cstrFunctionName)
    Set xmlResponseElement = xmlXmlIn.selectSingleNode("RESPONSE")
    If xmlResponseElement Is Nothing Then
        Err.Raise eXMLMISSINGELEMENT, cstrFunctionName, "Missing RESPONSE element"
    End If
    CheckXMLResponse = CheckXMLResponseNode(xmlResponseElement, vxmlInResponseElement, vblnRaiseError)
    Set xmlXmlIn = Nothing
    Set xmlResponseElement = Nothing
End Function
Public Function CheckXMLResponseNode(ByVal vxmlResponseNodeToCheck As IXMLDOMElement, _
                            Optional ByVal vxmlResponseNodeToAdd As IXMLDOMElement, _
                            Optional ByVal vblnRaiseError As Boolean = False) As Long
    Dim lngErrNo As Long
    Dim strErrSource As String, _
        strErrDesc As String
    lngErrNo = eUNSPECIFIEDERROR
    strErrSource = App.Title
    If vxmlResponseNodeToCheck Is Nothing Then
        Err.Raise eXMLMISSINGELEMENT, strErrSource, "Missing RESPONSE element"
    End If
    If vxmlResponseNodeToCheck.nodeName <> "RESPONSE" Then
        Err.Raise eXMLMISSINGELEMENT, strErrSource, "RESPONSE must be the top level tag"
    End If
    If Not vxmlResponseNodeToAdd Is Nothing Then
        If vxmlResponseNodeToAdd.nodeName <> "RESPONSE" Then
            Err.Raise eXMLMISSINGELEMENT, strErrSource, "RESPONSE must be the top level tag"
        End If
    End If
    Dim xmlResponseTypeNode As IXMLDOMNode
    Set xmlResponseTypeNode = vxmlResponseNodeToCheck.Attributes.getNamedItem("TYPE")
    If Not xmlResponseTypeNode Is Nothing Then
        If xmlResponseTypeNode.Text = "WARNING" Then
            If Not vxmlResponseNodeToAdd Is Nothing Then
                vxmlResponseNodeToAdd.setAttribute "TYPE", "WARNING"
                Dim xmlMessageList As IXMLDOMNodeList
                Dim xmlMessageElem As IXMLDOMElement
                Dim xmlFirstChild  As IXMLDOMElement
                Set xmlFirstChild = vxmlResponseNodeToAdd.firstChild
                Set xmlMessageList = vxmlResponseNodeToCheck.selectNodes("MESSAGE")
                ' insert messages at the top of the response to add element
                For Each xmlMessageElem In xmlMessageList
                    vxmlResponseNodeToAdd.insertBefore xmlMessageElem.cloneNode(True), xmlFirstChild
                Next
            End If
        ElseIf xmlResponseTypeNode.Text <> "SUCCESS" Then
            
            Dim xmlResponseErrorNumber As IXMLDOMNode, _
                xmlResponseErrorSource As IXMLDOMNode, _
                xmlResponseErrorDesc As IXMLDOMNode
            Set xmlResponseErrorNumber = vxmlResponseNodeToCheck.selectSingleNode("ERROR/NUMBER")
            If Not xmlResponseErrorNumber Is Nothing Then
                If IsNumeric(xmlResponseErrorNumber.Text) = True Then
                    lngErrNo = CSafeLng(xmlResponseErrorNumber.Text)
                End If
            End If
            Set xmlResponseErrorNumber = Nothing
            If vblnRaiseError Then
                
                Set xmlResponseErrorSource = vxmlResponseNodeToCheck.selectSingleNode("ERROR/SOURCE")
                If Not xmlResponseErrorSource Is Nothing Then
                    If Len(xmlResponseErrorSource.Text) > 0 Then
                        strErrSource = xmlResponseErrorSource.Text
                    End If
                End If
                Set xmlResponseErrorSource = Nothing
                Set xmlResponseErrorDesc = vxmlResponseNodeToCheck.selectSingleNode("ERROR/DESCRIPTION")
                If Not xmlResponseErrorDesc Is Nothing Then
                    If Len(xmlResponseErrorDesc.Text) > 0 Then
                        strErrDesc = xmlResponseErrorDesc.Text
                    End If
                End If
                Set xmlResponseErrorDesc = Nothing
                If Len(strErrDesc) = 0 Then
                    strErrDesc = "Unspecified Error"
                End If
                            
                Err.Raise lngErrNo, strErrSource, strErrDesc
            End If
        Else
            lngErrNo = 0
        End If
    End If
    CheckXMLResponseNode = lngErrNo
End Function
Public Function CreateElementRequestFromNode( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrMasterTagName As String, _
    ByVal blnConvertChildren As Boolean, _
    Optional ByVal vstrNewTagName As String = "") _
    As FreeThreadedDOMDocument40
Const cstrFunctionName As String = "CreateElementRequestFromNode"
Dim xmlDOMDocument As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDestParentNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
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
            CopyAttribute vxmlNode, xmlRequestNode, xmlNode.nodeName
        Next
        'Extract the node to convert
        Set xmlNodeList = vxmlNode.selectNodes(vstrMasterTagName)
        If xmlNodeList.length = 0 Then
            Err.Raise eXMLMISSINGELEMENT, cstrFunctionName, "No matching nodes found for: " & vstrMasterTagName
        End If
        For Each xmlNode In xmlNodeList
            Set xmlNode = MakeNodeElementBased(xmlNode, blnConvertChildren, vstrNewTagName)
            Set xmlDestParentNode = xmlRequestNode.appendChild(xmlNode)
        Next
        Set CreateElementRequestFromNode = xmlDOMDocument
    End If
CreateElementRequestFromNodeExit:
    Set xmlDOMDocument = Nothing
    Set xmlRequestNode = Nothing
    Set xmlDestParentNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
End Function
Public Function GetNodeText( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As String
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = GetNode(vxmlParentNode, vstrMatchPattern)
    If Not xmlNode Is Nothing Then
        GetNodeText = xmlNode.Text
        Set xmlNode = Nothing
    End If
End Function
Public Function CreateAttributeBasedResponse( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal blnConvertChild As Boolean) _
    As IXMLDOMNode
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
                Set xmlNode = CreateAttributeBasedResponse(xmlNode, blnConvertChild)
                xmlParent.appendChild xmlNode
            End If
            blnNoChildren = False
        Next
        Set CreateAttributeBasedResponse = xmlParent
    End If
CreateAttributeBasedResponseExit:
    Set xmlAttrib = Nothing
    Set xmlParent = Nothing
    Set xmlNode = Nothing
End Function
