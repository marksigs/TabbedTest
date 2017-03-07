Attribute VB_Name = "XMLHelper"
'EPSOM Specific History
'Prog    Date        AQR     Description
'DRC     06/09/2006  EP1118  Modified GetAttributeNode for non-existant nodes
Option Explicit



Public Function GetMandatoryNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrMatchPattern As String) _
    As IXMLDOMNode

    Set GetMandatoryNode = GetXMLNode(vxmlParentNode, vstrMatchPattern, True)

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
        'added 16/02/2006
        LogDetails 2, "Error Source : " & strErrSource
        LogDetails 2, "Error Description : " & Err.Description
            
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
            "[@" & vstrAttribName & "]"
    Else
        GetMandatoryAttributeText = xmlAttrib.Text
    End If

End Function

Public Function GetMandatoryAttributeAsLong( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Long
        
    GetMandatoryAttributeAsLong = _
        CSafeLng(GetMandatoryAttributeText(vobjNode, vstrAttribName))

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
    'EP1128 DRC
    If Not vxmlNode Is Nothing Then
        Set GetAttributeNode = vxmlNode.Attributes.getNamedItem(vstrAttributeName)
    End If
    'EP1128 - End
    
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
        Err.Raise eXMLMISSINGELEMENT, cstrFunctionName, "Match pattern:- " & vstrMatchPattern
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
            "[@" & vstrAttribName & "]"
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


