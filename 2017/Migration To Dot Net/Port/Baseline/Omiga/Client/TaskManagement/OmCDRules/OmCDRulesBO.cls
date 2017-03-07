VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "OmCDRulesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------
'PROG   Date        Description
'LH     11/07/2006  EP979   Added XML request file tracing
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Public Function RunRules(ByVal vxmlRequest As String) As String

    gobjContext.SetComplete

    On Error GoTo RunRulesVbErr
            
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
            
            
    Call Trace.Main

    Debug.Print vxmlRequest
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    xmlRequestDoc.loadXML vxmlRequest
    
    'xmlRequestDoc.Save "C:\omiga4trace\CriticalDatRequest" & Format(Now(), "ddmmyyyyhhmmss") & ".xml"
    LogDetails 1, "Saving the Request to REQ_RES_Files folder", xmlRequestDoc, "CriticalDataRequest"
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
            
    Set xmlResponseElem = xmlResponseDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlResponseDoc.appendChild(xmlResponseElem)
    
    LogDetails 1, "----------------------------------------------------------------------------------------------------------------"
    LogDetails 1, "RunRules : Start"
    LogDetails 1, "xmlTransform: Start - Running XSL transformation using CDRules.xsl"
    
    'Run critical data rules (contained in xsl file) against XML request doc
    xmlTransform xmlRequestDoc, xmlResponseDoc, "CDRules.xsl"
    
    LogDetails 1, "xmlTransform: End"
    
    'xmlResponseDoc.Save "C:\omiga4trace\CriticalDatResponse" & Format(Now(), "ddmmyyyyhhmmss") & ".xml"
    LogDetails 1, "Saving the Response to REQ_RES_Files folder", xmlResponseDoc, "CriticalDataResponse"
    LogDetails 1, "RunRules : End"
    'AW 26/05/06    EP526   -   Close log file
    LogDetails 10, "", xmlResponseDoc
    
RunRulesExit:
    
    RunRules = xmlResponseDoc.xml
    Debug.Print RunRules
    
    Exit Function
    
RunRulesVbErr:

    xmlResponseElem.setAttribute "TYPE", "APPERR"       'MAR1021 GHun
    
    Set xmlElem = xmlResponseDoc.createElement("ERROR")
    Set xmlNode = xmlResponseElem.appendChild(xmlElem)
    
    Set xmlElem = xmlResponseDoc.createElement("NUMBER")        'MAR1021 GHun
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
    LogDetails 2, "Error Number : " & Err.Number
        
    Set xmlElem = xmlResponseDoc.createElement("SOURCE")        'MAR1021 GHun
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    LogDetails 2, "Error Source : " & Err.Source
    
    Set xmlElem = xmlResponseDoc.createElement("DESCRIPTION")       'MAR1021 GHun
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    LogDetails 2, "Error Description : " & Err.Description
    
    GoTo RunRulesExit

End Function



' ======================================================================================
' xml helper functions =================================================================

Private Function xmlGetNode( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrPattern As String) _
    As IXMLDOMNode
    
    If Not vxmlParentNode Is Nothing Then
        Set xmlGetNode = vxmlParentNode.selectSingleNode(vstrPattern)
    End If

End Function

Private Function xmlGetNodeText( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrPattern As String) _
    As String
    
    If Not vxmlParentNode Is Nothing Then
        If Not vxmlParentNode.selectSingleNode(vstrPattern) Is Nothing Then
            xmlGetNodeText = vxmlParentNode.selectSingleNode(vstrPattern).Text
        End If
    End If

End Function

Private Function xmlGetChildNodeByAttribute( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrChildNodeName As String, _
    ByVal vstrAttributeName As String) _
    As IXMLDOMNode
    
    If Not vxmlParentNode Is Nothing Then
        Dim strPattern As String
        strPattern = vstrChildNodeName & "[" & vstrAttributeName & "'"
        Set xmlGetChildNodeByAttribute = vxmlParentNode.selectSingleNode(strPattern)
    End If

End Function

Private Function xmlGetChildNodeByAttributeValue( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrChildNodeName As String, _
    ByVal vstrAttributeName As String, _
    ByVal vstrAttributeValue As String) _
    As IXMLDOMNode
    
    If Not vxmlParentNode Is Nothing Then
        Dim strPattern As String
        strPattern = vstrChildNodeName & "[@" & vstrAttributeName & "='" & vstrAttributeValue & "']"
        Set xmlGetChildNodeByAttributeValue = vxmlParentNode.selectSingleNode(strPattern)
    End If

End Function

Private Function xmlGetNodeList( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrPattern As String) _
    As IXMLDOMNodeList
    
    If Not vxmlParentNode Is Nothing Then
        Set xmlGetNodeList = vxmlParentNode.selectNodes(vstrPattern)
    End If

End Function

Private Function xmlAttributeExists( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            If Len(vxmlNode.Attributes.getNamedItem(vstrAttribName).Text) > 0 Then
                xmlAttributeExists = True
            End If
        End If
    End If

End Function

Private Function xmlGetAttributeText( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As String
    
    If Not vxmlNode Is Nothing Then
        If Not vxmlNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
            xmlGetAttributeText = vxmlNode.Attributes.getNamedItem(vstrAttribName).Text
        End If
    End If

End Function

Private Function xmlGetAttributeAsBoolean( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Boolean
    
    Dim strValue As String
    
    strValue = UCase(xmlGetAttributeText(vxmlNode, vstrAttribName))
    
    If strValue = "1" Or strValue = "Y" Or strValue = "YES" Then
        xmlGetAttributeAsBoolean = True
    End If
    
End Function

Private Function xmlGetAttributeAsDate( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Date
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsDate(strValue) Then
        xmlGetAttributeAsDate = CDate(strValue)
    End If

End Function

Private Function xmlGetAttributeAsDouble( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Double
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsDouble = CDbl(strValue)
    End If

End Function

Private Function xmlGetAttributeAsInteger( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Integer
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsInteger = CInt(strValue)
    End If

End Function

Private Function xmlGetAttributeAsLong( _
    ByVal vxmlNode As IXMLDOMNode, _
    ByVal vstrAttribName As String) _
    As Long
    
    Dim strValue As String
    
    strValue = xmlGetAttributeText(vxmlNode, vstrAttribName)
    
    If IsNumeric(strValue) Then
        xmlGetAttributeAsLong = CLng(strValue)
    End If

End Function

Private Sub xmlCopyAttribute( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrAttribName As String)
    
    If Not vxmlSrceNode Is Nothing Then
        If Not vxmlDestNode Is Nothing Then
            If Not vxmlSrceNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
                vxmlDestNode.Attributes.setNamedItem _
                    vxmlSrceNode.Attributes.getNamedItem(vstrAttribName).cloneNode(True)
            End If
        End If
    End If

End Sub

Private Sub xmlCopyAttributeValue( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode, _
    ByVal vstrSrceAttribName As String, _
    ByVal vstrDestAttribName As String)
    
    If Not vxmlSrceNode Is Nothing Then
        If Not vxmlDestNode Is Nothing Then
            If Not vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName) Is Nothing Then
                If Len(vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text) > 0 Then
                    Dim xmlAttrib As IXMLDOMAttribute
                    Set xmlAttrib = vxmlDestNode.ownerDocument.createAttribute(vstrDestAttribName)
                    xmlAttrib.Value = vxmlSrceNode.Attributes.getNamedItem(vstrSrceAttribName).Text
                    vxmlDestNode.Attributes.setNamedItem xmlAttrib
                    Set xmlAttrib = Nothing
                End If
            End If
        End If
    End If

End Sub

Private Sub xmlCreateAttribute( _
    ByVal vxmlParentNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    ByVal vstrAttribValue As String)
    
    If Not vxmlParentNode Is Nothing Then
        
        Dim xmlAttrib As IXMLDOMAttribute
        Set xmlAttrib = vxmlParentNode.ownerDocument.createAttribute(vstrAttribName)
        
        xmlAttrib.Value = vstrAttribValue
        
        vxmlParentNode.Attributes.setNamedItem xmlAttrib
        
        Set xmlAttrib = Nothing
    
    End If

End Sub
    

Private Sub xmlTransform(ByVal xmlRequestDoc As FreeThreadedDOMDocument40, ByVal xmlResponseDoc As FreeThreadedDOMDocument40, ByVal vstrXSLFileName)

    Dim xsldoc As FreeThreadedDOMDocument40
    
    Dim strFileSpec As String
    
    strFileSpec = App.Path & "\" & vstrXSLFileName
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set xsldoc = New FreeThreadedDOMDocument40
    xsldoc.async = False
    xsldoc.Load strFileSpec
    
    If xsldoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, "", _
                "error loading xslt document: " & _
                strFileSpec & _
                " " & xsldoc.parseError.reason
    End If
    
    xmlResponseDoc.loadXML xmlRequestDoc.transformNode(xsldoc)
    
    If xsldoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, "", _
                "error in transformation: " & _
                strFileSpec & _
                " " & xsldoc.parseError.reason
    End If

End Sub

