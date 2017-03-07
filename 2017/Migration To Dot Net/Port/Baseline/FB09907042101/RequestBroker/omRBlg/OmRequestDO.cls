VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OmRequestDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------
'EPSOM Specific History:
'Prog   Date        AQR         Description
'IK     01/09/2006  EP11        created to handle (legacy) omRB calls
'-----------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl

Public Function OmDataRequest(ByVal vstrXMLRequest As String) As String
    On Error GoTo OmDataRequestVbErr
    Const cstrFunctionName As String = "OmDataRequest"
    Dim strStatus As String
    strStatus = "test 1.0"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlTemplate As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.async = False
    xmlIn.setProperty "NewParser", True
    xmlIn.validateOnParse = False
    xmlIn.loadXML vstrXMLRequest
    If xmlIn.parseError.errorCode <> 0 Then
        xmlParserError xmlIn.parseError, cstrFunctionName
    End If
    '   IK_DEBUG
    #If IK_DEBUG Then
        Dim fso As Object
        Dim tname As String
        Dim sPath As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        tname = fso.GetTempName
        tname = Left(tname, Len(tname) - 4)
        sPath = App.Path & "\debug"
        If Not fso.FolderExists(sPath) Then
            fso.CreateFolder sPath
        End If
        xmlIn.Save sPath & "\omRbRequest_" & tname & ".xml"
        Set fso = Nothing
    #End If
    Set xmlRequestNode = xmlGetMandatoryNode(xmlIn, "REQUEST")
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.async = False
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    
    ' load named schema file (if present)
    If xmlAttributeValueExists(xmlRequestNode, "RB_SCHEMA") Then
        adoLoadNamedSchema xmlGetAttributeText(xmlRequestNode, "RB_SCHEMA")
    Else
        adoLoadNamedSchema "default"
    End If
    
    ' merge RB_TEMPLATE etc.
    If xmlAttributeValueExists(xmlRequestNode, "RB_TEMPLATE") Then
        MergeRequestTemplate xmlIn
        Set xmlRequestNode = xmlGetMandatoryNode(xmlIn, "REQUEST")
    End If
    
    ' pre-process REQUEST
    If xmlAttributeValueExists(xmlRequestNode, "RB_PREPROCREF") Then
        PreProcessRequest xmlIn
        Set xmlRequestNode = xmlGetMandatoryNode(xmlIn, "REQUEST")
    End If
    
    Debug.Print "REQUEST node"
    Debug.Print xmlRequestNode.xml
    
    GetDataRequest xmlRequestNode, xmlResponseNode
    
    ' post-process RESPONSE
    ' add REQUEST attributes to RESPONSE
    Set xmlResponseElem = xmlGetMandatoryNode(xmlOut, "RESPONSE")
    For Each xmlAttrib In xmlRequestNode.Attributes
        If xmlResponseElem.Attributes.getNamedItem(xmlAttrib.Name) Is Nothing Then
            xmlResponseElem.Attributes.setNamedItem xmlAttrib.cloneNode(True)
        End If
    Next
    ' add derived fiels etc.
    If xmlAttributeValueExists(xmlResponseElem, "RB_POSTPROCREF") Then
        PostProcessResponse xmlOut
        Set xmlResponseElem = xmlGetMandatoryNode(xmlOut, "RESPONSE")
    End If
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "STATUS", strStatus
    OmDataRequest = xmlOut.xml
    '   IK_DEBUG
    #If IK_DEBUG Then
        xmlOut.Save sPath & "\omRbResponse_" & tname & ".xml"
    #End If
OmDataRequestExit:
    
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Exit Function
OmDataRequestVbErr:
    If errIsSystemError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    OmDataRequest = errCreateErrorResponse
    '   go to clean-up section
    Resume OmDataRequestExit
End Function
Private Sub MergeRequestTemplate(ByVal vxmlRequestDoc As FreeThreadedDOMDocument40)
    
    Dim xmlTemplateDoc As FreeThreadedDOMDocument40
    Set xmlTemplateDoc = New FreeThreadedDOMDocument40
    xmlTemplateDoc.async = False
    Dim strTemplateName As String
    strTemplateName = vxmlRequestDoc.selectSingleNode("REQUEST/@RB_TEMPLATE").Text
    LoadRequestTemplate xmlTemplateDoc, strTemplateName
    If xmlTemplateDoc.parseError.errorCode = 0 Then
        
        If Not xmlTemplateDoc.selectSingleNode("REQUEST") Is Nothing Then
            
            MergeTemplateNode _
                xmlTemplateDoc.selectSingleNode("REQUEST"), _
                vxmlRequestDoc.selectSingleNode("REQUEST")
            vxmlRequestDoc.setProperty "NewParser", True
            vxmlRequestDoc.validateOnParse = False
            vxmlRequestDoc.loadXML xmlTemplateDoc.xml
        End If
    End If
End Sub
Private Sub LoadRequestTemplate(ByVal vxmlTemplateDoc As FreeThreadedDOMDocument40, _
                                ByVal vstrTemplateName As String)
    Dim strFileSpec As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & vstrTemplateName & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    vxmlTemplateDoc.setProperty "NewParser", True
    vxmlTemplateDoc.validateOnParse = False
    vxmlTemplateDoc.Load strFileSpec
    If vxmlTemplateDoc.parseError.errorCode <> 0 Then
        App.LogEvent _
            "omRB template load error, file : " & strFileSpec, _
            vbLogEventTypeWarning
        Exit Sub
    End If
    If Not vxmlTemplateDoc.selectNodes("REQUEST").length = 1 Then
        App.LogEvent _
            "omRB template error no REQUEST root node, file : " & strFileSpec, _
            vbLogEventTypeWarning
        Exit Sub
    End If
End Sub
Private Sub MergeTemplateNode(ByVal vxmlTemplateNode As IXMLDOMNode, ByVal vxmlRequestNode As IXMLDOMNode)
    
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlNode As IXMLDOMNode
    For Each xmlAttrib In vxmlRequestNode.Attributes
        If vxmlTemplateNode.Attributes.getNamedItem(xmlAttrib.Name) Is Nothing Then
            vxmlTemplateNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
        End If
    Next
    For Each xmlNode In vxmlTemplateNode.childNodes
        If Not vxmlRequestNode.selectSingleNode(xmlNode.nodeName) Is Nothing Then
            MergeTemplateNode _
                xmlNode, _
                vxmlRequestNode.selectSingleNode(xmlNode.nodeName)
        End If
    Next
End Sub
Private Sub PreProcessRequest(ByVal vxmlRequestDoc As FreeThreadedDOMDocument40)
    
    Dim objBO As Object
    On Error Resume Next
    Set objBO = GetObjectContext.CreateInstance("omRBPreProc.omRBPreProcBO")
    If Not objBO Is Nothing Then
        vxmlRequestDoc.setProperty "NewParser", True
        vxmlRequestDoc.validateOnParse = False
        vxmlRequestDoc.loadXML objBO.omRequest(vxmlRequestDoc.xml)
        Set objBO = Nothing
    End If
End Sub
Private Sub PostProcessResponse(ByVal vxmlResponseDoc As FreeThreadedDOMDocument40)
    
    Dim objBO As Object
    On Error Resume Next
    Set objBO = GetObjectContext.CreateInstance("omRBPostProc.omRBPostProcBO")
    If Not objBO Is Nothing Then
        vxmlResponseDoc.setProperty "NewParser", True
        vxmlResponseDoc.validateOnParse = False
        vxmlResponseDoc.loadXML objBO.omRequest(vxmlResponseDoc.xml)
        Set objBO = Nothing
    End If
End Sub
Private Sub GetDataRequest(ByVal xmlRequest As IXMLDOMNode, ByVal xmlResponse As IXMLDOMNode)
    
    Dim intIndex As Integer
    Dim blnDoComboLookUp As Boolean
    blnDoComboLookUp = xmlGetAttributeAsBoolean(xmlRequest, "COMBOLOOKUP")
    For intIndex = 0 To (xmlRequest.childNodes.length - 1)
        GetRequestChildNode _
            xmlRequest.childNodes.Item(intIndex), _
            xmlResponse, _
            blnDoComboLookUp
    Next
    ' IK_TODO ELEMENT based output
    If xmlGetAttributeText(xmlRequest, "OUTTYPE") = "ELEM" Then
        ConvertResponseToElements xmlResponse
    End If
End Sub
Private Sub GetRequestChildNode(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode, _
                                ByVal vblnDoComboLookUp As Boolean)
    Dim xmlSchemaNode As IXMLDOMNode
    Dim strSchemaName As String
    Dim strNodeName As String
        
    If Not xmlAttributeValueExists(vxmlRequestNode, gstrSchemaNameId) Then
        Err.Raise _
            vbObjectError + 512 + oeXMLMissingAttribute, _
            "GetRequestChildNode", _
            "missing _SCHEMA_ value in REQUEST root"
    End If
    strNodeName = UCase(vxmlRequestNode.nodeName)
    strSchemaName = xmlGetAttributeText(vxmlRequestNode, gstrSchemaNameId)
    Set xmlSchemaNode = _
        adoLegacyGetSchemaNode( _
            "OMSCHEMAS/SCHEMA[@NAME=""" & _
                strSchemaName & _
                """]//" & strNodeName & _
                "[@DATASRCE]")
    If xmlSchemaNode Is Nothing Then
        Err.Raise _
            vbObjectError + 512 + oeXMLMissingElement, _
            "GetRequestChildNode", _
            "missing SCHEMA node, SCHEMA name: " & strSchemaName & _
            "DATASRCE NODE name: " & strNodeName
    End If
    adoLegacyGetDatasrce vxmlRequestNode, vxmlResponseNode, xmlSchemaNode, vblnDoComboLookUp
    Set xmlSchemaNode = Nothing
End Sub
Private Sub ConvertResponseToElements(ByVal vxmlResponseNode As IXMLDOMNode)
    Dim xmlNode As IXMLDOMNode
    For Each xmlNode In vxmlResponseNode.childNodes
        ConvertResponseNodeToElements xmlNode
    Next
End Sub
Private Sub ConvertResponseNodeToElements(ByVal vxmlResponseNode As IXMLDOMNode)
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlElem As IXMLDOMElement
    For Each xmlNode In vxmlResponseNode.childNodes
        ConvertResponseNodeToElements xmlNode
    Next
    If vxmlResponseNode.Attributes.length > 0 Then
        For Each xmlAttrib In vxmlResponseNode.Attributes
            If (Right(xmlAttrib.Name, 5) <> "_TEXT") _
                And _
                (Right(xmlAttrib.Name, 4) <> "_RAW") _
            Then
                
                Set xmlElem = vxmlResponseNode.ownerDocument.createElement(xmlAttrib.Name)
                xmlElem.Text = xmlAttrib.Text
                If xmlAttributeValueExists(vxmlResponseNode, xmlAttrib.Name & "_TEXT") Then
                    xmlCopyAttributeValue _
                        vxmlResponseNode, xmlElem, xmlAttrib.Name & "_TEXT", "TEXT"
                End If
                        
                If xmlAttributeValueExists(vxmlResponseNode, xmlAttrib.Name & "_RAW") Then
                    xmlCopyAttributeValue _
                        vxmlResponseNode, xmlElem, xmlAttrib.Name & "_RAW", "RAW"
                End If
                vxmlResponseNode.appendChild xmlElem
            End If
        Next
        For Each xmlAttrib In vxmlResponseNode.Attributes
            vxmlResponseNode.Attributes.removeNamedItem xmlAttrib.Name
        Next
    End If
End Sub

Private Sub ObjectControl_Activate()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
'    adoEmptySchemaCache
End Sub
