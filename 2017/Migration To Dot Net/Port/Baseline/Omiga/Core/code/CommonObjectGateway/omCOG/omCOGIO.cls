VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "omCOGIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCOGIO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:   Omiga Common Object Gateway Interface Object
'               Used as a common entry point into the Omiga system. Runs any client
'               specified rules before running the specified Omiga object.
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     25/05/2001  Created
'PSC    01/06/2001  Amended so that the client validation response was not overwritten if
'                   the validation failed
'PSC    10/10/01    SYS2788 - Amend to take into account methods that take no parameters
'                   or non xml parameters
'------------------------------------------------------------------------------------------
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
Public Function omCOGRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo omCOGRequestVbErr
    Const strFunctionName As String = "omCOGRequest"
    Dim xmlInDoc As FreeThreadedDOMDocument40, _
        xmlValidateResponseDoc As FreeThreadedDOMDocument40, _
        xmlPrimaryRequestDoc As FreeThreadedDOMDocument40, _
        xmlPrimaryResponseDoc As FreeThreadedDOMDocument40, _
        xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlCVINode As IXMLDOMNode
    Dim xmlOPINode As IXMLDOMNode
    Dim xmlValidateResponseNode As IXMLDOMNode
    Dim xmlPrimaryRequestNode As IXMLDOMNode
    Dim xmlPrimaryResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    Dim blnFailedValidation As Boolean
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete
    ' create REQUEST & RESPONSE FreeThreadedDomDocument4040s
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlInDoc.async = False
    xmlOutDoc.async = False
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXMLRequest
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    ' OPI (Omiga Processing Interface) Element is mandatory
    Set xmlOPINode = xmlGetMandatoryNode(xmlRequestNode, "OPI")
    xmlCheckMandatoryAttribute xmlOPINode, "COMPONENT"
    xmlCheckMandatoryAttribute xmlOPINode, "OBJECT"
    xmlCheckMandatoryAttribute xmlOPINode, "METHOD"
    ' CVI (Client Validation Interface) Element is optional
    Set xmlCVINode = xmlRequestNode.selectSingleNode("CVI")
    If Not xmlCVINode Is Nothing Then
        
        xmlCheckMandatoryAttribute xmlCVINode, "COMPONENT"
        xmlCheckMandatoryAttribute xmlCVINode, "OBJECT"
        xmlCheckMandatoryAttribute xmlCVINode, "METHOD"
        Set xmlValidateResponseDoc = New FreeThreadedDOMDocument40
        xmlValidateResponseDoc.async = False
        CallUserValidation xmlRequestNode, xmlValidateResponseDoc
        Set xmlValidateResponseNode = _
            xmlValidateResponseDoc.selectSingleNode("RESPONSE")
        If Not xmlValidateResponseNode Is Nothing Then
            
            If xmlGetAttributeText(xmlValidateResponseNode, "TYPE") <> "SUCCESS" Then
                
                omCOGRequest = xmlValidateResponseDoc.xml
                blnFailedValidation = True
            End If
        End If
    End If
            
    If Not blnFailedValidation Then
        Set xmlPrimaryRequestDoc = New FreeThreadedDOMDocument40
        xmlPrimaryRequestDoc.async = False
        ConvertRequest xmlRequestNode, xmlPrimaryRequestDoc
        Set xmlPrimaryRequestNode = xmlPrimaryRequestDoc.selectSingleNode("REQUEST")
        Set xmlPrimaryResponseDoc = New FreeThreadedDOMDocument40
        xmlPrimaryResponseDoc.async = False
        ForwardRequest xmlRequestNode, xmlPrimaryRequestNode, xmlPrimaryResponseDoc
        Set xmlPrimaryResponseNode = xmlPrimaryResponseDoc.selectSingleNode("RESPONSE")
        If ResponseConversionRequired(xmlOPINode.Attributes.getNamedItem("COMPONENT").Text) _
        Then
            ConvertResponse xmlPrimaryResponseNode, xmlOutDoc
        End If
        ' PSC 01/06/2001
        omCOGRequest = xmlOutDoc.xml
        ' TODO - add any MESSAGE response blocks from validation
    End If
    ' debug only
    ' Terminate xmlOutDoc, "ConvertResponse"
    Debug.Print omCOGRequest
omCOGRequestExit:
    
    ' clean up ================================================================
    ' release XML DOM objects
    Exit Function
omCOGRequestVbErr:
    
    ' create ERROR response block from Err object
    omCOGRequest = CreateErrorResponse()
    Debug.Print omCOGRequest
    ' perform clean up
    GoTo omCOGRequestExit
End Function
Private Sub CallUserValidation( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseDoc As FreeThreadedDOMDocument40)
    Const cstrFunctionName As String = "omCOG.CallUserValidation"
    Dim strComponent As String, _
        strObject As String, _
        strProgId As String, _
        strMethod As String, _
        strErrDesc As String
    Dim lngErrNo As Long
    Dim varResponse As Variant
    Dim objBO As Object
        
    ' get routing details
    strComponent = vxmlRequestNode.selectSingleNode("CVI/@COMPONENT").Text
    strObject = vxmlRequestNode.selectSingleNode("CVI/@OBJECT").Text
    strMethod = vxmlRequestNode.selectSingleNode("CVI/@METHOD").Text
    strProgId = strComponent & "." & strObject
    Err.Clear
    On Error Resume Next
#If DEBUG_ON Then
    Set objBO = CreateObject(strProgId)
#Else
    Set objBO = gobjContext.CreateInstance(strProgId)
#End If
    ' disable error handling in this module
    On Error GoTo 0
    If objBO Is Nothing Then
        Err.Raise _
            oeObjectNotCreatable, _
            cstrFunctionName, _
            "omCOG CVI object(" & strProgId & ") not found"
    End If
    On Error Resume Next
    varResponse = CallByName(objBO, strMethod, VbMethod, vxmlRequestNode.xml)
    lngErrNo = Err.Number
    strErrDesc = Err.Description
    Set objBO = Nothing
    On Error GoTo 0
    If lngErrNo <> 0 Then
        Err.Raise _
            lngErrNo, _
            "omCOG", _
            "omCOG CVI method(" & strProgId & "." & strMethod & ") " & strErrDesc
    End If
    If IsEmpty(varResponse) Then
        Err.Raise _
            oeNotImplemented, _
            cstrFunctionName, _
            "omCOG CVI method(" & strProgId & "." & strMethod & ") no response"
    End If
    vxmlResponseDoc.setProperty "NewParser", True
    vxmlResponseDoc.validateOnParse = False
    vxmlResponseDoc.loadXML varResponse
End Sub
Private Sub ConvertRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlPrimaryRequestDoc As FreeThreadedDOMDocument40)
    Dim xmlForwardReqNode As IXMLDOMNode
    Dim xmlOPINode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlParameterList As IXMLDOMNodeList
    Dim strComponent As String
    Set xmlOPINode = vxmlRequestNode.selectSingleNode("OPI")
    strComponent = xmlOPINode.Attributes.getNamedItem("COMPONENT").Text
    Set xmlForwardReqNode = _
        vxmlPrimaryRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    If Not xmlOPINode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        If GetConversionType(strComponent) = 2 Then
            ' phase 2 components
            ' copy OPERATION to REQUEST node
            xmlCopyAttribute xmlOPINode, xmlForwardReqNode, "OPERATION"
        Else
            ' phase 1 components
            ' if OPERATION tag present,
            ' create as child of REQUEST (e.g. original SEARCH, CREATE, UPDATE tags)
            Set xmlElem = _
                vxmlPrimaryRequestDoc.createElement(xmlOPINode.Attributes.getNamedItem("OPERATION").Text)
            Set xmlForwardReqNode = xmlForwardReqNode.appendChild(xmlElem)
        End If
    End If
    Set xmlParameterList = xmlOPINode.selectNodes(".//OPI_PARAMETER")
        
    ' Only convert if the component requires element based XML and XML is being passed
    ' into the methods
    If RequestConversionRequired(strComponent) = True And _
       xmlParameterList.length = 0 Then
        For Each xmlNode In xmlOPINode.childNodes
            xmlToElement xmlNode, xmlForwardReqNode
        Next
    Else
        For Each xmlNode In xmlOPINode.childNodes
            xmlForwardReqNode.appendChild xmlNode.cloneNode(True)
        Next
    End If
    Debug.Print "request: "
    Debug.Print vxmlPrimaryRequestDoc.xml
End Sub
Private Sub ForwardRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlPrimaryRequestNode As IXMLDOMNode, _
    ByVal vxmlPrimaryResponseDoc As FreeThreadedDOMDocument40)
    Const cstrFunctionName As String = "ForwardRequest"
    Dim strComponent As String, _
        strObject As String, _
        strProgId As String, _
        strMethod As String, _
        strErrDesc As String
    Dim lngErrNo As Long
    Dim varResponse As Variant
    Dim varParameters() As Variant
    Dim xmlParmeterList As IXMLDOMNodeList
    Dim xmlParameter As IXMLDOMElement
    Dim objBO As Object
        
    ' get routing details
    strComponent = vxmlRequestNode.selectSingleNode("OPI/@COMPONENT").Text
    strObject = vxmlRequestNode.selectSingleNode("OPI/@OBJECT").Text
    strMethod = vxmlRequestNode.selectSingleNode("OPI/@METHOD").Text
    strProgId = strComponent & "." & strObject
    Err.Clear
    On Error Resume Next
#If DEBUG_ON Then
    Set objBO = CreateObject(strProgId)
#Else
    Set objBO = gobjContext.CreateInstance(strProgId)
#End If
    ' disable error handling in this module
    On Error GoTo 0
    If objBO Is Nothing Then
        Err.Raise _
            oeObjectNotCreatable, _
            cstrFunctionName, _
            "omCOG OPI object(" & strProgId & ") not found"
    End If
    On Error Resume Next
    Set xmlParmeterList = vxmlPrimaryRequestNode.selectNodes(".//OPI_PARAMETER")
    ' PSC 10/10/01 SYS2788 - Start
    ' Check if we have a proper request, are passing non XML arguments to the
    ' method, or passing nothing to the method.
    ' If we are passing a proper request send it down to CallByNameEx.
    ' If we are passing individual parameters then pull off each parameter and
    ' send them down as an array of parameters parameters to CallByNameEx.
    ' If there are no parameters send no parameters down to CallByNameEx
    If vxmlPrimaryRequestNode.childNodes.length > 0 Or xmlParmeterList.length > 0 Then
        If xmlParmeterList.length > 0 Then
            ReDim varParameters(xmlParmeterList.length - 1)
            For Each xmlParameter In xmlParmeterList
                varParameters(UBound(varParameters)) = xmlParameter.getAttribute("VALUE")
            Next
        ElseIf vxmlPrimaryRequestNode.childNodes.length > 0 Then
            ReDim varParameters(0)
            varParameters(UBound(varParameters)) = vxmlPrimaryRequestNode.xml
        End If
        varResponse = CallByNameEx(objBO, strMethod, VbMethod, varParameters)
    ElseIf vxmlPrimaryRequestNode.childNodes.length = 0 Then
        varResponse = CallByNameEx(objBO, strMethod, VbMethod)
    End If
    ' PSC 10/10/01 SYS2788 - Start
    lngErrNo = Err.Number
    strErrDesc = Err.Description
    Set objBO = Nothing
    Set xmlParmeterList = Nothing
    Set xmlParameter = Nothing
    On Error GoTo 0
    If lngErrNo <> 0 Then
        Err.Raise _
            lngErrNo, _
            "omCOG", _
            "omCOG OPI method(" & strProgId & "." & strMethod & ") " & strErrDesc
    End If
    If IsEmpty(varResponse) Then
        Err.Raise _
            oeNotImplemented, _
            cstrFunctionName, _
            "omCOG OPI method(" & strProgId & "." & strMethod & ") no response"
    End If
    vxmlPrimaryResponseDoc.setProperty "NewParser", True
    vxmlPrimaryResponseDoc.validateOnParse = False
    vxmlPrimaryResponseDoc.loadXML varResponse
End Sub
    
Private Sub ConvertResponse( _
    ByVal vxmlPrimaryResponseNode As IXMLDOMNode, _
    ByVal vxmlResponseDoc As FreeThreadedDOMDocument40)
    Dim xmlResponseNode As IXMLDOMNode, _
        xmlChildNode As IXMLDOMNode
    Set xmlResponseNode = _
        vxmlResponseDoc.appendChild(vxmlPrimaryResponseNode.cloneNode(False))
    For Each xmlChildNode In vxmlPrimaryResponseNode.childNodes
        xmlToAttrib xmlChildNode, xmlResponseNode
    Next
End Sub
Private Sub xmlToElement( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode)
    Dim xmlThisElem As IXMLDOMElement
    Dim xmlThisParentNode As IXMLDOMNode, _
        xmlThisChildNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
        
    Set xmlThisElem = _
        vxmlDestNode.ownerDocument.createElement(vxmlSrceNode.nodeName)
    Set xmlThisParentNode = vxmlDestNode.appendChild(xmlThisElem)
    For Each xmlAttrib In vxmlSrceNode.Attributes
        
        Set xmlThisElem = _
            vxmlDestNode.ownerDocument.createElement(xmlAttrib.Name)
        xmlThisElem.Text = xmlAttrib.Text
        xmlThisParentNode.appendChild xmlThisElem
    Next
    For Each xmlThisChildNode In vxmlSrceNode.childNodes
        xmlToElement xmlThisChildNode, xmlThisParentNode
    Next
End Sub
Private Sub xmlToAttrib( _
    ByVal vxmlSrceNode As IXMLDOMNode, _
    ByVal vxmlDestNode As IXMLDOMNode)
    Dim xmlThisParentNode As IXMLDOMNode, _
        xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
        
    Set xmlThisParentNode = vxmlDestNode.appendChild(vxmlSrceNode.cloneNode(False))
    For Each xmlNode In vxmlSrceNode.childNodes
        If xmlNode.hasChildNodes Then
            If xmlNode.childNodes.length = 1 And _
                xmlNode.firstChild.nodeType = NODE_TEXT _
            Then
                Set xmlAttrib = _
                    vxmlDestNode.ownerDocument.createAttribute(xmlNode.nodeName)
                xmlAttrib.Text = xmlNode.Text
                xmlThisParentNode.Attributes.setNamedItem xmlAttrib
            Else
                xmlToAttrib xmlNode, xmlThisParentNode
            End If
        Else
            If xmlNode.nodeType = NODE_TEXT Then
                
                Set xmlAttrib = _
                    vxmlDestNode.ownerDocument.createAttribute("VALUE")
                xmlAttrib.Text = xmlNode.xml
                xmlThisParentNode.Attributes.setNamedItem xmlAttrib
            ElseIf xmlNode.Attributes.length > 0 Then
                
                xmlToAttrib xmlNode, xmlThisParentNode
            End If
        End If
    Next
End Sub
' create attribute based ERROR response
' (could be added to errAssist)
Private Function CreateErrorResponse() As String
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    Set xmlElem = xmlDoc.createElement("RESPONSE")
    xmlElem.setAttribute "TYPE", "APPERR"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    Set xmlElem = xmlDoc.createElement("ERROR")
    xmlElem.setAttribute "NUMBER", Err.Number
    xmlElem.setAttribute "SOURCE", Err.Source
    xmlElem.setAttribute "DESCRIPTION", Err.Description
    xmlElem.setAttribute "VERSION", App.Comments
    xmlNode.appendChild xmlElem
    CreateErrorResponse = xmlDoc.xml
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
End Function
Private Sub Terminate( _
    ByVal vxmlDoc As FreeThreadedDOMDocument40, _
    ByVal vstrFunctionName As String, _
    Optional ByVal vblnAbortTransaction As Boolean = False)
    vxmlDoc.save ("c:\temp\omCOGRequest.xml")
    Shell _
        "c:\program files\plus!\microsoft internet\iexplore.exe c:\temp\omCOGRequest.xml", _
        vbMaximizedFocus
    If vblnAbortTransaction Then
        gobjContext.SetAbort
    End If
    Err.Raise _
        oeUnspecifiedError, _
        TypeName(Me) & "." & vstrFunctionName, _
        "omCOGRequest Terminate Request"
End Sub
Private Function CallByNameEx(objObject As Object, _
                              strProcName As String, _
                              vctCallType As VbCallType, _
                              Optional vArgsArray As Variant) As Variant
    Dim objTLI As Object
    Dim lngProcId As Long
    Dim lngNumArgs As Long
    Dim lngIndex As Long
    Dim varParameters() As Variant
    Dim varResponse As Variant
    On Error GoTo CallByNameExVbErr
    Set objTLI = CreateObject("TLI.TLIApplication")
    lngProcId = objTLI.InvokeID(objObject, strProcName)
    If IsMissing(vArgsArray) Then
        varResponse = objTLI.InvokeHook(objObject, lngProcId, vctCallType)
    End If
    If IsArray(vArgsArray) Then
        lngNumArgs = UBound(vArgsArray)
        ReDim varParameters(lngNumArgs)
        For lngIndex = 0 To lngNumArgs
            varParameters(lngIndex) = vArgsArray(lngNumArgs - lngIndex)
        Next lngIndex
        varResponse = objTLI.InvokeHookArray(objObject, lngProcId, vctCallType, varParameters)
          
    End If
    CallByNameEx = varResponse
    Exit Function
CallByNameExVbErr:
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
