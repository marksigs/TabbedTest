VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omMQBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omMQBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Send a message to the Queue using XML string as a parameter.
'
'Dependencies:  omMSMQ, omOMMQ, omCOMP
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SF&PD  09/10/01    Created.
'AD     25/10/01    Rework how this component is structured. Will now use O4 Phase2 style
'                   BO, TxBO and DO to talk to the existing omMessageQueueClass. Means we
'                   we can have a component and still share the existing file so that
'                   compatibility is maintained.
'RF     01/03/02    SYS4215 omMQ.SendToQueue is producing "method ~ of object ~ failed"
'                   in release. Various improvements identified.
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'STB    08/07/02    SYS5040 Removed excessive auditing to the application event log.
'------------------------------------------------------------------------------------------
Option Explicit

'Implements IomMessageQueueBO
Implements ObjectControl

Private gobjContext As ObjectContext


' AD 29/10/01 New functions Omiplus 30

Public Function omRequest(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Public interface for this component.
' pass:
'   vstrXMLRequest
'       XML REQUEST details.
' return:
'       XML RESPONSE details (As String)
' exceptions:
'       none
'------------------------------------------------------------------------------------------
On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmRequest"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim xmlInDoc As DOMDocument
    Dim xmlOutDoc As DOMDocument
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' will be superceded by any subsequent SetAbort
    gobjContext.SetComplete
    
    ' create REQUEST & RESPONSE DOMDocuments
    Set xmlInDoc = New DOMDocument
    Set xmlOutDoc = New DOMDocument
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    Set xmlInDoc = xmlLoad(vstrXMLRequest, strFunctionName)
    
    ' check for request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        ' OPERATION attribute on REQUEST node
        ' do single operation
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        ' multiple operations
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                ' copy attributes from REQUEST node to OPERATION node
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If

    ' if we get here, everything is OK
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    omRequest = xmlOutDoc.xml
    Debug.Print omRequest
    
OmRequestExit:
    
    ' clean up ================================================================
    
    ' release global objects

    ' release XML DOM objects
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Exit Function

OmRequestVbErr:
    
    ' create ERROR response block from Err object
    omRequest = errCreateErrorResponse()
    Debug.Print omRequest
    
    ' perform clean up
    GoTo OmRequestExit

End Function


Private Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Request broker, private interface, performs single REQUEST OPERATION
'   Any results of operation (data retrieved, error info etc.) appended to vxmlResponseNode
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       raises oeNotImplemented if OPERATION value not in list of known functions.
'------------------------------------------------------------------------------------------
On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    adoBuildDbConnectionString
    
    Select Case strOperation
    
        Case "SENDTOQUEUE"
            SendToQueue vxmlRequestNode, vxmlResponseNode
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub SendToQueue( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified SendToQueue entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
On Error GoTo SendToQueueExit
    
    Const strFunctionName As String = "SendToQueue"
    
    'App.LogEvent "Start omMQBO " & strFunctionName, vbLogEventTypeInformation
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim objIMQTxBO As IomMQTxBO
    Dim nodeXML As IXMLDOMNode
    Dim nodeMESSAGEQUEUE As IXMLDOMNode
    Dim childNode As IXMLDOMNode
    Dim nodeTXBOResponse As IXMLDOMNode
    Dim strXML As String
    
    'DS Convert the contents of the XML node from an element to a string and make
    'it an attribute.
    
    Set nodeMESSAGEQUEUE = xmlGetMandatoryNode(vxmlRequestNode, "MESSAGEQUEUE")
    Set nodeXML = xmlGetMandatoryNode(nodeMESSAGEQUEUE, "XML")
    
    For Each childNode In nodeXML.childNodes
        strXML = strXML & childNode.xml
    Next childNode
    
    nodeMESSAGEQUEUE.removeChild nodeXML
    xmlSetAttributeValue nodeMESSAGEQUEUE, "XML", strXML
    
    If objIMQTxBO Is Nothing Then
        Set objIMQTxBO = gobjContext.CreateInstance(App.Title & ".omMQTxBO")
    End If
    
    Set nodeTXBOResponse = objIMQTxBO.SendToQueue(vxmlRequestNode)
    
    ' RF 11/01/02
    errCheckXMLResponseNode nodeTXBOResponse, , True
    
    Set vxmlResponseNode = nodeTXBOResponse
    
SendToQueueExit:
    
    'App.LogEvent "Exit omMQBO " & strFunctionName, vbLogEventTypeInformation
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    Set objIMQTxBO = Nothing
    Set nodeMESSAGEQUEUE = Nothing
    Set nodeXML = Nothing
    Set childNode = Nothing
    Set nodeTXBOResponse = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    '<VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName, TypeName(Me)

End Function
