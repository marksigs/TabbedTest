VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omCaseTrackBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCaseTrackBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Case Tracking Entry Business Object
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    11/04/01    SYS1927 Created
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private m_objContext As ObjectContext

Private m_objCaseTrackingBO As ICaseTrackingBO

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Public Function OmCaseTrackRequest(ByVal vstrXMLRequest As String) As String
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

    On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmCaseTrackRequest"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' will be superceded by any subsequent SetAbort
    m_objContext.SetComplete

    ' ensure global objects released before running request
    Set m_objCaseTrackingBO = Nothing
    
    ' create REQUEST & RESPONSE DOMDocuments
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.loadXML vstrXMLRequest
    
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
    
    OmCaseTrackRequest = xmlOutDoc.xml
    Debug.Print OmCaseTrackRequest
    
OmRequestExit:
    
    ' clean up ================================================================
    
    ' release global objects
    Set m_objCaseTrackingBO = Nothing

    ' release XML DOM objects
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmRequestVbErr:
    
    ' create ERROR response block from Err object
    OmCaseTrackRequest = errCreateErrorResponse()
    Debug.Print OmCaseTrackRequest
    
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
    
    Dim strOperation As String
    
    If m_objCaseTrackingBO Is Nothing Then
        Set m_objCaseTrackingBO = m_objContext.CreateInstance(App.Title & ".CaseTrackingBO")
    End If
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "FINDAPPLICATIONSFORCASETRACKING"
            m_objCaseTrackingBO.FindApplicationsForCaseTracking vxmlRequestNode, vxmlResponseNode
            
        Case "GETCASETRACKINGALLDATA"
            m_objCaseTrackingBO.GetCaseTrackingAllData vxmlRequestNode, vxmlResponseNode
            
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    errCheckError strFunctionName, TypeName(Me)

End Function
