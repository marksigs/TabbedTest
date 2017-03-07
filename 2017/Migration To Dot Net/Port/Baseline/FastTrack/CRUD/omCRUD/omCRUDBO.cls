VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omCRUDBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------------------------
' Procedure OmRequest
' Author:           IK
' Date:             24/02/2004
' Purpose:          The only public method and therefore single entry point of the
'                   class, from here the requested operation is brokered.
' Input parameters: vstrXmlIn - an XML String
' Output parameters: OmRequest returns an XML response string
'------------------------------------------------------------------------
Public Function OmRequest(ByVal vstrXmlIn As String) As String

    On Error GoTo OmRequestVbErr
    
    Const strFunctionName As String = "OmRequest"

    Dim xmlInDoc As DOMDocument40
    Dim xmlOutDoc As DOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlRequestClone As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlAttrib As IXMLDOMAttribute
    
    GetObjectContext.SetComplete
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New DOMDocument40
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
'   ik_wip_20050114 - will be overriden as required
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'   ik_wip_20050114 - ends
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    Debug.Print vstrXmlIn
    
    xmlInDoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    Set xmlRequestClone = xmlRequestNode.cloneNode(False)
    
    adoCRUDSaveRequest xmlInDoc, xmlRequestClone
    
    Set xmlRequestNode = PreProcInterface(xmlInDoc)
    
    If Not xmlRequestNode Is Nothing Then
    
        If xmlInDoc.selectNodes("REQUEST/OPERATION").length = 0 Then
            DoRequest xmlRequestNode, xmlResponseNode
        Else
            Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
            For Each xmlOperationNode In xmlOperationNodeList
                Dim xmlOpResponse As IXMLDOMNode
                Set xmlNode = xmlResponseNode.appendChild(xmlOperationNode.cloneNode(False))
                Set xmlElem = xmlOutDoc.createElement("RESPONSE")
                Set xmlOpResponse = xmlNode.appendChild(xmlElem)
                For Each xmlAttrib In xmlRequestNode.Attributes
                    xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
                Next
                If Not xmlOperationNode.Attributes.getNamedItem("ProgId") Is Nothing Then
                    NamedInterface xmlOperationNode, xmlOpResponse
                Else
                    DoRequest xmlOperationNode, xmlOpResponse
                End If
                Set xmlOpResponse = Nothing
            Next
        End If
    
        Set xmlResponseNode = xmlOutDoc.selectSingleNode("RESPONSE")
        
        If xmlGetAttributeAsBoolean(xmlRequestNode, "RETURNREQUEST") Then
            xmlResponseNode.appendChild xmlRequestNode.cloneNode(True)
        End If
        
        PostProcInterface xmlInDoc, xmlOutDoc
        
    Else
    
        xmlOutDoc.loadXML xmlInDoc.xml
        Set xmlResponseNode = xmlOutDoc.selectSingleNode("RESPONSE")
        
    End If
    
OmRequestExit:
    
    OmRequest = xmlOutDoc.xml
    Debug.Print OmRequest
    
    adoCRUDSaveResponse xmlOutDoc, xmlRequestClone
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestClone = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmRequestVbErr:
    
    Set xmlResponseElem = xmlOutDoc.selectSingleNode("RESPONSE")
    
    If xmlGetAttributeText(xmlResponseElem, "TYPE") = "SUCCESS" Then
        Err.Source = App.EXEName & "." & TypeName(Me) & "." & Err.Source
        xmlOutDoc.loadXML errCreateErrorResponse()
    Else
        Set xmlNode = xmlOutDoc.selectSingleNode("RESPONSE/ERROR/SOURCE")
        If Not xmlNode Is Nothing Then
            xmlNode.Text = App.EXEName & "." & TypeName(Me) & "." & xmlNode.Text
        End If
    End If
    
    GetObjectContext.SetAbort
    Debug.Print OmRequest
    
    GoTo OmRequestExit
    
End Function


