VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "FileVersioningBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'header ----------------------------------------------------------------------------------
'Workfile:      FileVersioningBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Helper functions for error handling
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'LD     26/01/01    Initial creation
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Private gobjContext As ObjectContext
Private gobjTxBO As IFileVersioningTxBO
Private gobjDO As IFileVersioningDO
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    ' ik_BM0200 traceassist support
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Public Function OmRequest(ByVal vstrXMLRequest As String) As String
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
    Const cstrFunctionName As String = "OmRequest"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
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
    gobjContext.SetComplete
    ' ensure global objects released before running request
    Set gobjTxBO = Nothing
    Set gobjDO = Nothing
    ' create REQUEST & RESPONSE FreeThreadedDOMDocument40s
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlInDoc.async = False
    xmlOutDoc.async = False
    ' create RESPONSE element
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    ' load XML REQUEST
    Debug.Print vstrXMLRequest
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXMLRequest
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
    xmlInDoc.save sPath & "\omFVSRequest_" & tname & ".xml"
    Set fso = Nothing
#End If
    
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
    OmRequest = xmlOutDoc.xml
    Debug.Print OmRequest
'   IK_DEBUG
#If IK_DEBUG Then
    xmlOutDoc.save sPath & "\omFVSResponse_" & tname & ".xml"
#End If
    
OmRequestExit:
    
    ' clean up ================================================================
    ' release global objects
    Set gobjTxBO = Nothing
    Set gobjDO = Nothing
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
    ' ik_BM0200 traceassist support
    gobjTrace.TraceResponse OmRequest
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
OmRequestVbErr:
    
    ' create ERROR response block from Err object
    OmRequest = errCreateErrorResponse()
    ' ik_BM0200 traceassist support
    gobjTrace.TraceIdentErrorResponse OmRequest
    Debug.Print OmRequest
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
    Const cstrFunctionName As String = "DoRequest"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim strOperation As String
    If vxmlRequestNode.nodeName = "REQUEST" Then
        ' get operation (function) name from REQUEST/@OPERATION
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        ' get operation (function) name from OPERATION/@NAME
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    strOperation = UCase(strOperation)
    Select Case strOperation
        Case _
            "CHECKOUTFILE", _
            "CHECKINFILE", _
            "CREATEFILE", _
            "CREATEFOLDER", _
            "UNLOCKFILE", _
            "SHAREFILE", _
            "GETFILE"
            CallTxBO strOperation, vxmlRequestNode, vxmlResponseNode
        Case _
            "GETFILES", _
            "GETFOLDERS"
            CallDO strOperation, vxmlRequestNode, vxmlResponseNode
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
DoRequestExit:
    
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CallTxBO( _
    ByVal strOperation As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Calls Operation in the TxBO
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
    On Error GoTo CallTxBOExit
    Const cstrFunctionName As String = "CallTxBO"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningTxBO")
    End If
    gobjTxBO.CallTxBO strOperation, vxmlRequestNode, vxmlResponseNode
CallTxBOExit:
    
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CallDO( _
    ByVal strOperation As String, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Calls Operation in the DO
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
    On Error GoTo CallDOExit
    Const cstrFunctionName As String = "CallDO"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(gstrFVS_COMPONENT & ".FileVersioningDO")
    End If
    Select Case strOperation
        Case "GETFILES"
            gobjDO.GetFiles vxmlRequestNode, vxmlResponseNode
        Case "GETFOLDERS"
            gobjDO.GetFolders vxmlRequestNode, vxmlResponseNode
        Case Else
            errThrowError _
                "FileVersioningRequest", _
                oeNotImplemented, _
                strOperation
    End Select
CallDOExit:
    
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const cstrFunctionName As String = "DoClientRequest"
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    '<VSA> Visual Studio Analyser Support
    Select Case strOperation
    'Enter Customer specific case details here
    Case Else
        errThrowError _
        cstrFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select
DoClientRequestExit:
         
    
    ' ik_BM0200 traceassist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Function
' ik_BM0200 traceassist support
Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStartElapsed As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStartElapsed
End Sub
