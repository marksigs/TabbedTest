VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "omPrintBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog   Date        Description
'MV     31/01/2005  E2EM00003299 - Performace Enhancements
'MV     01/02/2005  E2EM00003299 - Variable Errors
'MV     02/02/2005  BBG1908 - Amended FindDocumentNameList()
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Private gobjContext As ObjectContext
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    'IK_BM0200 traceAssist support
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Public Function omRequest(ByVal vstrXMLRequest As String) As String
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
    On Error GoTo omRequestVbErr
    Const cstrFunctionName As String = "omRequest"
    'IK_BM0200 traceAssist support
    gobjTrace.TraceRequest vstrXMLRequest
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    'IK_BM0200 traceAssist support ends
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
    xmlInDoc.save sPath & "\omPrintRequest_" & tname & ".xml"
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
    omRequest = xmlOutDoc.xml
    Debug.Print omRequest
omRequestExit:
    
    ' clean up ================================================================
'   IK_DEBUG
#If IK_DEBUG Then
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.loadXML omRequest
    xmlResponseDoc.save sPath & "\omPrintResponse_" & tname & ".xml"
    Set xmlResponseDoc = Nothing
#End If
    
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
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceResponse omRequest
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    Exit Function
omRequestVbErr:
    
    ' create ERROR response block from Err object
    omRequest = errCreateErrorResponse()
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceIdentErrorResponse omRequest
    Debug.Print omRequest
    ' perform clean up
    GoTo omRequestExit
End Function
Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
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
    ' IK_BM0200 traceAssist support
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
        Case "CREATECASETASK"
            CreateCaseTask vxmlRequestNode, vxmlResponseNode
        Case "GETCASETASKS"
            GetCaseTasks vxmlRequestNode, vxmlResponseNode
        Case "UPDATECASETASK"
            UpdateCaseTask vxmlRequestNode, vxmlResponseNode
        Case "FINDDOCUMENTNAMELIST"
            FindDocumentNameList vxmlRequestNode, vxmlResponseNode
        Case "GETPRINTATTRIBUTES"
            GetPrintAttributes vxmlRequestNode, vxmlResponseNode
        Case "PRINTDOCUMENT"
            PrintDocument vxmlRequestNode, vxmlResponseNode
        ' ik_BMIDS00885
        Case "REPRINTDOCUMENT"
            RePrintDocument vxmlRequestNode, vxmlResponseNode
        ' ik_BMIDS00885_ends
        'SYS4287 Rate Change Letters
        Case "PRINTLETTERS"
            PrintLetters vxmlRequestNode, vxmlResponseNode
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
DoRequestExit:
    
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CreateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseTaskExit
    Const cstrFunctionName As String = "CreateCaseTask"
    Dim objOmPrintTxBO As IOmPrintTxBO
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    
    If Not gobjContext Is Nothing Then
        Set objOmPrintTxBO = gobjContext.CreateInstance("OmPrint.OmPrintTxBO")
    Else
        Set objOmPrintTxBO = New OmPrintTxBO
    End If
   
    objOmPrintTxBO.CreateCaseTask vxmlRequestNode, vxmlResponseNode
CreateCaseTaskExit:
    
    Set objOmPrintTxBO = Nothing
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub GetCaseTasks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified CASETASK entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CASETASK node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
    On Error GoTo GetCaseTasksExit
    Const cstrFunctionName As String = "GetCaseTasks"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim objomPrintDO As IOmPrintDO
    
    ' validate REQUEST & select CASETASK node
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    If Not gobjContext Is Nothing Then
        Set objomPrintDO = gobjContext.CreateInstance("OmPrint.OmPrintDO")
    Else
        Set objomPrintDO = New OmPrintDO
    End If
    
    objomPrintDO.GetCaseTasksDetail xmlCaseTaskNode, vxmlResponseNode
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
GetCaseTasksExit:
    Set objomPrintDO = Nothing
    Set xmlCaseTaskNode = Nothing
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub UpdateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Private Sub FindDocumentNameList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo FindDocumentNameListExit
    Const cstrFunctionName As String = "FindDocumentNameList"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    Dim objomPrintDO As IOmPrintDO
    Dim xmlEntityNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("FINDDOCUMENTS")
    
    If Not gobjContext Is Nothing Then
        Set objomPrintDO = gobjContext.CreateInstance("omPrint.omPrintDO")
    Else
        Set objomPrintDO = New OmPrintDO
    End If
    
    objomPrintDO.FindDocumentNameList xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
FindDocumentNameListExit:
    
    Set objomPrintDO = Nothing
    Set xmlEntityNode = Nothing
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub GetPrintAttributes(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetPrintAttributesExit
    
    Const cstrFunctionName As String = "GetPrintAttributes"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    Dim xmlEntityNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objomPrintDO As IOmPrintDO
    
    Set xmlEntityNode = vxmlRequestNode.selectSingleNode("FINDATTRIBUTES")
    
    If Not gobjContext Is Nothing Then
        Set objomPrintDO = gobjContext.CreateInstance("OmPrint.OmPrintDO")
    Else
        Set objomPrintDO = New OmPrintDO
    End If
    
    objomPrintDO.GetPrintAttributes xmlEntityNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
GetPrintAttributesExit:
    
    Set objomPrintDO = Nothing
    Set xmlEntityNode = Nothing
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'Private Sub FindRecipientList( _
'    ByVal vxmlRequestNode As IXMLDOMNode, _
'    ByVal vxmlResponseNode As IXMLDOMNode)
'
'On Error GoTo FindRecipientListExit
'
'    Const cstrFunctionName As String = "FindRecipientList"
'
'    ' <VSA> Visual Studio Analyser Support
'    #If USING_VSA Then
'        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & cstrFunctionName)
'    #End If
'
'
'    Dim xmlTempResponseDoc As FreeThreadedDomDocument40
'    Set xmlTempResponseDoc = New FreeThreadedDomDocument40
'
'    'Get Queue parameter
'    Dim objAIPBO As AgreementInPrincipleBO
'    Dim strXMLResponse As String, strQueueName As String
'    Dim xmlDoc As FreeThreadedDomDocument40
'
'    If objAIPBO Is Nothing Then
'        Set objAIPBO = gobjContext.CreateInstance("omAIP.AgreementInPrincipleBO")
'    End If
'
'    'Call Print Data Manager and load response string into a FreeThreadedDomDocument40
'    xmlTempResponseDoc.loadXML objAIPBO.FindThirdPartyForCustomer(vxmlRequestNode.xml)
'    Debug.Print xmlTempResponseDoc.xml
'
'    'CheckResponse for errors
'    errCheckXMLResponse xmlTempResponseDoc.xml, True, vxmlResponseNode
'
'FindRecipientListExit:
'
'End Sub
Private Sub PrintDocument(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo PrintDocumentExit
    
    Const cstrFunctionName As String = "PrintDocument"
    Dim objOmPrintTxBO As IOmPrintTxBO
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    
    If Not gobjContext Is Nothing Then
        Set objOmPrintTxBO = gobjContext.CreateInstance(App.Title & ".omPrintTxBO")
    Else
        Set objOmPrintTxBO = New OmPrintTxBO
    End If
    
    objOmPrintTxBO.PrintDocument vxmlRequestNode, vxmlResponseNode
PrintDocumentExit:
    
    Set objOmPrintTxBO = Nothing
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
' ik_BMIDS00885
Private Sub RePrintDocument(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo RePrintDocumentExit
    Dim objOmPrintTxBO  As IOmPrintTxBO

    Const cstrFunctionName As String = "RePrintDocument"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If Not gobjContext Is Nothing Then
        Set objOmPrintTxBO = gobjContext.CreateInstance(App.Title & ".omPrintTxBO")
    Else
        Set objOmPrintTxBO = New OmPrintTxBO
    End If
    objOmPrintTxBO.RePrintDocument vxmlRequestNode, vxmlResponseNode
RePrintDocumentExit:
    
    Set objOmPrintTxBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
' ik_BMIDS00885_ends
'SYS4287 Used for Rate Change Letters from Batch Processing
Private Sub PrintLetters(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo PrintLettersExit
    
    Const cstrFunctionName As String = "PrintLetters"
    Dim objOmPrintTxBO  As IOmPrintTxBO
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    If Not gobjContext Is Nothing Then
        Set objOmPrintTxBO = gobjContext.CreateInstance(App.Title & ".omPrintTxBO")
    Else
        Set objOmPrintTxBO = New OmPrintTxBO
    End If
    objOmPrintTxBO.PrintLetters vxmlRequestNode, vxmlResponseNode
PrintLettersExit:
    
    Set objOmPrintTxBO = Nothing
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const cstrFunctionName As String = "DoClientRequest"
    ' IK_BM0200 traceAssist support
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
         
    
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodExit TypeName(Me), cstrFunctionName
    errCheckError cstrFunctionName, TypeName(Me)
End Function
