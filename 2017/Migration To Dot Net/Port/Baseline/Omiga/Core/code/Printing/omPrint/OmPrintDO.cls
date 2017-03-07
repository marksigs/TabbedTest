VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "OmPrintDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'IK     17/02/2003  BM0200 - add TraceAssist support
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IOmPrintDO
Private gobjContext As ObjectContext
Private Sub IOmPrintDO_FindDocumentNameList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo FindDocumentNameListExit
    
    ' ik_20030211
    gobjContext.SetComplete
    Const cstrFunctionName As String = "FindDocumentNameList"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "DOCUMENTS"
FindDocumentNameListExit:
    
    ' ik_20030211
    ' gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub IOmPrintDO_FindRecipientList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
End Sub
Private Sub IOmPrintDO_GetPrintAttributes( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetPrintAttributesExit
    
    ' ik_20030211
    gobjContext.SetComplete
    Const cstrFunctionName As String = "GetPrintAttributes"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ATTRIBUTES"
GetPrintAttributesExit:
    
    ' ik_20030211
    ' gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
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
' IOm4Ph2DO_CreateCaseTask header ---------------------------------------------------------
' description:
'   Creates specified CASETASK entities.
' pass:
'   vxmlRequestNodeList
'       list of XML CASETASK nodes from REQUEST
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
Private Sub IOmPrintDO_CreateCaseTask( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateCaseTaskExit
    ' ik_20030211
    gobjContext.SetComplete
    Const cstrFunctionName As String = "CreateCaseTask"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    adoCreateFromNodeList vxmlRequestNodeList, "CASETASK"
CreateCaseTaskExit:
    
    ' ik_20030211
    ' gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub IOmPrintDO_GetCaseTasksDetail( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetCaseTasksDetailExit
    ' ik_20030211
    gobjContext.SetComplete
    Const cstrFunctionName As String = "GetCaseTasksDetail"
    ' IK_BM0200 traceAssist support
    gobjTrace.TraceMethodEntry TypeName(Me), cstrFunctionName
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CASETASKDETAIL"
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
GetCaseTasksDetailExit:
    
    ' ik_20030211
    ' gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
