VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ValuationTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Component to link cost modelling to AiP
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog      Date     Description
' DJP/GD    12/03/01 Valuer Instructions (GD) and Valuation Report (DP)
' PSC       03/04/01 SYS2188 Add processing for batch payment
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IValuationTxBO

Private gobjContext As ObjectContext

Private Sub IValuationTxBO_ProcessValuationInstructions(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Private Sub IValuationTxBO_UpdateValuationReport(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateValuationReportVbErr
    
    Const strFunctionName As String = "IValuationTxBO_UpdateValuationReport"
    
    Dim xmlValuationList As IXMLDOMNodeList
    Dim xmlValuationNode As IXMLDOMNode
    
    Dim objDO As IValuationDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
    
    Set xmlValuationList = vxmlRequestParentNode.selectNodes("VALUATION")
    
    For Each xmlValuationNode In xmlValuationList
        objDO.UpdateValuationReport xmlValuationNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateValuationReportExit:
    
    Set objDO = Nothing
    Set xmlValuationList = Nothing
    Set xmlValuationNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateValuationReportVbErr:

    gobjContext.SetAbort

    GoTo UpdateValuationReportExit

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

Private Sub IValuationTxBO_CreateValuationReport( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CASETASK entities.
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

    On Error GoTo CreateValuationReportVbErr
    
    Const strFunctionName As String = "IValuationTxBO_CreateValuationReport"
    
    Dim xmlValNode As IXMLDOMNode
    Dim xmlValuationList As IXMLDOMNodeList
    Dim objDO As IValuationDO
    
    Set xmlValuationList = vxmlRequestParentNode.selectNodes("VALUATION")
        
    If xmlValuationList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
        For Each xmlValNode In xmlValuationList
            objDO.CreateValuationReport xmlValNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateValuationReportExit:
    
    Set objDO = Nothing
    Set xmlValuationList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateValuationReportVbErr:

    gobjContext.SetAbort
    
    GoTo CreateValuationReportExit

End Sub

Private Sub IValuationTxBO_UpdateValuerInstructions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateValuerInstructionsVbErr
    
    Const strFunctionName As String = "IValuationTxBO_UpdateValuerInstructions"
    
    Dim xmlValuerInstructionsList As IXMLDOMNodeList
    Dim xmlValuerInstructionsNode As IXMLDOMNode
    
    Dim objDO As IValuationDO
    
    Set objDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")

    Set xmlValuerInstructionsList = vxmlRequestParentNode.selectNodes("VALUERINSTRUCTION")
    
    For Each xmlValuerInstructionsNode In xmlValuerInstructionsList
        objDO.UpdateValuerInstructions xmlValuerInstructionsNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateValuerInstructionsExit:
    
    Set objDO = Nothing
    Set xmlValuerInstructionsList = Nothing
    Set xmlValuerInstructionsNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateValuerInstructionsVbErr:

    gobjContext.SetAbort

    GoTo UpdateValuerInstructionsExit

End Sub

Private Sub IValuationTxBO_CreateValuerInstructions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified ValuerInstruction entities.
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

    On Error GoTo CreateValuerInstructionsVbErr
    
    Const strFunctionName As String = "IValuationTxBO_CreateValuerInstructions"
    
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim objDO As IValuationDO
    
    Set xmlCaseTaskList = vxmlRequestParentNode.selectNodes("VALUERINSTRUCTION")
        
    If xmlCaseTaskList.length > 0 Then
        ' call CreateValuerInstructions if ValuerInstruction node in request
        Set objDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
        objDO.CreateValuerInstructions xmlCaseTaskList, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
    
CreateValuerInstructionsExit:
    
    Set objDO = Nothing
    Set xmlCaseTaskList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateValuerInstructionsVbErr:

    gobjContext.SetAbort
    
    GoTo CreateValuerInstructionsExit

End Sub

Private Sub IValuationTxBO_CreateValuationAudit( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified Valuation batch audit record.
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

    On Error GoTo CreateValuationAuditVbErr
    
    Const strFunctionName As String = "IValuationTxBO_CreateValuationAudit"
    
    Dim xmlValuationAuditList As IXMLDOMNodeList
    Dim xmlValuationAudit As IXMLDOMElement
    Dim objDO As IValuationDO
    
    Set xmlValuationAuditList = vxmlRequestParentNode.selectNodes("VALUATIONBATCHAUDIT")
        
    If xmlValuationAuditList.length > 0 Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
        
        For Each xmlValuationAudit In xmlValuationAuditList
            objDO.CreateValuationAudit xmlValuationAudit, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateValuationAuditExit:
    
    Set objDO = Nothing
    Set xmlValuationAuditList = Nothing
    Set xmlValuationAudit = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateValuationAuditVbErr:

    gobjContext.SetAbort
    
    GoTo CreateValuationAuditExit

End Sub



