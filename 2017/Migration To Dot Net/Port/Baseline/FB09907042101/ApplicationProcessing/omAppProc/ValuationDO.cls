VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ValuationDO"
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
' PSC       13/03/01 SYS1839 Amend GetValuerReport to get the lates one if sequence number not
'                            passed in
' DJP       13/03/01 SYS1839 Add RecordNotFound GetValuationReport
' DJP       13/03/01 SYS1839 Added new version of GetValuerInstructions from GD
' PSC       29/03/01 SYS2188 Add processing for batch payment
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IValuationDO

Private gobjContext As ObjectContext

'--------------------------------------------------------------------------------------
'BMIDS Specific History
'
'Prog   Date        AQR         Description
'DPF    03/09/2002  BMIDS00344  Added new method GetValuationReportForTemplate
'MV     03/09/2002  BMIDS00353  Added New Method IValuationDO_GetLatestValuationAmounts()
'--------------------------------------------------------------------------------------

Private Sub IValuationDO_GetValuationReportForTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

'------------------------------------------------------------------------------------
'This method checks for existing Valuation Reports for an application, if one exists
'details are pulled back
'------------------------------------------------------------------------------------

On Error GoTo GetValuationReportForTemplateExit
    
    Const strFunctionName As String = "IValuationDO_GetValuationReportForTemplate"
    
    Dim xmlViewElement As IXMLDOMElement
    Dim xmlReportElement As IXMLDOMElement
    Dim xmlAttribute As IXMLDOMAttribute
    
    Dim strSequenceNo As String
    
    strSequenceNo = xmlGetAttributeText(vxmlRequestNode, "INSTRUCTIONSEQUENCENO")
    
    If Len(strSequenceNo) > 0 Then
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETVALUATIONREPORT"
    Else
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETVALUATIONREPORTNOSEQNO"
    End If
        
GetValuationReportForTemplateExit:
    
    Set xmlViewElement = Nothing
    Set xmlReportElement = Nothing
    Set xmlAttribute = Nothing
    
    gobjContext.SetComplete
    
    'stop an error check taking place if no records were found
    If vxmlResponseNode.hasChildNodes = False Then
        Exit Sub
    End If
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_UpdateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdateValuationReportExit
    
    Const strFunctionName As String = "IValuationDO_UpdateValuationReport"
    
    adoUpdateFromNode vxmlRequestNode, "CREATEVALUATIONREPORTSUMMARY"
    adoUpdateFromNode vxmlRequestNode, "CREATEVALUATIONREPORTVALUATION"
    adoUpdateFromNode vxmlRequestNode, "CREATEVALUATIONREPORTPROPERTYRISKS"
    adoUpdateFromNode vxmlRequestNode, "CREATEVALUATIONREPORTPROPERTYSERVICES"
    adoUpdateFromNode vxmlRequestNode, "CREATEVALUATIONREPORTPROPERTYDETAILS"
    
UpdateValuationReportExit:
    
    gobjContext.SetComplete
    
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

Private Sub IValuationDO_CreateValuationReport( _
    ByVal vxmlRequestNodeList As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateValuationReportExit
    
    Const strFunctionName As String = "IValuationDO_CreateValuationReport"
    
    ' Five tables to update
    adoCreateFromNode vxmlRequestNodeList, "CREATEVALUATIONREPORTSUMMARY"
    adoCreateFromNode vxmlRequestNodeList, "CREATEVALUATIONREPORTVALUATION"
    adoCreateFromNode vxmlRequestNodeList, "CREATEVALUATIONREPORTPROPERTYRISKS"
    adoCreateFromNode vxmlRequestNodeList, "CREATEVALUATIONREPORTPROPERTYSERVICES"
    adoCreateFromNode vxmlRequestNodeList, "CREATEVALUATIONREPORTPROPERTYDETAILS"
    
CreateValuationReportExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_GetValuationReport( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetValuationReportExit
    
    Const strFunctionName As String = "IValuationDO_GetValuationReport"
    
    Dim xmlViewElement As IXMLDOMElement
    Dim xmlReportElement As IXMLDOMElement
    Dim xmlAttribute As IXMLDOMAttribute
    
    Dim strSequenceNo As String
    
    strSequenceNo = xmlGetAttributeText(vxmlRequestNode, "INSTRUCTIONSEQUENCENO")
    
    If Len(strSequenceNo) > 0 Then
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETVALUATIONREPORT"
    Else
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETVALUATIONREPORTNOSEQNO"
    End If
    
    ' SYS1839 Add RecordNotFound
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetValuationReportExit:
    
    Set xmlViewElement = Nothing
    Set xmlReportElement = Nothing
    Set xmlAttribute = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_GetValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    Dim xmlTempAttribute As IXMLDOMAttribute

    
    On Error GoTo GetValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationDO_GetValuerInstructions"
    
    'Check for presence of INSTRUCTIONSEQUENCENUMBER in REQUEST
    
    Set xmlTempAttribute = vxmlRequestNode.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO")
    
    If Not (xmlTempAttribute Is Nothing) Then 'ie.INSTRUCTIONSEQUENCENO tag present
        If Len(xmlTempAttribute.nodeValue) > 0 Then 'ie.INSTRUCTIONSEQUENCENO has a value
            'CALL adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTIONS"
            adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTIONS"
        Else ' in place, but no value (empty string)
            'REMOVE and CALL adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTIONSNOSEQNO"
            vxmlRequestNode.Attributes.removeNamedItem ("INSTRUCTIONSEQUENCENO")
            adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTIONSNOSEQNO"
        End If
    Else 'INSTRUCTIONSEQUENCENO  NOT present
        adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTIONSNOSEQNO"
    End If

    
GetValuerInstructionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_UpdateValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationDO_UpdateValuerInstructions"
    
    'GD TO CHANGE will this go in BO ??
    'If (vxmlRequestNode.Attributes.getNamedItem("APPLICATIONNUMBER") Is Nothing) Or _
     '   (vxmlRequestNode.Attributes.getNamedItem("APPLICATIONFACTFINDNUMBER") Is Nothing) Or _
     '   (vxmlRequestNode.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO") Is Nothing) Then
            'Report Error
            'From MsgTm AddCaseActivityGuid vxmlRequestNode
            
    'End If
    'GD TO CHANGE
    adoUpdateFromNode vxmlRequestNode, "VALUERINSTRUCTION"
    
UpdateValuerInstructionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_CreateValuerInstructions( _
    ByVal vxmlRequestNodeList As IXMLDOMNodeList, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationDO_CreateValuerInstructions"
    
   adoCreateFromNodeList vxmlRequestNodeList, "VALUERINSTRUCTION"

CreateValuerInstructionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_FindValuerInstructionList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo FindValuerInstructionListExit
    
    Const strFunctionName As String = "IValuationDO_FindValuerInstructionList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTION_SRT"
    
FindValuerInstructionListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_GetValuationFee( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetValuationFeeExit
    
    Const strFunctionName As String = "IValuationDO_GetValuationFee"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "VALUERINSTRUCTION"
    
GetValuationFeeExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_GetValuationPropertyDetails( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetValuationPropertyDetailsExit
    
    Const strFunctionName As String = "IValuationDO_GetValuationPropertyDetails"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "PROPERTY"
    
GetValuationPropertyDetailsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_FindInstructionsForBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Finds all valuation instructions that were received prior to the passed in number of
'   days before today
' pass:
'   vxmlRequestNode
'       XML Request
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo FindInstructionsForBatchExit
    
    Const strFunctionName As String = "IValuationDO_FindInstructionsForBatch"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "FINDINSTRUCTIONSFORBATCH"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

    
FindInstructionsForBatchExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_CreateValuationAudit( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates a valuation batch audit record
' pass:
'   vxmlRequestNode
'       Node to process
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo CreateValuationAuditExit
    
    Const strFunctionName As String = "IValuationDO_CreateValuationAudit"
    
   adoCreateFromNode vxmlRequestNode, "VALUATIONBATCHAUDIT"

CreateValuationAuditExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_GetLatestValuationAmounts(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                            ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetLatestValuationAmountsExit
    
    Const strFunctionName As String = "IValuationDO_GetLatestValuationAmounts"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETLATESTVALUATIONAMOUNTS"
    
    
    
GetLatestValuationAmountsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub IValuationDO_FindFailedInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Finds all valuation instructions that were received prior to the passed in number of
'   days before today
' pass:
'   vxmlRequestNode
'       XML Request
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo FindFailedInstructionsExit
    
    Const strFunctionName As String = "IValuationDO_FindFailedInstructions"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "FINDFAILEDINSTRUCTIONS"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

    
FindFailedInstructionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_FindTransactionSummaryList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Finds the summary for each valuer that had successful transactions processed for producing
'   output letters
' pass:
'   vxmlRequestNode
'       Node to process
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo FindTransactionSummaryListExit
    
    Const strFunctionName As String = "IValuationDO_FindTransactionSummaryList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TRANSACTIONSUMMARY"
    
FindTransactionSummaryListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_FindSuccessfulTransactionList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Finds the successful transactions processed for producing output letters
' pass:
'   vxmlRequestNode
'       Node to process
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------

    On Error GoTo FindSuccessfulTransactionListExit
    
    Const strFunctionName As String = "IValuationDO_FindSuccessfulTransactionList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "SUCCESSFULTRANSACTIONS"
    
FindSuccessfulTransactionListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationDO_FindTransactionList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Finds the all transactions regardless of status
' pass:
'   vxmlRequestNode
'       Node to process
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------

    On Error GoTo FindTransactionListExit
    
    Const strFunctionName As String = "IValuationDO_FindTransactionList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ALLTRANSACTIONS"
    
FindTransactionListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

