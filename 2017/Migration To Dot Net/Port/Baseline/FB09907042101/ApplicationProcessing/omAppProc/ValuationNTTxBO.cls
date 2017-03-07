VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "ValuationNTTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ValuationNTTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Transactional Business Object that requires a new transaction. Performs the
'               updates for an instance of a valuer instruction.
'
'Dependencies:  List any other dependent components
'               e.g. Omiga4DO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    11/04/99    Created
'SG     25/02/02    SYS4097 IValuationNTTxBO_ProcessValuerBatch global parameter now passed in.
'PSC    25/02/02    SYS4097 Amend where queue name is retrieved from
'------------------------------------------------------------------------------------------

Option Explicit

Implements IValuationNTTxBO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub IValuationNTTxBO_ProcessValuationInstruction(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                                         ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ProcessValuationInstructionsVbErr
    
    Const strFunctionName As String = "IValuationNTTxBO_ProcessValuationInstruction"
    
    Dim objThirdPartyBO As ThirdPartyBO
    Dim objAppManagerBO As ApplicationManagerBO
    Dim objApplicationBO As ApplicationBO
    Dim objIValuationDO As IValuationDO
    Dim objomTMBO As OmTmBO
        
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlLockRequest As IXMLDOMNode
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlValuationNode As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim dblInvoiceAmount As Double
    Dim strResponse As String
    Dim lngErrorNo As Long

    Set xmlValuationNode = xmlGetMandatoryNode(vxmlRequestParentNode, "VALUATION")
    
    ' Check invoice amount
    dblInvoiceAmount = xmlGetMandatoryAttributeAsDouble(xmlValuationNode, "VALUERINVOICEAMOUNT")
    
    If dblInvoiceAmount < 0.01 Then
        
        ' Write error report
        errThrowError strFunctionName, 312
    
    End If
    
    ' Get payment details
    Set xmlTempRequest = vxmlRequestParentNode.cloneNode(False)
    Set xmlRootNode = vxmlRequestParentNode.ownerDocument.createElement("PAYMENTDETAILS")
    xmlTempRequest.appendChild xmlRootNode
    
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("DIRECTORYGUID")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlValuationNode, "DIRECTORYGUID")
    xmlRootNode.appendChild xmlNode
    
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    strResponse = objThirdPartyBO.GetPaymentDetails(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)
    
    ' Payment details not present
    If lngErrorNo <> 0 Then
        
        ' Write error report
        errThrowError strFunctionName, 309
    
    End If
    
    Dim strPaymentMethod As String
    Dim strCompanyName As String
    Dim strBankSortCode As String
    Dim strBankName As String
    Dim strAccountName As String
    Dim strAccountNumber As String
    
    ' Check if payment details are complete
    strPaymentMethod = xmlGetNodeText(xmlTempResponseDoc, ".//PAYMENTMETHOD")
    strCompanyName = xmlGetNodeText(xmlTempResponseDoc, ".//COMPANYNAME")
    strBankSortCode = xmlGetNodeText(xmlTempResponseDoc, ".//BANKSORTCODE")
    strBankName = xmlGetNodeText(xmlTempResponseDoc, ".//BANKNAME")
    strAccountName = xmlGetNodeText(xmlTempResponseDoc, ".//ACCOUNTNAME")
    strAccountNumber = xmlGetNodeText(xmlTempResponseDoc, ".//ACCOUNTNUMBER")
    
    ' Payment details incomplete
    If Len(strPaymentMethod) = 0 Or Len(strCompanyName) = 0 Or Len(strBankSortCode) = 0 Or _
       Len(strBankName) = 0 Or Len(strAccountName) = 0 Or Len(strAccountNumber) = 0 Then
       
       ' Write error report
        errThrowError strFunctionName, 309
       
    End If
    
    ' Check if application is already locked
    Set xmlRootNode = Nothing
    
    Set xmlLockRequest = vxmlRequestParentNode.cloneNode(False)
    Set xmlRootNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONLOCK")
    xmlLockRequest.appendChild xmlRootNode
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONNUMBER")
    xmlRootNode.appendChild xmlNode
    
    Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strResponse = objAppManagerBO.GetLockData(xmlLockRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)

    ' Application is already locked
    If errGetOmigaErrorNumber(lngErrorNo) <> oeRecordNotFound Then
    
        ' Write error report
        errThrowError strFunctionName, 310

    End If
    
    ' Lock application
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("USERID")
    xmlNode.Text = xmlGetMandatoryAttributeText(vxmlRequestParentNode, "USERID")
    xmlRootNode.appendChild xmlNode
    
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("UNITID")
    xmlNode.Text = xmlGetMandatoryAttributeText(vxmlRequestParentNode, "UNITID")
    xmlRootNode.appendChild xmlNode
    
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("MACHINEID")
    xmlNode.Text = xmlGetMandatoryAttributeText(vxmlRequestParentNode, "MACHINEID")
    xmlRootNode.appendChild xmlNode
    
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("CHANNELID")
    xmlNode.Text = xmlGetMandatoryAttributeText(vxmlRequestParentNode, "CHANNELID")
    xmlRootNode.appendChild xmlNode
    
    strResponse = objAppManagerBO.CreateLock(xmlLockRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)

    ' Can't lock application
    If lngErrorNo <> 0 Then
    
        ' Write error report
        errThrowError strFunctionName, 311

    End If
    
    ' Update the instruction with the instruction date
    Set objIValuationDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    
    
    Set xmlRootNode = Nothing
    Set xmlRootNode = vxmlRequestParentNode.ownerDocument.createElement("VALUATION")
    xmlCopyMandatoryAttribute xmlValuationNode, xmlRootNode, "APPLICATIONNUMBER"
    xmlCopyMandatoryAttribute xmlValuationNode, xmlRootNode, "APPLICATIONFACTFINDNUMBER"
    xmlCopyMandatoryAttribute xmlValuationNode, xmlRootNode, "INSTRUCTIONSEQUENCENO"
    xmlSetAttributeValue xmlRootNode, "DATEVALUATIONPAID", Format(Date, "dd/mm/yyyy")

    On Error Resume Next
    
    objIValuationDO.UpdateValuerInstructions xmlRootNode, vxmlResponseNode
    
    lngErrorNo = Err.Number
    
    On Error GoTo ProcessValuationInstructionsVbErr
    
    If lngErrorNo <> 0 Then
    
        ' Write error report
        errThrowError strFunctionName, 315
    
    End If

    ' Create Case Task
    Dim strActivityID As String
    Dim strTaskId As String
    Dim strStageId As String
    Dim strCaseStageSeqNo As String
    Dim strPriority As String
    Dim strErrDesc As String
    Dim strTaskStatus As String
    
    strActivityID = GetMandatoryGlobalParamString("TMOmigaActivity")
    strTaskId = GetMandatoryGlobalParamString("TMValuationFeePaid")
    strTaskStatus = GetFirstComboValueId("TaskStatus", "CP")
    
    Set xmlTempRequest = vxmlRequestParentNode.cloneNode(False)
    Set xmlRootNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONSTAGE")
    xmlTempRequest.appendChild xmlRootNode
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONNUMBER")
    xmlRootNode.appendChild xmlNode
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONFACTFINDNUMBER")
    xmlRootNode.appendChild xmlNode
    
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    strResponse = objApplicationBO.GetCurrentApplicationStage(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)

    If lngErrorNo <> 0 Then
    
        ' Write error report
        strErrDesc = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//DESCRIPTION")
        errThrowError strFunctionName, 313, strErrDesc
        errThrowError strFunctionName, 313
    
    End If
   
    strStageId = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//STAGENUMBER")
    strCaseStageSeqNo = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//STAGESEQUENCENO")
    
    Set xmlTempRequest = vxmlRequestParentNode.cloneNode(False)
    Set xmlRootNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONPRIORITY")
    xmlTempRequest.appendChild xmlRootNode
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONNUMBER")
    xmlRootNode.appendChild xmlNode
    
    strResponse = objApplicationBO.GetApplicationPriority(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)
    
    If lngErrorNo <> 0 Then
    
        ' Write error report
        strErrDesc = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//DESCRIPTION")
        errThrowError strFunctionName, 313, strErrDesc
        errThrowError strFunctionName, 313
    
    End If
   
    strPriority = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//APPLICATIONPRIORITYVALUE")
    
    Set xmlTempRequest = vxmlRequestParentNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "CreateAdHocCaseTask"
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlNode, "APPLICATIONPRIORITY", strPriority
    xmlTempRequest.appendChild xmlNode
    Set xmlNode = vxmlRequestParentNode.ownerDocument.createElement("CASETASK")
    xmlSetAttributeValue xmlNode, "SOURCEAPPLICATION", "Omiga"
    xmlSetAttributeValue xmlNode, "ACTIVITYID", strActivityID
    xmlSetAttributeValue xmlNode, "ACTIVITYINSTANCE", "1"
    xmlSetAttributeValue xmlNode, "STAGEID", strStageId
    xmlSetAttributeValue xmlNode, "CASESTAGESEQUENCENO", strCaseStageSeqNo
    xmlSetAttributeValue xmlNode, "TASKID", strTaskId
    xmlSetAttributeValue xmlNode, "TASKSTATUS", strTaskStatus
    xmlCopyMandatoryAttributeValue xmlValuationNode, xmlNode, "APPLICATIONNUMBER", "CASEID"
    xmlCopyMandatoryAttribute xmlTempRequest, xmlNode, "USERAUTHORITYLEVEL"
    xmlTempRequest.appendChild xmlNode

    Set objomTMBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTMBO")
    strResponse = objomTMBO.OmTmRequest(xmlTempRequest.xml)

    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)

    If lngErrorNo <> 0 Then
    
        ' Write error report
        strErrDesc = xmlGetMandatoryNodeText(xmlTempResponseDoc, ".//DESCRIPTION")
        errThrowError strFunctionName, 313, strErrDesc
    
    End If

    ' Delete the lock
    strResponse = objAppManagerBO.DeleteLock(xmlLockRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    lngErrorNo = errCheckXMLResponseNode(xmlTempResponseDoc.documentElement, vxmlResponseNode)

    If lngErrorNo <> 0 Then
    
        ' Write error report
        errThrowError strFunctionName, 314
 
    End If
   
    gobjContext.SetComplete

ProcessValuationInstructionsExit:
       
    Set objThirdPartyBO = Nothing
    Set objAppManagerBO = Nothing
    Set objIValuationDO = Nothing
    Set objomTMBO = Nothing
        
    Set xmlTempRequest = Nothing
    Set xmlLockRequest = Nothing
    Set xmlRootNode = Nothing
    Set xmlNode = Nothing
    Set xmlValuationNode = Nothing
    Set xmlTempResponseDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
ProcessValuationInstructionsVbErr:

    gobjContext.SetAbort

    GoTo ProcessValuationInstructionsExit
    
End Sub

Private Sub IValuationNTTxBO_ProcessValuerBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ProcessValuerBatchVbErr
    
    Const strFunctionName As String = "IValuationNTTxBO_ProcessValuerBatch"
    
    Dim objIValuationDO As IValuationDO
    Dim objMessageQueue As IomMessageQueue
    Dim xmlTempInstructionRequest As IXMLDOMNode
    Dim xmlTempQueueRequest As IXMLDOMNode
    Dim xmlMessageQueue As IXMLDOMNode
    Dim xmlValuerInstruction As IXMLDOMNode
    Dim xmlValBatchAudit As IXMLDOMNode
    Dim xmlInstruction As IXMLDOMNode
    Dim xmlQueueResponse As IXMLDOMNode
    Dim xmlBatchAuditRequest As IXMLDOMNode
    Dim xmlBatchAudit As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objBatchSchedBO As BatchScheduleBO
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
        
    Dim strQueueName As String
    
    'SG 25/02/02 SYS4097
    'strQueueName = GetMandatoryGlobalParamString("ValuationQueueName")
    
    Set xmlInstruction = xmlGetMandatoryNode(vxmlRequestParentNode, ".//BATCHINSTRUCTION")
    
    'SG 25/02/02 SYS4097
    strQueueName = xmlGetMandatoryAttributeText(xmlInstruction, "QUEUENAME")
    
    ' Create the Batch Audit record
    Set xmlBatchAuditRequest = vxmlRequestParentNode.cloneNode(False)
    xmlSetAttributeValue xmlBatchAuditRequest, "OPERATION", "CreateBatchAudit"
    Set xmlBatchAudit = vxmlRequestParentNode.ownerDocument.createElement("BATCHAUDIT")
    xmlBatchAuditRequest.appendChild xmlBatchAudit
    xmlCopyMandatoryAttribute xmlInstruction, xmlBatchAudit, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlBatchAudit, "BATCHRUNNUMBER"
            
    Set objBatchSchedBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")

    strResponse = objBatchSchedBO.omBatchRequest(xmlBatchAuditRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
     
    ' Create the valuation batch audit record
    Set xmlValBatchAudit = vxmlRequestParentNode.ownerDocument.createElement("VALUATIONBATCHAUDIT")
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHAUDIT")
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "BATCHRUNNUMBER"
    xmlCopyMandatoryAttribute xmlNode, xmlValBatchAudit, "BATCHAUDITGUID"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "APPLICATIONNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "APPLICATIONFACTFINDNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "INSTRUCTIONSEQUENCENO"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValBatchAudit, "DIRECTORYGUID"
    xmlCopyAttribute xmlInstruction, xmlValBatchAudit, "VALUERINVOICEAMOUNT"
    
    Set objIValuationDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
        
    objIValuationDO.CreateValuationAudit xmlValBatchAudit, vxmlResponseNode

    ' Send instruction to queue
    Set xmlTempQueueRequest = vxmlRequestParentNode.cloneNode(False)
    xmlSetAttributeValue xmlTempQueueRequest, "OPERATION", "SendToQueue"
    Set xmlMessageQueue = vxmlRequestParentNode.ownerDocument.createElement("MESSAGEQUEUE")
    xmlTempQueueRequest.appendChild xmlMessageQueue
    xmlSetAttributeValue xmlMessageQueue, "QUEUENAME", strQueueName
    xmlSetAttributeValue xmlMessageQueue, "PROGID", App.Title & ".ValuationBO"
    
    Set xmlTempInstructionRequest = vxmlRequestParentNode.cloneNode(False)
    xmlSetAttributeValue xmlTempInstructionRequest, "OPERATION", "ProcessValuationInstruction"
    Set xmlValuerInstruction = vxmlRequestParentNode.ownerDocument.createElement("VALUATION")
    xmlTempInstructionRequest.appendChild xmlValuerInstruction
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "BATCHRUNNUMBER"
    xmlCopyMandatoryAttribute xmlValBatchAudit, xmlValuerInstruction, "BATCHAUDITGUID"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "APPLICATIONNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "APPLICATIONFACTFINDNUMBER"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "INSTRUCTIONSEQUENCENO"
    xmlCopyMandatoryAttribute xmlInstruction, xmlValuerInstruction, "DIRECTORYGUID"
    xmlCopyAttribute xmlInstruction, xmlValuerInstruction, "VALUERINVOICEAMOUNT"
    
    xmlSetAttributeValue xmlMessageQueue, "XML", xmlTempInstructionRequest.xml
    
    Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")

    Set xmlQueueResponse = objMessageQueue.SendToQueue(xmlTempQueueRequest)
    errCheckXMLResponseNode xmlQueueResponse, vxmlResponseNode, True
        
    gobjContext.SetComplete
    
ProcessValuerBatchExit:

    Set objIValuationDO = Nothing
    Set objMessageQueue = Nothing
    Set xmlTempInstructionRequest = Nothing
    Set xmlTempQueueRequest = Nothing
    Set xmlMessageQueue = Nothing
    Set xmlValuerInstruction = Nothing
    Set xmlValBatchAudit = Nothing
    Set xmlInstruction = Nothing
    Set xmlQueueResponse = Nothing
    Set xmlBatchAuditRequest = Nothing
    Set xmlBatchAudit = Nothing
    Set xmlNode = Nothing
    Set objBatchSchedBO = Nothing
    Set xmlTempResponseDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
ProcessValuerBatchVbErr:

    gobjContext.SetAbort

    GoTo ProcessValuerBatchExit

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

