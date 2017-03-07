VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ValuationBO"
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
' PSC       30/03/01 SYS2188 Add Processing for batch payments
' SG        25/02/02 SYS4097 IValuationBO_ProcessValuerBatch pass global parameter to ValuatioNTTxBo
' PSC       25/02/02 SYS4097 Add additional error processing
' PSC       01/03/02 SYS4214 Add enhance error info in ProcessValuerBatch
' SA        08/03/02 SYS4176 Add call to PrintDocuments in BatchSchedule
' GD        09/04/02 SYS4290 Change OnMessage to SetComplete(), and return success.
' MEVA      18/04/02 SYS3336 Checking Valuation Fee
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IValuationBO
Implements MessageQueueComponentVC2

Private gobjContext As ObjectContext

Private gobjTxBO As IValuationTxBO
Private gobjDO As IValuationDO

' Constants
Private Const cstrValuationComplete = 30
Private gobjAppProcBO As omAppProcBO



'---------------------------------------------------------------------------------------------
'BMIDS Specific History:

'Prog   Date        AQR         Description
'MV     14/08/2002  BMIDS00333  Modified GetApplicationData() - Core Ref AQR: SYS4919
'DPF    03/09/2002  BMIDS00344  Added new methods CreateValuationReportUsingTemplate
'                               and GetValuationReportForTemplate
'MV     03/09/2002  BMIDS00353  Added New Method IValuationBO_GetLatestValuationAmounts()
'DPF    28/10/2002  BMIDS00560  Added check for accepted quote before we reach GetMortgageSubQuoteData call
'MDC    18/11/2002  BMIDS00938 CC014
'BS     12/02/2003  BM0291 Amend ValidateValuationReport method to return old and new LTVs
'GHun   07/04/2003  BM0515 Amend ValidateValuationReport to include differences in LTV smaller than 1%
'BS     19/05/2003  BM0515 Amend ValidateValuationReport to compare numerical LTVs not textual values
'BS     20/05/2003  BM0515 Convert LTVs to doubles instead of long in ValidateValuationReport
'GHun   20/06/2003  BMIDS589 Amend ValidateValuationReport to prevent error 91
'HMA    21/01/2004  BMIDS696 Amend ValidateValuationReport to correct setting of VALUATIONTYPE.
'JD     15/09/2004  BMIDS868 Use correct xml tag in ValidateValuationReport
'JD     23/11/2004  BMIDS604    Set batch status correctly to Cancelled. THIS CHANGE HAS NOT BEEN TESTED
'JD     22/11/2004  BMIDS604    Amended CompleteBatch to look for Cancelled Batch error. THIS CHANGE HAS NOT BEEN TESTED
'
'----------------------------------------------------------------------------------------------
'INGDUK
'JD     28/10/2005  MAR342       Added CreateValuationReportNoInst.
'JD     17/03/2006  MAR1396     CreateValuationReportNoInst - copy some values from the previous valuerinstruction
'----------------------------------------------------------------------------------------------
Private Sub IValuationBO_GetLatestValuationAmounts(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetLatestValuationAmountsExit
    
    Const strFunctionName As String = "IValuationBO_GetLatestValuationAmounts"
    
    Dim xmlGetValuationRepValuationNode  As IXMLDOMNode
    
    Set xmlGetValuationRepValuationNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATIONREPVALUATION")
 
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
  
    gobjDO.GetLatestValuationAmounts xmlGetValuationRepValuationNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    

GetLatestValuationAmountsExit:

    Set xmlGetValuationRepValuationNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub IValuationBO_CreateValuationReportUsingTemplate(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------------------
'This routine takes in a request from omTm.omTMBO, checks if there are any existing
'valuation reports for an application and if so creates a new record, using the old
'one as a template, if it fails to find an existing report it still creates a new
'one but with only the very basic details.
'------------------------------------------------------------------------------------

On Error GoTo CreateValuationReportUsingTemplateExit

Const cstrFunctionName As String = "IValuationBO_CreateValuationReportUsingTemplate"
    
    'Declare Variables
    Dim strGetValuationReport As String
    Dim lngErrorNo As Long
    Dim intSeqNo As Integer
    
    Dim xmlCreateValElem As IXMLDOMElement
    Dim xmlGetValReqElem As IXMLDOMElement
    Dim xmlValuationElem As IXMLDOMElement
    Dim xmlCreateValReqElem As IXMLDOMElement
    Dim xmlGetValRespDOM As FreeThreadedDOMDocument40
    
    Set xmlGetValRespDOM = New FreeThreadedDOMDocument40
    xmlGetValRespDOM.validateOnParse = False
    xmlGetValRespDOM.setProperty "NewParser", True
    
    'ensure we have a link to omAppProcBO object as we need this
    If gobjAppProcBO Is Nothing Then
        Set gobjAppProcBO = gobjContext.CreateInstance("OmAppProc.omAppProcBO")
    End If
            
    'Check for any Existing Valuation Report for this Application
    Set xmlValuationElem = xmlGetMandatoryNode(vxmlRequestParentNode, "./VALUATION").cloneNode(True)
    intSeqNo = CSafeInt(xmlValuationElem.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO").Text)
    
    'check first that this isn't the first valuer isntruction for this application
    If CSafeInt(intSeqNo) > 1 Then
        intSeqNo = CSafeInt(intSeqNo) - 1
        xmlValuationElem.setAttribute "INSTRUCTIONSEQUENCENO", intSeqNo
        Set xmlGetValReqElem = xmlGetRequestNode(vxmlRequestParentNode)
        xmlGetValReqElem.setAttribute "OPERATION", "GetValuationReportForTemplate"
        xmlGetValReqElem.appendChild xmlValuationElem.cloneNode(True)
        'we need to manipulate the response so, call it through the component's interface
        strGetValuationReport = gobjAppProcBO.OmAppProcRequest(xmlGetValReqElem.xml)
        
        'Check Response
        errCheckXMLResponse strGetValuationReport, True
        
        'Load our response from GetValuationReportForTemplate into a DOM Document
        Set xmlGetValRespDOM = xmlLoad(strGetValuationReport, "CreateValuationReportUsingTemplate")
    Else
        'create a dummy <RESPONSE TYPE="SUCCESS"/> response as we don't call anything
        Set xmlValuationElem = xmlGetValRespDOM.createElement("RESPONSE")
        xmlValuationElem.setAttribute "TYPE", "SUCCESS"
        xmlGetValRespDOM.appendChild xmlValuationElem
    End If
    
    'Create new Valuation Report
    If xmlGetValRespDOM.documentElement.hasChildNodes = True Then
    'we have a template, so we can create a new record identical to the last
        Set xmlCreateValElem = xmlGetMandatoryNode(xmlGetValRespDOM.documentElement, "GETVALUATIONREPORT")
        xmlChangeNodeName xmlCreateValElem, "GETVALUATIONREPORT", "VALUATION"
        Set xmlCreateValReqElem = xmlGetRequestNode(vxmlRequestParentNode)
        xmlCreateValReqElem.setAttribute "OPERATION", "CreateValuationReport"
        intSeqNo = CSafeInt(intSeqNo) + 1
        xmlCreateValElem.setAttribute "INSTRUCTIONSEQUENCENO", intSeqNo
        xmlCreateValReqElem.appendChild xmlCreateValElem
        IValuationBO_CreateValuationReport xmlCreateValReqElem, vxmlResponseNode
    Else
    'No existing reports were found so build a basic shell from the request
        Set xmlCreateValReqElem = vxmlRequestParentNode.cloneNode(True)
        xmlCreateValReqElem.setAttribute "OPERATION", "CreateValuationReport"
        IValuationBO_CreateValuationReport xmlCreateValReqElem, vxmlResponseNode
    End If

'exit routine: reset variables, check for errors, exit sub routine
CreateValuationReportUsingTemplateExit:
    
    Set xmlCreateValElem = Nothing
    Set xmlGetValReqElem = Nothing
    Set xmlValuationElem = Nothing
    Set xmlCreateValReqElem = Nothing
    Set xmlGetValRespDOM = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
  
End Sub

Private Sub IValuationBO_GetValuationReportForTemplate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'-----------------------------------------------------------------------------------
'This method checks for an existing valuation report for the application to use as
'a template for a new record, if one is found it is returned
'-----------------------------------------------------------------------------------
    On Error GoTo GetValuationReportForTemplateExit

    Const strFunctionName As String = "IValuationBO_GetValuationReportForTemplate"

    Dim xmlValRepNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlValRepNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/VALUATION")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
    End If

    gobjDO.GetValuationReportForTemplate xmlValRepNode, vxmlResponseNode

GetValuationReportForTemplateExit:

    Set xmlValRepNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_ProcessValuationInstruction(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ProcessValuationInstructionExit

    Const strFunctionName As String = "IValuationBO_ProcessValuationInstruction"
    
    Dim objBatchSchedBO As BatchScheduleBO
    Dim objValuationNTTxBO As IValuationNTTxBO
    Dim objMessageQueue As IomMessageQueue
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempBatchRequest As IXMLDOMNode
    Dim xmlBatchAudit As IXMLDOMNode
    Dim xmlBatchAuditRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlBatchSchedule As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlValuationNode As IXMLDOMNode
    Dim xmlQueueResponse As IXMLDOMElement
    Dim xmlMessageQueue As IXMLDOMElement
    Dim xmlTempQueueRequest As IXMLDOMNode
    
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strCompleteWithErrors As String
    Dim strComplete As String
    Dim blnLastTransaction As Boolean
    
    'Validation
    Set xmlValuationNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATION")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlValuationNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlValuationNode, "BATCHRUNNUMBER")
    
    'Create request to get Batch data
    Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "GetBatchSchedule"
    Set xmlBatchSchedule = vxmlRequestNode.ownerDocument.createElement("BATCHSCHEDULE")
    xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchSchedule, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchSchedule, "BATCHRUNNUMBER"
    xmlTempBatchRequest.appendChild xmlBatchSchedule

    'Create and call the Batch Scheduler component to get current batch info
    Set objBatchSchedBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
    
    Set xmlTempRequest = Nothing
   
    ' Check if batch has been cancelled
    Dim strBatchStatus As String
    
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHSCHEDULE")
    strBatchStatus = xmlGetMandatoryAttributeText(xmlNode, "STATUS")
    
    Dim lngErrorNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    If IsValidationType("BatchScheduleStatus", CSafeInt(strBatchStatus), "CA") = False Then
        
        Set objValuationNTTxBO = gobjContext.CreateInstance(App.Title & ".ValuationNTTxBO")
        
        On Error Resume Next
        
        objValuationNTTxBO.ProcessValuationInstruction vxmlRequestNode, vxmlResponseNode
        
        lngErrorNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
        
        On Error GoTo ProcessValuationInstructionExit
    
        ' Instruction processed OK
        If lngErrorNo = 0 Then
            
            ' Update the batch success count
            xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "IncrementBatchScheduleSuccess"
            xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchSchedule, "BATCHAUDITGUID"
            strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
            Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
            
            Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHSCHEDULE")
            blnLastTransaction = xmlGetMandatoryAttributeAsBoolean(xmlNode, "LASTTRANSACTION")
            
        Else ' There's been an error
        
            Dim lngNumRetries As Long
            Dim lngCurrentRetries As Long
            Dim blnResendRequest As Boolean
            Dim strQueueName As String
            
            blnResendRequest = False
            
            ' Application is locked so check to see if the instruction is to be retried
            If errGetOmigaErrorNumber(lngErrorNo) = 310 Then
                        
                lngNumRetries = GetMandatoryGlobalParamAmount("ValuationRetries")
            
                ' Get the batch audit data in order to check number of retries
                Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
            
                xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "FindBatchAuditList"
                Set xmlBatchAudit = vxmlRequestNode.ownerDocument.createElement("BATCHAUDIT")
                xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchAudit, "BATCHNUMBER"
                xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchAudit, "BATCHRUNNUMBER"
                xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchAudit, "BATCHAUDITGUID"
                xmlTempBatchRequest.appendChild xmlBatchAudit
    
                strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
                Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
                
                Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHAUDIT")
                lngCurrentRetries = xmlGetAttributeAsLong(xmlNode, "NUMBEROFRETRIES")
                
                If lngCurrentRetries < lngNumRetries Then
                    blnResendRequest = True
                End If
            
            End If
            
            ' If the instruction is to be retried then send it back to the queue else it's a
            ' failure
            If blnResendRequest = True Then
            
                ' Increment the number of retries for this instruction
                xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "IncrementBatchAuditRetries"
                strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
                Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
    
                ' Set up queue request
                strQueueName = GetMandatoryGlobalParamString("ValuationQueueName")
    
                Set xmlTempQueueRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlTempQueueRequest, "OPERATION", "SendToQueue"
                Set xmlMessageQueue = xmlTempResponseDoc.createElement("MESSAGEQUEUE")
                xmlTempQueueRequest.appendChild xmlMessageQueue
                xmlMessageQueue.setAttribute "QUEUENAME", strQueueName
                xmlMessageQueue.setAttribute "PROGID", App.Title & ".ValuationBO"
                xmlMessageQueue.setAttribute "XML", vxmlRequestNode.xml
    
                Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
            
                Set xmlQueueResponse = objMessageQueue.SendToQueue(xmlTempQueueRequest)
                errCheckXMLResponseNode xmlQueueResponse, vxmlResponseNode, True
    
            Else
        
                ' Update the batch failure count
                Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "IncrementBatchScheduleFailure"
                xmlCopyMandatoryAttribute xmlValuationNode, xmlBatchSchedule, "BATCHAUDITGUID"
                xmlSetAttributeValue xmlBatchSchedule, "ERRORNUMBER", CStr(lngErrorNo)
                xmlSetAttributeValue xmlBatchSchedule, "ERRORSOURCE", strSource
                xmlSetAttributeValue xmlBatchSchedule, "ERRORDESCRIPTION", strDescription
                xmlTempBatchRequest.appendChild xmlBatchSchedule
                
                strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
                Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
            
                Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHSCHEDULE")
                blnLastTransaction = xmlGetMandatoryAttributeAsBoolean(xmlNode, "LASTTRANSACTION")
            
            End If
        End If
        
        ' Do final processing if last transaction in batch
        If blnLastTransaction = True Then
            CompleteBatch vxmlRequestNode, vxmlResponseNode
        End If
    Else
        'JD BMIDS604 Batch is cancelled, raise an error to prevent the batch status being set as Complete
        errThrowError strFunctionName, 7025
    End If
     
ProcessValuationInstructionExit:

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
End Sub

Private Function MessageQueueComponentVC2_OnMessage(ByVal in_xmlConfig As String, ByVal in_xmlData As String) As Long

    Const strFunctionName As String = "MessageQueueComponentVC2_OnMessage"
    Dim strErrMsg As String
    Dim objomAppProcBO As omAppProcBO
    Dim strResponse As String
    Dim lngErrorNo As Long

    Set objomAppProcBO = gobjContext.CreateInstance(App.Title & ".omAppProcBO")
    strResponse = objomAppProcBO.OmAppProcRequest(in_xmlData)
    lngErrorNo = errCheckXMLResponse(strResponse)
    
OnMessageExit:
'GD 09/04/02 SYS4290 - Start
    'If lngErrorNo = 0 Then
        'MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
    'Else
        'MessageQueueComponentVC2_OnMessage = MESSQ_RESP_RETRY_MOVE_MESSAGE
    'End If
    If lngErrorNo <> 0 Then
        strErrMsg = "Error Processing Message. " & strFunctionName & "." & vbCrLf & _
            "Number: " & CStr(Err.Number) & vbCrLf & _
            "Source: " & Err.Source & vbCrLf & _
            "Description: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
        gobjContext.SetComplete
    End If

    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
'GD 09/04/02 SYS4290 - End

End Function

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Private Sub IValuationBO_CreateValuationReportNoInst(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'This routine has been written to create a valuation report which has not been
'created as a result of a valuer instruction. This is a copy of the omTMBO method
'with the UpdateCaseTask removed. It is called from the ESurv Interface.

    On Error GoTo CreateValuationReportNoInstErr
    
    'Declare Variables
    Dim xmlExtractDOM As FreeThreadedDOMDocument40
    Dim xmlValuerResponse As FreeThreadedDOMDocument40
    Dim xmlValuerRequest As FreeThreadedDOMDocument40
    Dim xmlComboRequest As FreeThreadedDOMDocument40
    Dim xmlComboResponse As FreeThreadedDOMDocument40
    Set xmlExtractDOM = New FreeThreadedDOMDocument40
    xmlExtractDOM.validateOnParse = False
    xmlExtractDOM.setProperty "NewParser", True
    Set xmlValuerRequest = New FreeThreadedDOMDocument40
    xmlValuerRequest.validateOnParse = False
    xmlValuerRequest.setProperty "NewParser", True
    Set xmlValuerResponse = New FreeThreadedDOMDocument40
    xmlValuerResponse.validateOnParse = False
    xmlValuerResponse.setProperty "NewParser", True
    Set xmlComboRequest = New FreeThreadedDOMDocument40
    xmlComboRequest.validateOnParse = False
    xmlComboRequest.setProperty "NewParser", True
    Set xmlComboResponse = New FreeThreadedDOMDocument40
    xmlComboResponse.validateOnParse = False
    xmlComboResponse.setProperty "NewParser", True
    
    Dim xmlValElem As IXMLDOMElement
    Dim xmlInstElem As IXMLDOMElement
    Dim xmlValRequest As IXMLDOMElement
    Dim xmlInstRequest As IXMLDOMElement
    Dim xmlCaseReq As IXMLDOMElement
    Dim xmlValuerReqElem As IXMLDOMElement
    Dim xmlValuerList As IXMLDOMElement
    Dim xmlValuerPanelID As IXMLDOMElement
    Dim xmlValuerValType As IXMLDOMElement
    Dim xmlList As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlOldValNode As IXMLDOMNode 'MAR1396
        
    Dim objOmTmBo As Object
    Dim objAppProcBO As Object
    Dim objThirdPartyBO As Object
    Dim objComboDO As Object
    
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlValuationType As IXMLDOMNode
         
    Dim strCaseResp As String
    Dim strReportResponse As String
    Dim strInstResponse As String
    Dim intSeqNo As Integer
    Dim strValuerResponse As String
    Dim strDirectoryGUID As String
    Dim strComboList As String
    Dim strStandardValuer As String
    
    Const cstrFunctionName As String = "CreateValuationReportNoInst"
    
    'amend request node
    xmlChangeNodeName vxmlRequestNode, "APPLICATION", "VALUATION"
        
    'set up connection to Application Processing & Third Party Business Objects
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
    
    'DPF 06/11/2002 - BMIDS00852
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        
    If objThirdPartyBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO"
    End If
    
    If objAppProcBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    'DPF 06/11/2002 - BMIDS00852
    If objComboDO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrBASE_COMPONENT & ".ComboDO"
    End If
      
    'DPF 5/11/2002 - BMIDS00852 - use Combo Assist to pull back valuer id for Validation type 'L'
    strStandardValuer = "L"
    Set xmlList = xmlComboRequest.createElement("LIST")
    xmlComboRequest.appendChild xmlList
    Set xmlNode = xmlComboRequest.createElement("LISTNAME")
    xmlNode.Text = "ValuationType"
    xmlList.appendChild xmlNode
    
    strComboList = objComboDO.GetComboList(xmlList.xml)
    Set xmlComboResponse = xmlLoad(strComboList, "CreateValuationReportNoInst")
        
    Set xmlValuationType = xmlComboResponse.selectSingleNode("//LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE = '" & strStandardValuer & "']")
    'DPF - END OF BMIDS00852
      
    'Retrieve the Directory GUID of a default Valuer set up in supervisor
    Set xmlValuerReqElem = xmlGetRequestNode(vxmlRequestNode)
    xmlValuerReqElem.setAttribute "OPERATION", "FindPanelValuerList"
    Set xmlValuerList = xmlValuerRequest.createElement("VA_PANELVALUERLIST")
    Set xmlValuerValType = xmlValuerRequest.createElement("VALUATIONTYPE")
    Set xmlValuerPanelID = xmlValuerRequest.createElement("PANELID")
    xmlValuerPanelID.Text = "Default"
    xmlValuerValType.Text = xmlValuationType.selectSingleNode("VALUEID").Text 'DPF 05/11/2002 - BMIDS00852
    xmlValuerList.appendChild xmlValuerValType
    xmlValuerList.appendChild xmlValuerPanelID
    xmlValuerReqElem.appendChild xmlValuerList
    xmlValuerRequest.appendChild xmlValuerReqElem
        
    strValuerResponse = objThirdPartyBO.FindPanelValuerList(xmlValuerRequest.xml)
    
    'Check Response
    errCheckXMLResponse strValuerResponse, True
        
    'Assign directory guid retrieved into Variable
    xmlValuerResponse.loadXML (strValuerResponse)
    'AW     20/12/2002  BM0190
    'strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "VA_PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
    strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
        
    'MAR1396 - find values from previous valuerinstruction if present to pass on to new one
    Set xmlInstElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION").cloneNode(True)
    xmlChangeNodeName xmlInstElem, "VALUATION", "VALUERINSTRUCTION"
    xmlInstElem.setAttribute "_ORDERBY_", "INSTRUCTIONSEQUENCENO DESC"
    Set xmlInstRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlInstRequest.setAttribute "OPERATION", "FindValuerInstructionList"
    xmlInstRequest.appendChild xmlInstElem.cloneNode(True)
    
    strInstResponse = objAppProcBO.OmAppProcRequest(xmlInstRequest.xml)
    Set xmlExtractDOM = xmlLoad(strInstResponse, "FindValuerInstructionList")
    'The first in the list will be the latest as we did an ORDERBY
    Set xmlOldValNode = xmlExtractDOM.documentElement.selectNodes("VALUERINSTRUCTION").Item(0)
    
    'Create 'Dummy' Valuer Instruction record
    Set xmlInstElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION").cloneNode(True)
    xmlChangeNodeName xmlInstElem, "VALUATION", "VALUERINSTRUCTION"
    xmlInstElem.setAttribute "VALUERPANELNO", "Default"
    xmlInstElem.setAttribute "DIRECTORYGUID", strDirectoryGUID
    'MAR1396 if we have a previous valuerinstruction then copy some values across
    If Not xmlOldValNode Is Nothing Then
        xmlInstElem.setAttribute "VALUATIONTYPE", xmlGetAttributeText(xmlOldValNode, "VALUATIONTYPE")
        xmlInstElem.setAttribute "APPOINTMENTDATE", xmlGetAttributeText(xmlOldValNode, "APPOINTMENTDATE")
        xmlInstElem.setAttribute "DATEOFINSTRUCTION", xmlGetAttributeText(xmlOldValNode, "DATEOFINSTRUCTION")
    End If
    Set xmlInstRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlInstRequest.setAttribute "OPERATION", "CreateValuerInstructions"
    xmlInstRequest.appendChild xmlInstElem.cloneNode(True)
    
    strInstResponse = objAppProcBO.OmAppProcRequest(xmlInstRequest.xml)
    
    'Check Response
    errCheckXMLResponse strInstResponse, True
    
    'Extract Instruction Sequence No of record created
    Set xmlExtractDOM = xmlLoad(strInstResponse, "CreateValuationReportNoInst")
    intSeqNo = xmlExtractDOM.documentElement.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO").Text
              
    '-----------------------------------------------------------------------------
        
    'Call omAppproc.ValuationBO.CreateValuationReportUsingTemplate
    'passing in initial request data & Instruction Seq No of record just created
    Set xmlValElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION")
    xmlValElem.setAttribute "INSTRUCTIONSEQUENCENO", intSeqNo
    Set xmlValRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlValRequest.setAttribute "OPERATION", "CreateValuationReportUsingTemplate"
    xmlValRequest.appendChild xmlValElem.cloneNode(True)
            
    strReportResponse = objAppProcBO.OmAppProcRequest(xmlValRequest.xml)
    
    'add the seq num to the response node
    xmlSetAttributeValue vxmlResponseNode, "INSTRUCTIONSEQUENCENO", intSeqNo
    
    
    Set objAppProcBO = Nothing
        
    'Check Response
    errCheckXMLResponse strReportResponse, True
          
           
CreateValuationReportNoInstExit:
    
    Set xmlExtractDOM = Nothing
    Set xmlValuerRequest = Nothing
    Set xmlValuerResponse = Nothing
    Set xmlComboRequest = Nothing
    Set xmlComboResponse = Nothing
    
    Set xmlValElem = Nothing
    Set xmlInstElem = Nothing
    Set xmlValRequest = Nothing
    Set xmlInstRequest = Nothing
    Set xmlCaseReq = Nothing
    Set xmlValuerReqElem = Nothing
    Set xmlValuerList = Nothing
    Set xmlValuerPanelID = Nothing
    Set xmlValuerValType = Nothing
    Set xmlList = Nothing
    Set xmlNode = Nothing
        
    Set objOmTmBo = Nothing
    Set objThirdPartyBO = Nothing
    Set objComboDO = Nothing
    
    Set xmlTableNode = Nothing
    Set xmlValuationType = Nothing
            
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateValuationReportNoInstErr:
 
    'MsgBox "Source:  " & Err.Source & " Description:  " & Err.Description & " Number:  " & Err.Number
    GoTo CreateValuationReportNoInstExit
End Sub
Private Sub IValuationBO_CreateValuationReport(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                               ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve Valuation Report for the application number and instruction sequence number passed
'   in
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

    On Error GoTo CreateValuationReportExit
    
    Const strFunctionName As String = "IValuationBO_CreateValuationReport"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestParentNode, "VALUATION"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ValuationTXBO")
    End If
    
    gobjTxBO.CreateValuationReport vxmlRequestParentNode, vxmlResponseNode
    
CreateValuationReportExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_GetValuationPropertyDetails(ByVal vxmlRequestParentNode As IXMLDOMNode, _
                                                     ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetValuationPropertyDetailsExit
    
    Const strFunctionName As String = "IValuationBO_GetValuationPropertyDetails"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisGetValuationPropertyDetailsNode As IXMLDOMNode
    Dim xmlGetValuationPropertyDetailsNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    
    Dim strAttribValue As String
      
    Set xmlGetValuationPropertyDetailsNode = xmlGetMandatoryNode(vxmlRequestParentNode, "VALUATIONPROPERTYDETS")
 
    'Set xmlThisRequestDoc = New DOMDocument

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
  
    gobjDO.GetValuationPropertyDetails xmlGetValuationPropertyDetailsNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    

GetValuationPropertyDetailsExit:

    Set xmlGetValuationPropertyDetailsNode = Nothing
    Set xmlThisGetValuationPropertyDetailsNode = Nothing
    Set xmlThisRequestDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IValuationBO_GetValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LENDERSREFERENCE and CurrentLendersRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CURRENTLENDERSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetValuationReportExit

    Const strFunctionName As String = "IValuationBO_GetValuationReport"

    Dim xmlValRepNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlValRepNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/VALUATION")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
    End If

    gobjDO.GetValuationReport xmlValRepNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetValuationReportExit:

    Set xmlValRepNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_GetValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

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

    On Error GoTo GetValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationBO_GetValuerInstructions"
    
    Dim xmlGetValuerInstructionsNode As IXMLDOMNode
    
    ' validate REQUEST & select ValuerInstructions node
       
    Set xmlGetValuerInstructionsNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUERINSTRUCTIONS")
  
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    gobjDO.GetValuerInstructions xmlGetValuerInstructionsNode, vxmlResponseNode
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetValuerInstructionsExit:

    Set xmlGetValuerInstructionsNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_UpdateValuationReport(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified EMPLOYERSREFERENCE and PREVIOUSEMPLOYERSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateValuationReportExit
    
    Const strFunctionName As String = "IValuationBO_UpdateValuationReport"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ValuationTxBO")
    End If
    
    gobjTxBO.UpdateValuationReport vxmlRequestParentNode, vxmlResponseNode
    
UpdateValuationReportExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_ValidateInvoiceAmount(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve Valuation Report for the application number and instruction sequence number passed
'   in
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

    On Error GoTo ValidateInvoiceAmountExit
    
    Const strFunctionName As String = "IValuationBO_ValidateInvoiceAmount"
    
    Dim xmlValuationNode As IXMLDOMNode
    Dim strAuthorisationRequired As String
    Dim dblInvoiceAmount As Double
    Dim dblValuationFee As Double
    Dim dblValuationFeeTolerance As Double
    Dim strTypeOfMortgage As String
    Dim strAppNo As String
    Dim strAppFactFindNo As String
    Dim strInsSeqNo As String
    Dim objApplicationBO As ApplicationBO
    Dim strAppResp As String
    Dim domAppReq As FreeThreadedDOMDocument40
    Dim domAppResp As New FreeThreadedDOMDocument40
    domAppResp.validateOnParse = False
    domAppResp.setProperty "NewParser", True
    Dim xmlAppReq As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlTempRespNode As IXMLDOMNode
    
    ' validate REQUEST & select VALUATION node
    Set xmlValuationNode = xmlGetMandatoryNode(vxmlRequestParentNode, "VALUATION")
    
    '
    ' Get the Invoice Amount Provided
    '
    dblInvoiceAmount = xmlGetMandatoryAttributeAsDouble(xmlValuationNode, "INVOICEAMOUNT")

    strAppNo = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONNUMBER")
    strAppFactFindNo = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONFACTFINDNUMBER")
    strInsSeqNo = xmlGetAttributeText(xmlValuationNode, "INSTRUCTIONSEQUENCENO")

    '
    ' Get Application Fact Find data
    '
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Set xmlAppReq = xmlGetRequestNode(vxmlRequestParentNode)
    Set domAppReq = xmlAppReq.ownerDocument
    Set xmlAppNode = xmlAppReq.appendChild(xmlAppReq.ownerDocument.createElement("APPLICATIONFACTFIND"))

    xmlAppNode.appendChild(domAppReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlAppNode.appendChild(domAppReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo

    ' Do the call
    strAppResp = objApplicationBO.GetApplicationFactFindData(xmlAppReq.xml)

    ' Load the response and check it for errors
    Set domAppResp = xmlLoad(strAppResp, strFunctionName)

    Set xmlAppNode = domAppResp.firstChild
    errCheckXMLResponseNode xmlAppNode, , True
    Set xmlAppNode = domAppResp.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
    
    If Not xmlAppNode Is Nothing Then
        strTypeOfMortgage = xmlAppNode.Text
    End If
        
    '
    ' Get ValuerInstruction data GetValuerInstructions
    '
    Set domAppReq = New FreeThreadedDOMDocument40
    domAppReq.validateOnParse = False
    domAppReq.setProperty "NewParser", True
    Set xmlElem = domAppReq.appendChild(domAppReq.createElement("VALUERINSTRUCTIONS"))
    
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFindNo
    xmlElem.setAttribute "INSTRUCTIONSEQUENCENO", strInsSeqNo
    
    ' Do the call
    Set domAppResp = New FreeThreadedDOMDocument40
    domAppResp.validateOnParse = False
    domAppResp.setProperty "NewParser", True
    Set xmlTempRespNode = domAppResp.createElement("RESPONSE")
    
    ' First, get the Valuation Report
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
    End If
    
    ' Do the call
    gobjDO.GetValuerInstructions xmlElem, xmlTempRespNode
    
    'IValuationBO_GetValuerInstructions vxmlRequestParentNode, xmlTempRespNode
    
    Set xmlAppNode = xmlTempRespNode.selectSingleNode("RESPONSE/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
    
    If Not xmlAppNode Is Nothing Then
        dblValuationFee = xmlGetMandatoryAttributeText(xmlTempRespNode, "VALUATIONFEE")
    Else
        dblValuationFee = 0
    End If
    
    '
    ' Get Valuation Fee Tolerance from Global Parameters
    '
    dblValuationFeeTolerance = GetGlobalParamAmount("ValuationFeeTolerance")
    
    '
    ' Check Tolerence level not exceeded
    '
    If (strTypeOfMortgage = "30") Then
        '
        ' Further Advance
        '
        strAuthorisationRequired = 0
    Else
        '
        '
        '
        If (dblInvoiceAmount - dblValuationFee > dblValuationFeeTolerance) Then
            strAuthorisationRequired = "1"
        Else
            strAuthorisationRequired = "0"
        End If
        
    End If
    
    '
    '
    '
    Set domAppResp = New FreeThreadedDOMDocument40
    domAppResp.validateOnParse = False
    domAppResp.setProperty "NewParser", True
    Set xmlElem = domAppResp.appendChild(domAppResp.createElement("VALUATION"))
    xmlElem.setAttribute "AUTHORISATIONREQUIRED", strAuthorisationRequired
    vxmlResponseNode.appendChild xmlElem
    
ValidateInvoiceAmountExit:

    Set xmlAppReq = Nothing
    Set xmlAppNode = Nothing
    Set xmlElem = Nothing
    Set xmlTempRespNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_ValidateValuationReport(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve Valuation Report for the application number and instruction sequence number passed
'   in
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

    On Error GoTo ValidateValuationReportExit
    
    Const strFunctionName As String = "IValuationBO_ValidateValuationReport"
    Dim strAppNo As String
    Dim strInsSeqNo As String
    Dim strAppFactFindNo As String
    Dim strTypeOfMortgage As String
    Dim strCustomerNumber(5) As String
    Dim strCustomerVersionNumber(5) As String
    Dim intCustomerIndex As Integer
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlValuationNode As IXMLDOMNode
    Dim domReq As FreeThreadedDOMDocument40
    Dim domResp As FreeThreadedDOMDocument40
    Dim xmlTempRespNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim objApplicationBO As ApplicationBO
    ' Get Valuation node
    Set xmlValuationNode = xmlGetMandatoryNode(vxmlRequestParentNode, "/REQUEST/VALUATION")
    
    ' Get input data
    
    strAppNo = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONNUMBER")
    strAppFactFindNo = xmlGetMandatoryAttributeText(xmlValuationNode, "APPLICATIONFACTFINDNUMBER")
    strInsSeqNo = xmlGetAttributeText(xmlValuationNode, "INSTRUCTIONSEQUENCENO")
    strTypeOfMortgage = xmlGetAttributeText(xmlValuationNode, "TYPEOFMORTGAGE")
    Set xmlCustomerList = xmlValuationNode.selectNodes("CUSTOMER")
    intCustomerIndex = 0
    For Each xmlCustomerNode In xmlCustomerList
        intCustomerIndex = intCustomerIndex + 1
        strCustomerNumber(intCustomerIndex) = xmlGetAttributeText(xmlCustomerNode, "CUSTOMERNUMBER")
        strCustomerVersionNumber(intCustomerIndex) = xmlGetAttributeText(xmlCustomerNode, "CUSTOMERVERSIONNUMBER")
    Next
  
    
    ' First, get the Valuation Report
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ValuationDO")
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''
    ' Get the Valuation report
    Dim xmlValReportNode As IXMLDOMNode
    
    Set domReq = New FreeThreadedDOMDocument40
    domReq.validateOnParse = False
    domReq.setProperty "NewParser", True
    Set xmlElem = domReq.appendChild(domReq.createElement("GETVALUATIONDETAILS"))
    
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFindNo
    If Len(strInsSeqNo) > 0 Then
        xmlElem.setAttribute "INSTRUCTIONSEQUENCENO", strInsSeqNo
    End If
    ' Do the call
    Set domResp = New FreeThreadedDOMDocument40
    domResp.validateOnParse = False
    domResp.setProperty "NewParser", True
    Set xmlTempRespNode = domResp.createElement("RESPONSE")
    Set xmlTempRespNode = domResp.appendChild(xmlTempRespNode)
    gobjDO.GetValuationReport xmlElem, xmlTempRespNode
    
        
    Set xmlValReportNode = xmlTempRespNode.selectSingleNode("GETVALUATIONREPORT")
    ' JR - SYS2256, check if we have a Seq No. as its required for other operations
    If strInsSeqNo = "" Then
        If Not xmlValReportNode Is Nothing Then
            strInsSeqNo = xmlGetAttributeText(xmlValReportNode, "INSTRUCTIONSEQUENCENO")
        End If
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''
    ' Get the Valuation Instructions
    ''''''''''''''''''''''''''''''''''''''''''
    Set domReq = New FreeThreadedDOMDocument40
    domReq.validateOnParse = False
    domReq.setProperty "NewParser", True
    Set xmlElem = domReq.appendChild(domReq.createElement("VALUERINSTRUCTIONS"))
    
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFindNo
    xmlElem.setAttribute "INSTRUCTIONSEQUENCENO", strInsSeqNo
    
    ' Do the call
    Set domResp = New FreeThreadedDOMDocument40
    domResp.validateOnParse = False
    domResp.setProperty "NewParser", True
    Set xmlTempRespNode = domResp.createElement("RESPONSE")
    gobjDO.GetValuerInstructions xmlElem, xmlTempRespNode
    
    ''''''''''''''''''''''''''''''''''''''''''
    ' Get the application fact find details
    ''''''''''''''''''''''''''''''''''''''''''
    Dim strAppResp As String
    Dim domAppReq As FreeThreadedDOMDocument40
    Dim domAppResp As FreeThreadedDOMDocument40
    Dim xmlAppReq As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Set xmlAppReq = xmlGetRequestNode(vxmlRequestParentNode)
    Set domAppReq = xmlAppReq.ownerDocument
    Set xmlAppNode = xmlAppReq.appendChild(xmlAppReq.ownerDocument.createElement("APPLICATIONFACTFIND"))

    xmlAppNode.appendChild(domAppReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlAppNode.appendChild(domAppReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo

    ' Do the call
    strAppResp = objApplicationBO.GetApplicationFactFindData(xmlAppReq.xml)

    ' Load the response and check it for errors
    Set domAppResp = New FreeThreadedDOMDocument40
    domAppResp.validateOnParse = False
    domAppResp.setProperty "NewParser", True
    Set domAppResp = xmlLoad(strAppResp, strFunctionName)

    Set xmlAppNode = domAppResp.firstChild
    errCheckXMLResponseNode xmlAppNode, , True
    Set xmlAppNode = domAppResp.selectSingleNode("RESPONSE/APPLICATIONFACTFIND")
    ''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''
    ' JR - SYS2256, plugged-in the Valuation rules
    RunValuationRules vxmlRequestParentNode, vxmlResponseNode
   
    ''''''''''''''''''''''''''''''''''''
    ' Update the Valuation Instructions
    ''''''''''''''''''''''''''''''''''''
    Dim xmlValElem As IXMLDOMElement
    Dim xmlValInstructionReq As IXMLDOMElement
    Dim xmlDOMVal As FreeThreadedDOMDocument40
    Set xmlValInstructionReq = xmlGetRequestNode(vxmlRequestParentNode)
    Set xmlDOMVal = xmlValInstructionReq.ownerDocument

    ' JR - SYS2256, changed element name to VALUERINSTRUCTION
    
    ' DRC - SYS2093, This section commented out because it is done in the GUI(AP205) on
    ' a successful return
'    Set xmlValElem = xmlDOMVal.createElement("VALUERINSTRUCTION")
'    xmlValElem.setAttribute "APPLICATIONNUMBER", strAppNo
'    xmlValElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFactFindNo
'    xmlValElem.setAttribute "INSTRUCTIONSEQUENCENO", strInsSeqNo
'    xmlValElem.setAttribute "VALUATIONSTATUS", cstrValuationComplete
'    xmlValInstructionReq.appendChild xmlValElem
'
'    IValuationBO_UpdateValuerInstructions xmlValInstructionReq, vxmlResponseNode
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' DC - SYS2256
    ' Update the latest valuation

    '''''''''''''''''''''''''''''''''''''''''''''
    Dim dblLatestValuation As Double
    Dim dblPresentValuation As Double
    Dim dblPostWorksValuation As Double
    
    dblPresentValuation = xmlGetAttributeAsDouble(xmlValReportNode, "PRESENTVALUATION")
    dblPostWorksValuation = xmlGetAttributeAsDouble(xmlValReportNode, "POSTWORKSVALUATION")
    
    If dblPostWorksValuation > 0# Then
        dblLatestValuation = dblPostWorksValuation
    ElseIf dblPresentValuation > 0# Then
        dblLatestValuation = dblPresentValuation
    Else
        dblLatestValuation = 0#
    End If
    
    Dim strNewPropertyResp As String
    Dim domNewPropertyReq As FreeThreadedDOMDocument40
    Dim domNewPropertyResp As FreeThreadedDOMDocument40
    Dim xmlNewPropertyReq As IXMLDOMNode
    Dim xmlNewPropertyNode As IXMLDOMNode
    Dim objNewPropertyBO As NewPropertyBO
    Dim xmlValInstNode As IXMLDOMNode           ' BMIDS696
    
    Set objNewPropertyBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyBO")
    Set xmlNewPropertyReq = xmlGetRequestNode(vxmlRequestParentNode)
    Set domNewPropertyReq = xmlNewPropertyReq.ownerDocument
    Set xmlNewPropertyNode = xmlNewPropertyReq.appendChild(xmlNewPropertyReq.ownerDocument.createElement("NEWPROPERTY"))
    xmlNewPropertyNode.appendChild(domNewPropertyReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlNewPropertyNode.appendChild(domNewPropertyReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo
        
    'BMIDS696  Select the VALUERINSTRUCTIONS node
    Set xmlValInstNode = xmlTempRespNode.selectSingleNode("VALUERINSTRUCTIONS")
        
    'BMIDS00938 MDC 18/11/2002 - CC014
    'If strTypeOfMortgage = "30" Then
    If IsValidationType("TypeOfMortgage", strTypeOfMortgage, "F") Then
    'BMIDS00938 MDC 18/11/2002 - End
    
    '   ''''''''''''''''''''''''''''''''''''''''''''''
    ' Update the property record with the latest valuation
    '   ''''''''''''''''''''''''''''''''''''''''''''''
    
        Dim strValuerType As String
  
        xmlNewPropertyNode.appendChild(domNewPropertyReq.createElement("VALUATIONPRICE")).Text = CStr(dblLatestValuation)
        
        'BMIDS696  Select the VALUATIONTYPE from the VALUERINSTRUCTIONS node
        strValuerType = ""
        If Not xmlValInstNode Is Nothing Then
            ' JD BMIDS868 use correct tag    strValuerType = xmlGetAttributeText(xmlValInstNode, "VALUATIONTYPE")
            strValuerType = xmlGetAttributeText(xmlValInstNode, "NEWPROPERTYVALUATIONTYPE")
        End If
        
        'strValuerType = xmlGetAttributeText(xmlTempRespNode, "VALUATIONTYPE")
        
        xmlNewPropertyNode.appendChild(domNewPropertyReq.createElement("VALUATIONTYPE")).Text = strValuerType
        ' Do the call
        strNewPropertyResp = objNewPropertyBO.UpdateNewProperty(xmlNewPropertyReq.xml)
    
        ' Load the response and check it for errors
        Set domNewPropertyResp = xmlLoad(strNewPropertyResp, strFunctionName)
        Set xmlNewPropertyNode = domNewPropertyResp.firstChild
        errCheckXMLResponseNode xmlNewPropertyNode, , True
        
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the Discount Amount from the property details
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim dblDiscountAmount As Double
    ' Do the call
    strNewPropertyResp = objNewPropertyBO.GetFullNewPropertyDetails(xmlNewPropertyReq.xml)

    ' Load the response and check it for errors
    Set domNewPropertyResp = xmlLoad(strNewPropertyResp, strFunctionName)
    Set xmlNewPropertyNode = domNewPropertyResp.firstChild
    errCheckXMLResponseNode xmlNewPropertyNode, , True
    Set xmlNewPropertyNode = domNewPropertyResp.selectSingleNode("RESPONSE/NEWPROPERTY/DESCRIPTIONOFPROPERTY")
    dblDiscountAmount = xmlGetNodeAsDouble(xmlNewPropertyNode, "DISCOUNTAMOUNT")
    ''''''''''''''''''''''''''''''''''''''''''
    '
    ' Check to see if Loan to Value percentage has altered
    ' Get data for LTV calcs - amount requested and present LTV
    '
    ''''''''''''''''''''''''''''''''''''''''''
    Dim strLTVChanged As String
    strLTVChanged = "0"
    ' Get the sub quote details
    ''''''''''''''''''''''''''''''''''''''''''
    Dim strStoredQuoteResp As String
    Dim domStoredQuoteReq As FreeThreadedDOMDocument40
    Dim domStoredQuoteResp As FreeThreadedDOMDocument40
    Dim xmlStoredQuoteReq As IXMLDOMNode
    Dim xmlStoredQuoteNode As IXMLDOMNode
    Dim xmlStoredQuoteNodeList As IXMLDOMNode
    Dim objQuotationBO As QuotationBO
    
    Set objQuotationBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    Set xmlStoredQuoteReq = xmlGetRequestNode(vxmlRequestParentNode)
    Set domStoredQuoteReq = xmlStoredQuoteReq.ownerDocument
    Set xmlStoredQuoteNode = xmlStoredQuoteReq.appendChild(xmlStoredQuoteReq.ownerDocument.createElement("SEARCH"))
    Set xmlStoredQuoteNode = xmlStoredQuoteNode.appendChild(xmlStoredQuoteReq.ownerDocument.createElement("QUOTATION"))

    xmlStoredQuoteNode.appendChild(domStoredQuoteReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlStoredQuoteNode.appendChild(domStoredQuoteReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo

    ' Do the call
    strStoredQuoteResp = objQuotationBO.FindStoredQuoteDetails(xmlStoredQuoteReq.xml)

    ' Load the response and check it for errors
    Set domStoredQuoteResp = New FreeThreadedDOMDocument40
    domStoredQuoteResp.validateOnParse = False
    domStoredQuoteResp.setProperty "NewParser", True
    Set domStoredQuoteResp = xmlLoad(strStoredQuoteResp, strFunctionName)
    Set xmlStoredQuoteNode = domStoredQuoteResp.firstChild
    errCheckXMLResponseNode xmlStoredQuoteNode, , True
    Set xmlStoredQuoteNodeList = domStoredQuoteResp.selectSingleNode("RESPONSE/STOREDQUOTATIONLIST")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Amount Requested & LTV for accepted sub-quote
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strAcceptedQuoteNumber As String
    Dim strStoredQuoteNumber As String
    Dim strQuotationMortgageSubQuoteNo As String
    Dim strMortgageSubQuoteMortgageSubQuoteNo As String
    Dim xmlLoanComponentNode As IXMLDOMNode
    
    strAcceptedQuoteNumber = xmlGetNodeText(xmlAppNode, "ACCEPTEDQUOTENUMBER")
    'DPF 28/10/2002 (AQR:  BMIDS00560) - check to see if we actually have an accepted quote
    'we can't carry on without one!
    If strAcceptedQuoteNumber = "" Then
        errThrowError strFunctionName, oeNoAcceptedQuote
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'go thru list for correct mortgage sub quote
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each xmlStoredQuoteNode In xmlStoredQuoteNodeList.childNodes
        strStoredQuoteNumber = xmlGetNodeText(xmlStoredQuoteNode, "QUOTATIONNUMBER")
        If strStoredQuoteNumber = strAcceptedQuoteNumber Then
            strQuotationMortgageSubQuoteNo = xmlGetNodeText(xmlStoredQuoteNode, "MORTGAGESUBQUOTENUMBER")
        End If
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'get the  mortgage sub quote
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set domReq = New FreeThreadedDOMDocument40
    domReq.validateOnParse = False
    domReq.setProperty "NewParser", True
      
    Dim xmlMortgageSubQuoteNode As IXMLDOMNode
    Dim domMortgageSubQuoteRespDoc As FreeThreadedDOMDocument40
    Dim objMortgageSubQuoteBO As MortgageSubQuoteBO
    Dim strMortgageSubQuote As String
    Set objMortgageSubQuoteBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    Set xmlMortgageSubQuoteNode = domReq.appendChild(domReq.createElement("MORTGAGESUBQUOTE"))
    xmlMortgageSubQuoteNode.appendChild(domReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlMortgageSubQuoteNode.appendChild(domReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo
    xmlMortgageSubQuoteNode.appendChild(domReq.createElement("MORTGAGESUBQUOTENUMBER")).Text = strQuotationMortgageSubQuoteNo

'do call
    strMortgageSubQuote = objMortgageSubQuoteBO.GetMortgageSubQuoteData(xmlMortgageSubQuoteNode.xml)
    Set domMortgageSubQuoteRespDoc = New FreeThreadedDOMDocument40
    domMortgageSubQuoteRespDoc.validateOnParse = False
    domMortgageSubQuoteRespDoc.setProperty "NewParser", True
    Set domMortgageSubQuoteRespDoc = xmlLoad(strMortgageSubQuote, strFunctionName)
'    check for error
    Set xmlMortgageSubQuoteNode = domMortgageSubQuoteRespDoc.firstChild
    errCheckXMLResponseNode xmlMortgageSubQuoteNode, , True
    Set xmlMortgageSubQuoteNode = domMortgageSubQuoteRespDoc.selectSingleNode("RESPONSE/MORTGAGESUBQUOTE")
    Dim strAmountRequested As String
    Dim strLTV As String
   strAmountRequested = xmlGetNodeText(xmlMortgageSubQuoteNode, "AMOUNTREQUESTED")
   strLTV = xmlGetNodeText(xmlMortgageSubQuoteNode, "LTV")
'
'Build xmlNode for call to CalcCostModelLTV

    Dim strCalcCostModelLTVResp As String
    Dim domCalcCostModelLTVReq As FreeThreadedDOMDocument40
    Dim domCalcCostModelLTVResp As FreeThreadedDOMDocument40
    Dim xmlCalcCostModelLTVReq As IXMLDOMNode
    Dim xmlCalcCostModel As IXMLDOMNode
    Dim xmlCalcCostModelLTV As IXMLDOMNode
    Dim xmlCustomerNodeList As IXMLDOMNode
    Dim intCustCount As Integer
    Dim xmlTempNode As IXMLDOMNode
    Dim objApplicationQuoteBO As ApplicationQuoteBO
    
    
    Set objApplicationQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    Set xmlCalcCostModel = xmlGetRequestNode(vxmlRequestParentNode)
    Set domCalcCostModelLTVReq = xmlCalcCostModel.ownerDocument
    Set xmlCalcCostModel = xmlCalcCostModel.appendChild(xmlCalcCostModel.ownerDocument.createElement("REQUEST"))
    Set xmlCalcCostModelLTV = xmlCalcCostModel.appendChild(domCalcCostModelLTVReq.createElement("LTV"))
    xmlCalcCostModelLTV.appendChild(domCalcCostModelLTVReq.createElement("APPLICATIONNUMBER")).Text = strAppNo
    xmlCalcCostModelLTV.appendChild(domCalcCostModelLTVReq.createElement("APPLICATIONFACTFINDNUMBER")).Text = strAppFactFindNo
    xmlCalcCostModelLTV.appendChild(domCalcCostModelLTVReq.createElement("AMOUNTREQUESTED")).Text = strAmountRequested
    Set xmlCustomerNodeList = xmlCalcCostModelLTV.appendChild(domCalcCostModelLTVReq.createElement("CUSTOMERLIST"))
    For intCustCount = 1 To intCustomerIndex
        Set xmlTempNode = xmlCustomerNodeList.appendChild(domCalcCostModelLTVReq.createElement("CUSTOMER"))
        xmlTempNode.appendChild(domCalcCostModelLTVReq.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber(intCustCount)
        xmlTempNode.appendChild(domCalcCostModelLTVReq.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber(intCustCount)
    Next
    
    'BMIDS00938 MDC 18/11/2002 - CC014
'    If strTypeOfMortgage = "30" Then
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Check Re-instatement value of property if Further Advance
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Dim dblReinstatementAmount As Double
'        Dim dblTestAmount As Double
'        Dim xmlAddressReq As FreeThreadedDOMDocument40
'        Dim xmlCustomerAddress As IXMLDOMNode
'        Dim xmlCustomerAddressList As IXMLDOMNode
'        Dim xmlThisCustomerAddress
'        Dim strTempResponse As String
'        Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
'        Dim xmlAddress As IXMLDOMNode
'        Dim objCustomerBO As CustomerBO
'        dblReinstatementAmount = 0#
'        Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
'        For intCustCount = 1 To intCustomerIndex
'            ' Set up request to get customer address
'            Set xmlCustomerAddress = xmlGetRequestNode(vxmlRequestParentNode)
'            Set xmlAddressReq = xmlCustomerAddress.ownerDocument
'            Set xmlCustomerAddress = xmlCustomerAddress.appendChild(xmlCustomerAddress.ownerDocument.createElement("REQUEST"))
'            Set xmlCustomerAddress = xmlCustomerAddress.appendChild(xmlAddressReq.createElement("CUSTOMERADDRESS"))
'            xmlCustomerAddress.appendChild(xmlAddressReq.createElement("CUSTOMERNUMBER")).Text = strCustomerNumber(intCustCount)
'            xmlCustomerAddress.appendChild(xmlAddressReq.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNumber(intCustCount)
'        ' Only interested in current address
'            xmlCustomerAddress.appendChild(xmlAddressReq.createElement("ADDRESSTYPE")).Text = "1"
'        ' Get the  current address sequence number for this customer
'            strTempResponse = objCustomerBO.FindCustomerAddressList(xmlCustomerAddress.xml)
'            Set xmlTempResponseDoc = xmlLoad(strTempResponse, strFunctionName)
'            errCheckXMLResponseNode xmlTempResponseDoc.firstChild, , True
'            Set xmlCustomerAddressList = xmlTempResponseDoc.selectSingleNode("RESPONSE/CUSTOMERADDRESSLIST")
'            If Not xmlCustomerAddressList Is Nothing Then
'             ' must loop thru in case there is more than one current address for this customer - tho unlikely
'             For Each xmlThisCustomerAddress In xmlCustomerAddressList.childNodes
'                 Set xmlTempNode = xmlCustomerAddress.appendChild(xmlAddressReq.createElement("CUSTOMERADDRESSSEQUENCENUMBER"))
'                 xmlTempNode.Text = xmlGetNodeText(xmlThisCustomerAddress, "CUSTOMERADDRESSSEQUENCENUMBER")
'            ' Get the customer address details to find the reinstatement amount
'                 strTempResponse = objCustomerBO.GetCustomerAddress(xmlCustomerAddress.xml)
'                 Set xmlTempResponseDoc = xmlLoad(strTempResponse, strFunctionName)
'                 errCheckXMLResponseNode xmlTempResponseDoc.firstChild, , True
'                 Set xmlAddress = xmlTempResponseDoc.selectSingleNode("RESPONSE/CUSTOMERADDRESS/CURRENTPROPERTY")
'                 dblTestAmount = xmlGetNodeAsDouble(xmlAddress, "REINSTATEMENTAMOUNT")
'            ' Make sure that reinstatement amount is the maximum
'                 If dblTestAmount > dblReinstatementAmount Then
'                   dblReinstatementAmount = dblTestAmount
'                 End If
'                 xmlCustomerAddress.removeChild xmlTempNode 'i.e. "CUSTOMERADDRESSSEQUENCENUMBER"
'              Next
'             End If
'        Next
'        If dblReinstatementAmount > 0 And _
'            dblReinstatementAmount < xmlGetAttributeAsDouble(xmlValReportNode, "REINSTATEMENTVALUE") Then
'            ' Do the call to Calc Cost Model LTV
'            strCalcCostModelLTVResp = objApplicationQuoteBO.CalcCostModelLTV(xmlCalcCostModel.xml)
'            ' Load the response and check it for errors
'            Set domCalcCostModelLTVResp = xmlLoad(strCalcCostModelLTVResp, strFunctionName)
'            Set xmlCalcCostModel = domCalcCostModelLTVResp.firstChild
'            errCheckXMLResponseNode xmlCalcCostModel, , True
'            Set xmlCalcCostModel = domCalcCostModelLTVResp.selectSingleNode("RESPONSE/LTV")
'        Else
'            Set xmlCalcCostModel = Nothing
'        End If
'    Else
    'BMIDS00938 MDC 18/11/2002 - End
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'check the property value and recalculate LTV if necessary
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
        'BMIDS589 GHun 20/06/2003 PURCHASEPRICEORESTIMATEDVALUE is a Node, not an attribute
        'If Abs((xmlGetAttributeAsDouble(xmlAppNode, "PURCHASEPRICEORESTIMATEDVALUE") +
        If Abs((xmlGetNodeAsDouble(xmlAppNode, "PURCHASEPRICEORESTIMATEDVALUE") + _
                dblDiscountAmount) _
            - dblPresentValuation) > 1# _
        Then
        'BMIDS589 End
        
            ' Do the call to Calc Cost Model LTV
            strCalcCostModelLTVResp = objApplicationQuoteBO.CalcCostModelLTV(xmlCalcCostModel.xml)
            ' Load the response and check it for errors
            Set domCalcCostModelLTVResp = New FreeThreadedDOMDocument40
            domCalcCostModelLTVResp.validateOnParse = False
            domCalcCostModelLTVResp.setProperty "NewParser", True
            Set domCalcCostModelLTVResp = xmlLoad(strCalcCostModelLTVResp, strFunctionName)
            Set xmlCalcCostModel = domCalcCostModelLTVResp.firstChild
            errCheckXMLResponseNode xmlCalcCostModel, , True
            Set xmlCalcCostModel = domCalcCostModelLTVResp.selectSingleNode("RESPONSE/LTV")
            
        Else
            Set xmlCalcCostModel = Nothing
        End If
        
    ' End If    'BMIDS00938 MDC 18/11/2002
    
    ''''''''''''''''''''''''''''''''''''''''''
    ' Report changed LTV
    ''''''''''''''''''''''''''''''''''''''''''
    If Not xmlCalcCostModel Is Nothing Then
        'BS BM0515 20/05/03 Changed to Doubles
        'BS BM0515 19/05/03 Re-introduce numerical comparison (but test > 0) because comparing text
        'values means 50 and 50.00 are considered different
'        If Abs(ConvertAssistEx.CSafeLng(xmlCalcCostModel.Text) - _
'               ConvertAssistEx.CSafeLng(strLTV)) > 0 _
'        Then
        If Abs(ConvertAssistEx.CSafeDbl(xmlCalcCostModel.Text) - _
               ConvertAssistEx.CSafeDbl(strLTV)) > 0 _
        Then
        'BM0515 strLTVChanged should be set for any difference, not only > 1%
        'If Abs(ConvertAssistEx.CSafeLng(xmlCalcCostModel.Text) - _
        '       ConvertAssistEx.CSafeLng(strLTV)) > 1 _
        'Then
        'If xmlCalcCostModel.Text <> strLTV Then
        'BM0515 End
        'BS BM0515 End 19/05/03
            strLTVChanged = "1"
        End If
    End If
' Indicate change to LTV in Response node

    Set domResp = New FreeThreadedDOMDocument40
    domResp.validateOnParse = False
    domResp.setProperty "NewParser", True
    Set xmlElem = domResp.appendChild(domResp.createElement("LTV"))
    xmlElem.setAttribute "LTVCHANGED", strLTVChanged
    
    'BMIDS589 GHun 20/06/2003 Only return old and new LTV if the LTV has changed
    If strLTVChanged = "1" Then
    'BMIDS589 End
        
        'BS BM0291 12/02/03
        xmlElem.setAttribute "OLDLTV", strLTV
        xmlElem.setAttribute "NEWLTV", xmlCalcCostModel.Text
        'BS BM0291 End 12/02/03
    
    'BMIDS589 GHun 20/06/2003
    End If
    'BMIDS589 End
    
    vxmlResponseNode.appendChild xmlElem
  
ValidateValuationReportExit:

    Set xmlValuationNode = Nothing
    Set xmlCustomerList = Nothing
    Set domReq = Nothing
    Set xmlElem = Nothing
    Set domResp = Nothing
    Set xmlTempRespNode = Nothing
    Set objApplicationBO = Nothing
    Set domAppReq = Nothing
    Set xmlAppNode = Nothing
    Set xmlAppReq = Nothing
    Set domAppResp = Nothing
    Set xmlAppNode = Nothing
    Set xmlValInstructionReq = Nothing
    Set xmlDOMVal = Nothing
    Set xmlValElem = Nothing
    Set objNewPropertyBO = Nothing
    Set xmlNewPropertyReq = Nothing
    Set domNewPropertyReq = Nothing
    Set xmlNewPropertyNode = Nothing
    Set domNewPropertyResp = Nothing
    Set objQuotationBO = Nothing
    Set xmlStoredQuoteReq = Nothing
    Set domStoredQuoteReq = Nothing
    Set xmlStoredQuoteNode = Nothing
    Set domStoredQuoteResp = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set xmlMortgageSubQuoteNode = Nothing
    Set domMortgageSubQuoteRespDoc = Nothing
    Set objApplicationQuoteBO = Nothing
    Set domCalcCostModelLTVReq = Nothing
    Set xmlCalcCostModel = Nothing
    Set xmlCalcCostModelLTV = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlTempNode = Nothing
    'BMIDS00938 MDC 18/11/2002 - CC014
    'Set xmlCustomerAddress = Nothing
    'Set xmlAddressReq = Nothing
    'Set xmlTempResponseDoc = Nothing
    'Set xmlCustomerAddressList = Nothing
    'BMIDS00938 MDC 18/11/2002 - End
    Set domCalcCostModelLTVResp = Nothing
    Set xmlCalcCostModel = Nothing
    Set domResp = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IValuationBO_UpdateValuerInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationBO_UpdateValuerInstructions"
    
    'CHECK HERE FOR PRESENCE OF Primarys
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ValuationTxBO")
        
    End If
    
    gobjTxBO.UpdateValuerInstructions vxmlRequestNode, vxmlResponseNode
    
UpdateValuerInstructionsExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_CreateValuerInstructions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified VALUERINSTRUCTION entities.
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

    On Error GoTo CreateValuerInstructionsExit
    
    Const strFunctionName As String = "IValuationBO_CreateValuerInstructions"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "VALUERINSTRUCTION"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ValuationTxBO")
    End If
    
    gobjTxBO.CreateValuerInstructions vxmlRequestNode, vxmlResponseNode
    
CreateValuerInstructionsExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_FindValuerInstructionList( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo FindValuerInstructionListExit
    
    Const strFunctionName As String = "IValuationBO_FindValuerInstructionList"
    
    Dim xmlValuerInstructionNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    
    Dim strAttribValue As String
      
    Set xmlValuerInstructionNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUERINSTRUCTION")
 
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")

    End If
  
    gobjDO.FindValuerInstructionList xmlValuerInstructionNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    

FindValuerInstructionListExit:

    Set xmlValuerInstructionNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IValuationBO_GetValuationFee( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetValuationFeeExit
    
    Const strFunctionName As String = "IValuationBO_GetValuationFee"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisGetValuationFeeNode As IXMLDOMNode
    Dim xmlGetValuationFeeNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    
    Dim strAttribValue As String
      
    Set xmlGetValuationFeeNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUERINSTRUCTION")
 
    'Set xmlThisRequestDoc = New DOMDocument

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
  
    gobjDO.GetValuationFee xmlGetValuationFeeNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    

GetValuationFeeExit:

    Set xmlGetValuationFeeNode = Nothing
    Set xmlThisGetValuationFeeNode = Nothing
    Set xmlThisRequestDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IValuationBO_ProcessValuerBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Processes the batch passed in by getting all relevant instructions and putting them
'   on a queue for processing
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo ProcessValuerBatchExit

    Const strFunctionName As String = "IValuationBO_ProcessValuerBatch"
    
    Dim objBatchSchedBO As BatchScheduleBO
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempBatchRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlBatchInstruction As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlScheduleNode As IXMLDOMNode
    Dim xmlBatch As IXMLDOMNode
    Dim xmlInstructionList As IXMLDOMNodeList
    Dim xmlInstruction As IXMLDOMNode
    Dim xmlBatchAuditRequest As IXMLDOMNode
    Dim xmlBatchAudit As IXMLDOMNode
    Dim xmlValuation As IXMLDOMNode
    Dim objNTTxBO As IValuationNTTxBO
    
    Dim lngErrorNo As Long
    Dim strSource As String
    Dim strDescription As String
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strInstructSeqNo As String
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String

    'Validation
    Set xmlBatch = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatch, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatch, "BATCHRUNNUMBER")

    'Create request to get Batch data
    Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "GetBatchSchedule"
    Set xmlScheduleNode = xmlTempBatchRequest.ownerDocument.createElement("BATCHSCHEDULE")
    xmlCopyMandatoryAttribute xmlBatch, xmlScheduleNode, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlBatch, xmlScheduleNode, "BATCHRUNNUMBER"
    xmlTempBatchRequest.appendChild xmlScheduleNode

    'Create and call the Batch Scheduler component
    Set objBatchSchedBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
       
    ' Check if batch has been cancelled
    Dim strBatchStatus As String
    Dim strFailedRunNo As String
    Dim intReceiptDays As Integer
    
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHSCHEDULE")
    strBatchStatus = xmlGetMandatoryAttributeText(xmlNode, "STATUS")
    strFailedRunNo = xmlGetAttributeText(xmlNode, "FAILEDBATCHRUNNUMBER")
    
    If IsValidationType("BatchScheduleStatus", CSafeInt(strBatchStatus), "CA") = False Then
    
        If gobjDO Is Nothing Then
            Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
        End If
        
        Set xmlTempResponse = Nothing
        Set xmlTempResponse = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTIONS")

        ' Check if this is a new run
        If CSafeLng(strFailedRunNo) = 0 Then
            intReceiptDays = GetMandatoryGlobalParamAmount("ValuationReceiptDays")
            
            
            ' Get the valuer instructions that are outstanding
            Set xmlBatchInstruction = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTION")
            Set xmlNode = xmlTempResponseDoc.createAttribute("VALUATIONRECEIPTDAYS")
            xmlBatchInstruction.setAttribute "VALUATIONRECEIPTDAYS", CStr(intReceiptDays)
            
            On Error Resume Next
            
            gobjDO.FindInstructionsForBatch xmlBatchInstruction, xmlTempResponse
            
            lngErrorNo = Err.Number
            strSource = Err.Source
            strDescription = Err.Description
            
            On Error GoTo ProcessValuerBatchExit
            
            ' PSC 25/02/02 SYS4097 - Start
            If lngErrorNo <> 0 Then
                If errGetOmigaErrorNumber(lngErrorNo) = oeRecordNotFound Then
                    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                    Set xmlValuation = vxmlRequestNode.ownerDocument.createElement("VALUATION")
                    xmlTempRequest.appendChild xmlValuation
                    xmlCopyAttribute xmlBatch, xmlValuation, "BATCHNUMBER"
                    xmlCopyAttribute xmlBatch, xmlValuation, "BATCHRUNNUMBER"
                    CompleteBatch xmlTempRequest, vxmlResponseNode
                Else
                    Err.Raise lngErrorNo, strSource, strDescription
                End If
            End If
            ' PSC 25/02/02 SYS4097 - End


        Else
             ' Get the valuer instructions that have failed
            Set xmlBatchInstruction = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTION")
            xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchInstruction, "BATCHNUMBER"
            xmlSetAttributeValue xmlBatchInstruction, "BATCHRUNNUMBER", strFailedRunNo
            
            On Error Resume Next
            
            gobjDO.FindFailedInstructions xmlBatchInstruction, xmlTempResponse
                          
            lngErrorNo = Err.Number
            strSource = Err.Source
            strDescription = Err.Description
            
            On Error GoTo ProcessValuerBatchExit
            
            ' PSC 25/02/02 SYS4097 - Start
            If lngErrorNo <> 0 Then
                If errGetOmigaErrorNumber(lngErrorNo) = oeRecordNotFound Then
                    errThrowError strFunctionName, 575
                Else
                    Err.Raise lngErrorNo, strSource, strDescription
                End If
            
            End If
            ' PSC 25/02/02 SYS4097 - End
        
        End If
                
        Set xmlInstructionList = xmlTempResponse.selectNodes(".//BATCHINSTRUCTION")
        
        ' PSC 25/02/02 SYS4097
        If xmlInstructionList.length > 0 Then
        
            ' Update the batch schedule with the record count
            xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "UpdateBatchSchedule"
            xmlSetAttributeValue xmlScheduleNode, "TOTALNUMBEROFRECORDS", CStr(xmlInstructionList.length)
            
            strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
            Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
                
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            
            Set objNTTxBO = gobjContext.CreateInstance(App.Title & ".ValuationNTTxBO")
            
            ' Set up common data in Batch Audit
            Set xmlBatchAuditRequest = vxmlRequestNode.cloneNode(False)
            xmlSetAttributeValue xmlBatchAuditRequest, "OPERATION", "CreateBatchAudit"
            Set xmlBatchAudit = vxmlRequestNode.ownerDocument.createElement("BATCHAUDIT")
            xmlBatchAuditRequest.appendChild xmlBatchAudit
            xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchAudit, "BATCHNUMBER"
            xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchAudit, "BATCHRUNNUMBER"
            
            'SG 25/02/02 SYS4097
            Dim strQueueName As String
            strQueueName = GetMandatoryGlobalParamString("ValuationQueueName")
            'SG 25/02/02 SYS4097
                        
            ' For each instruction create a batch audit record, valuation batch audit
            ' record and send it to the queue
            For Each xmlInstruction In xmlInstructionList
                    
                ' Copy batch keys onto instruction
                xmlCopyMandatoryAttribute xmlScheduleNode, xmlInstruction, "BATCHNUMBER"
                xmlCopyMandatoryAttribute xmlScheduleNode, xmlInstruction, "BATCHRUNNUMBER"
                
                'SG 25/02/02 SYS4097
                xmlSetAttributeValue xmlInstruction, "QUEUENAME", strQueueName

                xmlTempRequest.appendChild xmlInstruction
                
                On Error Resume Next
                
                objNTTxBO.ProcessValuerBatch xmlTempRequest, vxmlResponseNode
                
                lngErrorNo = Err.Number
                strDescription = Err.Description
                strSource = Err.Source
                
                On Error GoTo ProcessValuerBatchExit
                                           
                If lngErrorNo <> 0 Then
                
                    ' Create the batch audit record in order to write the error to
                    strResponse = objBatchSchedBO.omBatchRequest(xmlBatchAuditRequest.xml)
                    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
                    
                    ' PSC 01/03/2002 SYS4214 - Start
                    strApplicationNo = xmlGetAttributeText(xmlInstruction, "APPLICATIONNUMBER")
                    strApplicationFFNo = xmlGetAttributeText(xmlInstruction, "APPLICATIONFACTFINDNUMBER")
                    strInstructSeqNo = xmlGetAttributeText(xmlInstruction, "INSTRUCTIONSEQUENCENO")
                    
                    strDescription = "Application No: " & strApplicationNo & _
                                     " Application Fact Find No: " & strApplicationFFNo & _
                                     " Instruction No: " & strInstructSeqNo & " " & strDescription
                    ' PSC 01/03/2002 SYS4214 - End
                                      
                    ' Update the batch failure count
                    Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHAUDIT")
                    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "IncrementBatchScheduleFailure"
                    xmlCopyMandatoryAttribute xmlNode, xmlScheduleNode, "BATCHAUDITGUID"
                    xmlSetAttributeValue xmlScheduleNode, "ERRORNUMBER", CStr(lngErrorNo)
                    xmlSetAttributeValue xmlScheduleNode, "ERRORSOURCE", strSource
                    xmlSetAttributeValue xmlScheduleNode, "ERRORDESCRIPTION", strDescription
                    
                    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
                    
                    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
                    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
                
                End If
                
                ' Remove this instruction from the request
                xmlTempRequest.removeChild xmlInstruction
            Next
        End If
    Else
        'JD BMIDS604 Batch is cancelled, raise an error to prevent the batch status being set as Complete
        errThrowError strFunctionName, 7025
    End If
    
ProcessValuerBatchExit:

    Set objBatchSchedBO = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempBatchRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlBatchInstruction = Nothing
    Set xmlNode = Nothing
    Set xmlScheduleNode = Nothing
    Set xmlBatch = Nothing
    Set xmlInstructionList = Nothing
    Set xmlInstruction = Nothing
    Set xmlBatchAuditRequest = Nothing
    Set xmlBatchAudit = Nothing
    Set objNTTxBO = Nothing
    Set xmlValuation = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_FindInstructionsForBatch( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieves the instructions ready for batch processing,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo FindInstructionsForBatchExit
    
    Const strFunctionName As String = "IValuationBO_FindInstructionsForBatch"
    
    Dim xmlBatchInstructionNode As IXMLDOMNode
    
    ' validate REQUEST & select BatchInstruction node
    Set xmlBatchInstructionNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHINSTRUCTION")
  
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    gobjDO.FindInstructionsForBatch xmlBatchInstructionNode, vxmlResponseNode
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindInstructionsForBatchExit:

    Set xmlBatchInstructionNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_CreateValuationAudit(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates a valuation batch audit record
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo CreateValuationAuditExit
    
    Const strFunctionName As String = "IValuationBO_CreateValuationAudit"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "VALUATIONBATCHAUDIT"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".ValuationTXBO")
    End If
    
    gobjTxBO.CreateValuationAudit vxmlRequestNode, vxmlResponseNode
    
CreateValuationAuditExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_FindFailedInstructions( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retireves the instructions which have failed batch processing
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo FindFailedInstructionsExit
    
    Const strFunctionName As String = "IValuationBO_FindFailedInstructions"
    
    Dim xmlBatchInstructionNode As IXMLDOMNode
    
    ' validate REQUEST & select BatchInstruction node
    Set xmlBatchInstructionNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHINSTRUCTION")
  
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    gobjDO.FindFailedInstructions xmlBatchInstructionNode, vxmlResponseNode
    
    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindFailedInstructionsExit:

    Set xmlBatchInstructionNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IValuationBO_ProduceValuerOutput(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Produces the financial transactions and letters for all valuers that have successful
'   instructions
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------

On Error GoTo ProduceValuerOutputExit

    Const strFunctionName As String = "IValuationBO_ProduceValuerOutput"
    
    Dim objBatchSchedBO As BatchScheduleBO
    Dim objPrintBO As omPrintBO
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlBatchElement As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlScheduleNode As IXMLDOMNode
    Dim xmlTransactionSummaryList As IXMLDOMNode
    Dim xmlTransactionListNode As IXMLDOMNode
    Dim xmlValuerSummary As IXMLDOMNode
    Dim xmlValuerSummaryList As IXMLDOMNodeList
    Dim xmlValuerTransactionList As IXMLDOMNodeList
    Dim xmlValuerTransaction As IXMLDOMNode
    Dim xmlPrintRequestNode As IXMLDOMNode
    
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String

    'Validation
    Set xmlScheduleNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATION")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlScheduleNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlScheduleNode, "BATCHRUNNUMBER")

    'Create request to get Batch data
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetBatchSchedule"
    Set xmlBatchElement = vxmlRequestNode.ownerDocument.createElement("BATCHSCHEDULE")
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHRUNNUMBER"
    xmlTempRequest.appendChild xmlBatchElement

    'Create and call the Batch Scheduler component
    Set objBatchSchedBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
    
    Set xmlTempRequest = Nothing
   
    ' Check if batch has been cancelled
    Dim strBatchStatus As String
    
    Set xmlNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHSCHEDULE")
    strBatchStatus = xmlGetMandatoryAttributeText(xmlNode, "STATUS")
    
    If IsValidationType("BatchScheduleStatus", CSafeInt(strBatchStatus), "CA") = False Then
        
        Dim strPaymentType As String
        Dim strDirectoryGUID As String
        Dim strPattern As String
        
        If gobjDO Is Nothing Then
            Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
        End If
        
        ' Get a list of successful transactions
        Set xmlTransactionListNode = vxmlRequestNode.ownerDocument.createElement("TRANSACTIONLIST")
        gobjDO.FindSuccessfulTransactionList xmlScheduleNode, xmlTransactionListNode
        
        ' Get the transaction summary for the valuers
        Set xmlTransactionSummaryList = vxmlRequestNode.ownerDocument.createElement("TRANSACTIONSUMMARYLIST")
        gobjDO.FindTransactionSummaryList xmlScheduleNode, xmlTransactionSummaryList
        
        Set xmlValuerSummaryList = xmlTransactionSummaryList.selectNodes("TRANSACTIONSUMMARY")
        
        ' Get the print configuration data
        Set xmlPrintRequestNode = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlPrintRequestNode, "OPERATION", "PrintDocument"

'       Add the print control data into the request
'       AddValuerLetterPrintControlData xmlPrintRequestNode
        Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")

        ' Produce either a BACS or cheque transaction together with a letter for each valuer
        For Each xmlValuerSummary In xmlValuerSummaryList
            
            strPaymentType = xmlGetMandatoryAttributeText(xmlValuerSummary, "PAYMENTTYPE")
            
            ' Payment type must be BACS or cheque
            If strPaymentType <> "B" And strPaymentType <> "CH" Then
                'Error for this valuer
            End If
            
            If strPaymentType = "B" Then            ' BACS
                ' Create BACS transaction
            Else                                    ' Cheque
                ' Create Cheque transaction
            End If
            
            ' Get a list of successful transactions for this valuer
            strDirectoryGUID = xmlGetMandatoryAttributeText(xmlValuerSummary, "DIRECTORYGUID")
            
            strPattern = "TRANSACTION[@DIRECTORYGUID = '" & strDirectoryGUID & "']"
            Set xmlValuerTransactionList = xmlTransactionListNode.selectNodes(strPattern)
            
            AddValuerLetterData xmlPrintRequestNode, xmlValuerSummary, xmlValuerTransactionList
                        
'            strResponse = objPrintBO.omRequest(xmlPrintRequestNode.xml)
'            errCheckXMLResponse strResponse, True
                        
            Set xmlValuerTransactionList = Nothing
        Next
        
        ' Produce a report listing all transactions processed regardless of success or failure
'        PrintValuationReport vxmlRequestNode
        
        ' Produce a report listing all transactions which have failed
'        PrintFailureReport vxmlRequestNode

        ' Reprint of cheques
'        if Reprint of cheques then
'            ReprintCheques
'        End If
    End If

ProduceValuerOutputExit:

    Set objBatchSchedBO = Nothing
    Set objPrintBO = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlBatchElement = Nothing
    Set xmlNode = Nothing
    Set xmlScheduleNode = Nothing
    Set xmlTransactionSummaryList = Nothing
    Set xmlTransactionListNode = Nothing
    Set xmlValuerSummary = Nothing
    Set xmlValuerSummaryList = Nothing
    Set xmlValuerTransactionList = Nothing
    Set xmlValuerTransaction = Nothing
    Set xmlPrintRequestNode = Nothing

    errCheckError strFunctionName, TypeName(Me)


End Sub

Private Sub AddValuerLetterData(ByVal vxmlPrintRequestNode As IXMLDOMNode, _
                                  ByVal vxmlValuerSummary As IXMLDOMNode, _
                                  ByVal vxmlTransactionList As IXMLDOMNodeList)
' header ----------------------------------------------------------------------------------
' description:
'   Produces a letter for the valuer passed in and the associated transactions
' pass:
'   vxmlPrintRequestNode
'       Node for the print request
'   vxmlValuerSummary
'       Summary data for the valuer
'   vxmlTransactionList
'       List of successful transactions for the valuer
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

On Error GoTo AddValuerLetterDataExit

    Const strFunctionName As String = "AddValuerLetterData"
    
    Dim xmlTransaction As IXMLDOMNode
    Dim xmlTemplateData As IXMLDOMNode
    Dim xmlOldTemplateData As IXMLDOMNode
    
    ' Set up the template data for this letter
    Set xmlTemplateData = vxmlPrintRequestNode.ownerDocument.createElement("TEMPLATEDATA")
    xmlTemplateData.appendChild vxmlValuerSummary.cloneNode(True)
    
    For Each xmlTransaction In vxmlTransactionList
        xmlTemplateData.appendChild xmlTransaction.cloneNode(True)
    Next
    
    ' Replace old data with new data
    Set xmlOldTemplateData = vxmlPrintRequestNode.selectSingleNode(".//TEMPLATEDATA")

    If Not xmlOldTemplateData Is Nothing Then
        vxmlPrintRequestNode.replaceChild xmlTemplateData, xmlOldTemplateData
    Else
        vxmlPrintRequestNode.appendChild xmlTemplateData
    End If
        
AddValuerLetterDataExit:

    Set xmlTransaction = Nothing
    Set xmlTemplateData = Nothing
    Set xmlOldTemplateData = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ReprintCheques()
' header ----------------------------------------------------------------------------------
' description:
'   Reprints the cheques
' pass:
'   vxmlValuerSummary
'       Summary data for the valuer
'   vxmlTransactionList
'       List of successful transactions for the valuer
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

On Error GoTo ReprintChequesExit

    Const strFunctionName As String = "ReprintCheques"
    
'    If report print is part of omiga then
'        Print By batch Or extract
'        if cheque numbers need to be configured then
'            Renumber cheques
'        End If
'
'        Send to print queue
'    End If

ReprintChequesExit:

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub CompleteBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Does the processing for the final transaction in the batch. Sets the batch to complete
'   if no transactions failed or to complete with errors if errors have occured in the batch
'   run
' pass:
'   vxmlRequestNode
'       XML request Node
'   vxmlResponseNode
'       XML response Node
' return:
'       none
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

On Error GoTo CompleteBatchExit

    Const strFunctionName As String = "CompleteBatch"
    
    Dim objBatchSchedBO As BatchScheduleBO
    Dim xmlBatchElement As IXMLDOMNode
    Dim xmlScheduleNode As IXMLDOMNode
    Dim xmlTempBatchRequest As IXMLDOMNode
    Dim xmlInstructionListNode As IXMLDOMNode
    Dim xmlInstructionList As IXMLDOMNodeList
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim colValueIDs As Collection
    
    Dim strResponse As String
    Dim strStatus As String
        
    ' Get the valuer instructions that have failed
    Set xmlScheduleNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATION")

    Set xmlBatchElement = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTION")
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHRUNNUMBER"
    Set xmlInstructionListNode = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTIONLIST")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    gobjDO.FindFailedInstructions xmlBatchElement, xmlInstructionListNode
    
    Set xmlInstructionList = xmlInstructionListNode.selectNodes(".//BATCHINSTRUCTION")
    
    ' No failures is complete, failures is complete with errors
    If xmlInstructionList.length = 0 Then
        strStatus = GetFirstComboValueId("BatchScheduleStatus", "C")
    Else
        strStatus = GetFirstComboValueId("BatchScheduleStatus", "CWE")
        'JD BMIDS604 check the error number on the failed instructions. We are looking for a 'Cancelled' app
        Dim nCancelledErr As Long
        Dim xmlCancelledNodeList As IXMLDOMNodeList
        nCancelledErr = 7025 + vbObjectError + 512
        Set xmlCancelledNodeList = xmlInstructionListNode.selectNodes("BATCHINSTRUCTION[@ERRORNUMBER='" & CStr(nCancelledErr) & "']")
        If xmlCancelledNodeList.length > 0 Then
            strStatus = GetFirstComboValueId("BatchScheduleStatus", "CA")
        End If
    End If
        
    Set objBatchSchedBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    
    ' Produce batch schedule output report
    IValuationBO_ProduceValuerOutput vxmlRequestNode, vxmlResponseNode

    Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
    Set xmlBatchElement = vxmlRequestNode.ownerDocument.createElement("BATCHSCHEDULE")
    xmlTempBatchRequest.appendChild xmlBatchElement
    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "SetBatchStatus"
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHRUNNUMBER"
    xmlSetAttributeValue xmlBatchElement, "STATUS", strStatus
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
    
    'SYS4176 Now Print the Report
    
    xmlChangeNodeName xmlBatchElement, "BATCHSCHEDULE", "BATCH"
    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "PrintDocuments"
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
    
CompleteBatchExit:

    If errGetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

    Set objBatchSchedBO = Nothing
    Set xmlBatchElement = Nothing
    Set xmlScheduleNode = Nothing
    Set xmlTempBatchRequest = Nothing
    Set xmlInstructionList = Nothing
    Set xmlTempResponseDoc = Nothing
    Set colValueIDs = Nothing
    Set xmlInstructionList = Nothing


    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub AddValuerLetterPrintControlData(ByVal vxmlPrintRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Gets the configuration data used for printing
' pass:
'   vxmlPrintRequestNode
'       XML print request Node
' return:
'       Print configuration data
' exceptions:
'
'------------------------------------------------------------------------------------------

On Error GoTo AddValuerLetterPrintControlDataExit

    Const strFunctionName As String = "AddValuerLetterPrintControlData"
    
    Dim objPrintBO As omPrintBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlPrintAttributes As IXMLDOMNode
    Dim xmlControlData As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    
    Dim lngTemplateId As Long
    
    lngTemplateId = GetMandatoryGlobalParamAmount("PrintValuation")
    
    Set xmlTempRequest = vxmlPrintRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
    Set xmlPrintAttributes = vxmlPrintRequestNode.ownerDocument.createElement("FINDATTRIBUTES")
    xmlTempRequest.appendChild xmlPrintAttributes
    xmlSetAttributeValue xmlPrintAttributes, "DOCUMENTID", CStr(lngTemplateId)
    
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    
    Dim strResponse  As String
    
    strResponse = objPrintBO.omRequest(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc, , True
    
    Set xmlPrintAttributes = Nothing
    Set xmlPrintAttributes = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
    
    ' Do some sort of validation here such as is the template active etc
    
    Set xmlControlData = vxmlPrintRequestNode.ownerDocument.createElement("CONTROLDATA")
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "HOSTTEMPLATEID", "DOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DPSTEMPLATEID ", "DPSDOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DEFAULTCOPIES ", "COPIES"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "REMOTEPRINTERLOCATION ", "PRINTER"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "PRINTERDESTINATIONTYPE ", "DESTINATIONTYPE "
    
    vxmlPrintRequestNode.appendChild xmlControlData
    
AddValuerLetterPrintControlDataExit:

    Set objPrintBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPrintAttributes = Nothing
    Set xmlControlData = Nothing
    Set xmlTempResponseDoc = Nothing


    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub PrintFailureReport(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Gets the configuration data used for printing
' pass:
'   vxmlRequestNode
'       XML request Node
' return:
'       Print configuration data
' exceptions:
'
'------------------------------------------------------------------------------------------

On Error GoTo PrintFailureReportExit

    Const strFunctionName As String = "PrintFailureReport"
    
    Dim objPrintBO As omPrintBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlPrintAttributes As IXMLDOMNode
    Dim xmlControlData As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlPrintRequestNode As IXMLDOMNode
    Dim xmlTemplateData As IXMLDOMNode
    Dim xmlInstruction As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlInstructionList As IXMLDOMNodeList
    Dim xmlInstructionListNode As IXMLDOMNode
    Dim xmlBatchElement As IXMLDOMNode
    Dim xmlScheduleNode As IXMLDOMNode
    
    Dim lngTemplateId As Long
    
    Set xmlScheduleNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATION")

    ' Get a list of failed instructions
    Set xmlBatchElement = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTION")
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHNUMBER"
    xmlCopyMandatoryAttribute xmlScheduleNode, xmlBatchElement, "BATCHRUNNUMBER"

    Set xmlInstructionListNode = vxmlRequestNode.ownerDocument.createElement("BATCHINSTRUCTIONLIST")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    gobjDO.FindFailedInstructions xmlBatchElement, xmlInstructionListNode
    
    Set xmlInstructionList = xmlInstructionListNode.selectNodes(".//BATCHINSTRUCTION")

    
    ' Get document attributes
    lngTemplateId = GetMandatoryGlobalParamAmount("PrintValuationFailure")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
    Set xmlPrintAttributes = vxmlRequestNode.ownerDocument.createElement("FINDATTRIBUTES")
    xmlTempRequest.appendChild xmlPrintAttributes
    xmlSetAttributeValue xmlPrintAttributes, "DOCUMENTID", CStr(lngTemplateId)
    
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    
    Dim strResponse  As String
    
    strResponse = objPrintBO.omRequest(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc, , True
    
    Set xmlPrintAttributes = Nothing
    Set xmlPrintAttributes = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
    
    ' Do some sort of validation here such as is the template active etc
    
    Set xmlPrintRequestNode = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlPrintRequestNode, "OPERATION", "PrintDocument"
    
   ' Set up document control data
    Set xmlControlData = xmlPrintRequestNode.ownerDocument.createElement("CONTROLDATA")
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "HOSTTEMPLATEID", "DOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DPSTEMPLATEID ", "DPSDOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DEFAULTCOPIES ", "COPIES"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "REMOTEPRINTERLOCATION ", "PRINTER"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "PRINTERDESTINATIONTYPE ", "DESTINATIONTYPE "
    
    xmlPrintRequestNode.appendChild xmlControlData
    
    Set xmlTemplateData = vxmlRequestNode.ownerDocument.createElement("TEMPLATEDATA")
    xmlPrintRequestNode.appendChild xmlTemplateData
        
    For Each xmlInstruction In xmlInstructionList
        xmlTemplateData.appendChild xmlInstruction
    Next
    
    ' Send data to print
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    strResponse = objPrintBO.omRequest(xmlPrintRequestNode.xml)
    errCheckXMLResponse strResponse, True
    
PrintFailureReportExit:

    Set objPrintBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPrintAttributes = Nothing
    Set xmlControlData = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlPrintRequestNode = Nothing
    Set xmlTemplateData = Nothing
    Set xmlInstruction = Nothing
    Set xmlNode = Nothing
    Set xmlInstructionList = Nothing
    Set xmlInstructionListNode = Nothing
    Set xmlBatchElement = Nothing
    Set xmlScheduleNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub PrintValuationReport(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Gets the configuration data used for printing
' pass:
'   vxmlRequestNode
'       XML request Node
' return:
'       Print configuration data
' exceptions:
'
'------------------------------------------------------------------------------------------

On Error GoTo PrintValuationReportExit

    Const strFunctionName As String = "PrintValuationReport"
    
    Dim objPrintBO As omPrintBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlPrintAttributes As IXMLDOMNode
    Dim xmlControlData As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlPrintRequestNode As IXMLDOMNode
    Dim xmlTemplateData As IXMLDOMNode
    Dim xmlTransaction As IXMLDOMNode
    Dim xmlTransactionListNode As IXMLDOMNode
    Dim xmlTransactionList As IXMLDOMNodeList
    Dim xmlScheduleNode As IXMLDOMNode
    
    Dim lngTemplateId As Long
    
    Set xmlScheduleNode = xmlGetMandatoryNode(vxmlRequestNode, "VALUATION")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".ValuationDO")
    End If
    
    ' Get list of all transactions
    Set xmlTransactionListNode = vxmlRequestNode.ownerDocument.createElement("TRANSACTIONLIST")
    gobjDO.FindTransactionList xmlScheduleNode, xmlTransactionListNode
    
    Set xmlTransactionList = xmlTransactionListNode.selectNodes(".//TRANSACTION")

    ' Get document attributes
    lngTemplateId = GetMandatoryGlobalParamAmount("PrintValuationReport")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
    Set xmlPrintAttributes = vxmlRequestNode.ownerDocument.createElement("FINDATTRIBUTES")
    xmlTempRequest.appendChild xmlPrintAttributes
    xmlSetAttributeValue xmlPrintAttributes, "DOCUMENTID", CStr(lngTemplateId)
    
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    
    Dim strResponse  As String
    
    strResponse = objPrintBO.omRequest(xmlTempRequest.xml)
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlTempResponseDoc, , True
    
    Set xmlPrintAttributes = Nothing
    Set xmlPrintAttributes = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
    
    ' Do some sort of validation here such as is the template active etc
    
    Set xmlPrintRequestNode = vxmlRequestNode.cloneNode(False)
    
    ' Set up control data
    Set xmlControlData = xmlPrintRequestNode.ownerDocument.createElement("CONTROLDATA")
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "HOSTTEMPLATEID", "DOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DPSTEMPLATEID ", "DPSDOCUMENTID"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DEFAULTCOPIES ", "COPIES"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "REMOTEPRINTERLOCATION ", "PRINTER"
    xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "PRINTERDESTINATIONTYPE ", "DESTINATIONTYPE "
    Set xmlPrintRequestNode = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlPrintRequestNode, "OPERATION", "PrintDocument"
    
    Set xmlTemplateData = vxmlRequestNode.ownerDocument.createElement("TEMPLATEDATA")
    xmlPrintRequestNode.appendChild xmlTemplateData
    
    For Each xmlTransaction In xmlTransactionList
        xmlTemplateData.appendChild xmlTransaction
    Next
    
    ' Send to print
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    strResponse = objPrintBO.omRequest(xmlPrintRequestNode.xml)
    errCheckXMLResponse strResponse, True
    
PrintValuationReportExit:

    Set objPrintBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPrintAttributes = Nothing
    Set xmlControlData = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlPrintRequestNode = Nothing
    Set xmlTemplateData = Nothing
    Set xmlTransaction = Nothing
    Set xmlTransactionListNode = Nothing
    Set xmlTransactionList = Nothing
    Set xmlScheduleNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, ByVal vxmlParentNode As IXMLDOMNode)

    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib
    
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Text = "APValnRBTemplate"
    xmlNode.Attributes.setNamedItem xmlAttrib
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
    
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next
    
GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub RunValuationRules(ByVal vxmlRequestNode As IXMLDOMNode, _
                              ByVal vxmlResponseNode As IXMLDOMNode)
                              
    On Error GoTo RunValuationRulesExit
    Const strFunctionName As String = "RunValuationRules"

    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40, _
        xmlThisResponseDoc As FreeThreadedDOMDocument40, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationNode As IXMLDOMNode, _
        xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement

    Dim objAPRules As Object, _
        gobjAppBOProc As Object
    
    Dim strComments As String
    Dim iUnderReviewInd As Integer
    iUnderReviewInd = 0
        
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "RunValuationReportRules"
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("VALUATION")
    
    GetApplicationData xmlApplicationNode, xmlThisRequestNode
        
    Set objAPRules = gobjContext.CreateInstance("omAPRules.APRulesBO")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    xmlThisResponseDoc.loadXML objAPRules.RunRules(xmlThisRequestDoc.xml)
    Set objAPRules = Nothing
    
    Set xmlResponseList = _
        xmlThisResponseDoc.selectNodes("RESPONSE/VALUATIONREPORTRULE[@RESULT='0']")
    
    If xmlResponseList.length > 0 Then
        
        iUnderReviewInd = 1
        
        For Each xmlThisNode In xmlResponseList
        
            If Len(strComments) > 0 Then
                strComments = strComments & ";"
            End If
            
            strComments = strComments & xmlGetAttributeText(xmlThisNode, "MESSAGE")
        
        Next
        
        ' build CreateApplicationReviewHistory request =================================
        
        ' strip child nodes from request
        For Each xmlThisNode In xmlThisRequestNode.childNodes
            xmlThisRequestNode.removeChild xmlThisNode
        Next
        
        ' set operation value
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = _
            "CreateApplicationReviewHistory"
    
        Set xmlThisElem = xmlThisRequestDoc.createElement("APPLICATIONREVIEWHISTORY")
        xmlCopyAttribute xmlApplicationNode, xmlThisElem, "APPLICATIONNUMBER"
        xmlCopyAttribute xmlApplicationNode, xmlThisElem, "APPLICATIONFACTFINDNUMBER"
        xmlThisElem.setAttribute "REVIEWREASON", 10
        xmlSetSysDateToNodeAttrib xmlThisElem, "REVIEWDATETIME"
        xmlThisElem.setAttribute _
            "REVIEWUSERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
        xmlThisElem.setAttribute _
            "REVIEWUNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
        xmlThisElem.setAttribute "REVIEWCOMMENTS", strComments

        ' Complete the XML request by appending
        xmlThisRequestNode.appendChild xmlThisElem
        
        Debug.Print "omAppProc.omAppProcBO Request"
        Debug.Print xmlThisRequestDoc.xml

        Set gobjAppBOProc = gobjContext.CreateInstance("omAppProc.omAppProcBO")
        
        xmlThisResponseDoc.loadXML _
            gobjAppBOProc.OmAppProcRequest(xmlThisRequestDoc.xml)
        
        Debug.Print "omAppProc.omAppProcBO Response"
        Debug.Print xmlThisResponseDoc.xml
        
        Set gobjAppBOProc = Nothing
        
        errCheckXMLResponseNode _
            xmlThisResponseDoc.selectSingleNode("RESPONSE"), _
            vxmlResponseNode, True
    
    End If
    
    'Attach the following to vxmlResponseNode
    Set xmlThisElem = xmlThisResponseDoc.createElement("APPSTATUS")
    xmlThisElem.setAttribute "UNDERREVIEWIND", iUnderReviewInd
    xmlThisElem.setAttribute "REVIEWMESSAGETEXT", strComments
    
    vxmlResponseNode.appendChild xmlThisElem
    
RunValuationRulesExit:

    Set objAPRules = Nothing
    Set gobjAppBOProc = Nothing
    
    Set xmlResponseList = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlThisNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

