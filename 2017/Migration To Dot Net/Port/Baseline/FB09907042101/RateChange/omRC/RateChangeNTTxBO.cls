VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "RateChangeNTTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------------------------------
'Workfile:      RateChangeNTTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Rate Change (New) Transactional Business Object.
'               MTSTransactionMode: Requires New Transaction
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MC     20/08/2001  Created
'MC     06/12/2001  SYS3018 - Fix for Transactioning/SQL Server locking problems
'MC     07/12/2001  SYS3440 - Various Batch fixes
'MC     20/12/2001  SYS3541 - Write error info to event log
'PSC    25/02/2002  SYS4097 - Update ProcessApplication to have QueueName passed in
'DRC    02/04/2002  SYS4343 - Change Attribute name to match database field name
'SA     04/04/2002  SYS4340 - Change call in ProcessApplication to pass string instead of node.
'-------------------------------------------------------------------------------------------------------
Option Explicit

Implements IRateChangeNTTxBO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Enum LOCKSTATUS
    lsLocked = 1
    lsUnlocked = 0
    lsError = -1
End Enum

'-------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MC     09/01/2003  BM0230 - Changes for failed batch runs in ProcessApplication
'GHun   19/02/2003  BM0367 - Add future dated EXECUTEAFTERDATE before adding a message to the queue
'GHun   15/08/2003  BMIDS622 - Restructure
'JD     22/11/2004  BMIDS604 - If batch is cancelled throw an error to indicate this.
'-------------------------------------------------------------------------------------------------------

'BMIDS622 No longer used, but cannot be deleted because OOSS will just replace it
Private Sub IRateChangeNTTxBO_ProcessApplication(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    ProcessApplication
' Author:       Mark Coleman
' Date:         05/09/2001
' Purpose:      Initialise a single item in a Rate Change batch job
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
' Notes:        This method creates a BatchAudit record for a single
'               application affected by a Rate Change. It also creates
'               RateChangeBatchAudit records for each Loan Component the
'               application quote has. It the posts a message to the
'               Rate Change message queue for processing.
'------------------------------------------------------------------------

'
'On Error GoTo ProcessApplicationVbErr
'
'Dim strFunctionName As String
'strFunctionName = "IRateChangeNTTxBO_ProcessApplication"
'
'Dim objBatchSchedBO As BatchScheduleBO
'Dim objMessageQueue As IomMessageQueue
'Dim objRateChangeDO As IRateChangeDO
''SYS4340 Should be declared as BO not IBO
''Dim objMortgageProductBO As IMortgageProductBO
'Dim objMortgageProductBO As MortgageProductBO
'Dim xmlDoc As FreeThreadedDOMDocument40
'Dim xmlTempRequest As IXMLDOMNode
'Dim xmlTempResponse As IXMLDOMNode
'Dim xmlMPRequestElem As IXMLDOMElement
''SYS4340
''Dim xmlMPResponse As IXMLDOMNode
'Dim xmlMPResponse As FreeThreadedDOMDocument40
'Dim xmlNode As IXMLDOMNode
'Dim xmlNodeList As IXMLDOMNodeList
'Dim xmlLoanCompNode As IXMLDOMNode
'Dim xmlMessage As IXMLDOMElement
'Dim xmlRateChangeNode As IXMLDOMNode
'
'Dim strResponse As String
'Dim strQueueName As String
'Dim strBatchNumber As String
'Dim strBatchRunNumber As String
''SYS3541 MDC 20/12/2001
'Dim strErrMsg As String
'Dim strBatchAuditGUID As String
''SYS3541 End
'Dim xmlMPResponseNode As IXMLDOMNode    'SYS4340
'
''BM0230 MDC 08/01/2003
'Dim xmlRateChangeAudit As IXMLDOMNode
'Dim xmlRateChangeAuditListNode As IXMLDOMNode
''BM0230 MDC 08/01/2003 - End
'
'    'Validation
'    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
'    strBatchNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
'    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
'    Set xmlRateChangeNode = xmlGetMandatoryNode(xmlNode, "RATECHANGE")
'
'    ' PSC 25/02/02 SYS4097
'    strQueueName = xmlGetMandatoryAttributeText(xmlNode, "QUEUENAME")
'
'    'Instantiate Objects
'    Set objBatchSchedBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
'    Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
'    Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
'    Set objMortgageProductBO = gobjContext.CreateInstance("omMP.MortgageProductBO")
'    Set xmlDoc = New FreeThreadedDOMDocument40
'
'    'Create Batch Audit --------------------------------------------------------------------
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "CreateBatchAudit"
'    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
'    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
'    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
'    xmlSetAttributeValue xmlNode, "NUMBEROFRETRIES", "1"
'    xmlTempRequest.appendChild xmlNode
'    strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'
'    'Get BatchAuditGuid of newly created record
'    xmlDoc.loadXML strResponse
'    Set xmlNode = xmlDoc.selectSingleNode(".//BATCHAUDIT")
'    strBatchAuditGUID = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")
'
'    'For Each Loan Component related to this Mortgage Sub Quote ----------------------------
'    Set xmlNodeList = vxmlRequestNode.selectNodes(".//LOANCOMPONENT")
'    For Each xmlLoanCompNode In xmlNodeList
'
'        'Create RateChangeBatchAudit record
'        Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
'        Set xmlNode = xmlDoc.createElement("RATECHANGEBATCHAUDIT")
'        xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
'        xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
'        xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGUID
'        xmlCopyMandatoryAttribute xmlRateChangeNode, xmlNode, "APPLICATIONNUMBER"
'        xmlCopyMandatoryAttribute xmlRateChangeNode, xmlNode, "APPLICATIONFACTFINDNUMBER"
'        xmlCopyMandatoryAttribute xmlRateChangeNode, xmlNode, "QUOTATIONNUMBER"
'        xmlCopyMandatoryAttribute xmlRateChangeNode, xmlNode, "MORTGAGESUBQUOTENUMBER"
'        xmlCopyMandatoryAttribute xmlLoanCompNode, xmlNode, "LOANCOMPONENTSEQUENCENUMBER"
'        xmlCopyMandatoryAttribute xmlRateChangeNode, xmlNode, "APPLICATIONDATE"
'' AQR 4176 - need to call omMP.MortgageProductBO with element - based request
'        Set xmlMPRequestElem = vxmlRequestNode.cloneNode(False)
'        Set xmlTempRequest = xmlMPRequestElem.appendChild(xmlLoanCompNode).parentNode
'        Set xmlMPRequestElem = xmlCreateElementRequestFromNode(xmlTempRequest, "LOANCOMPONENT", True).firstChild
'        'SYS4340 Need to pass in string not node
'        'set xmlMPResponse = objMortgageProductBO.GetInterestRateForRateChange(xmlMPRequestElem)
'
'        'BM0230 MDC 08/01/2003 - If required, get old interest from RateChangeBatchAudit
'        If xmlGetAttributeAsBoolean(vxmlRequestNode, "BASERATEAPPLIED") = True Then
'            Set xmlRateChangeAuditListNode = xmlDoc.createElement("RATECHANGEBATCHAUDITLIST")
'            Set xmlTempRequest = xmlNode.cloneNode(False)
'            xmlCopyAttributeValue xmlRateChangeNode, xmlTempRequest, "FAILEDBATCHNUMBER", "BATCHRUNNUMBER"
'            xmlTempRequest.Attributes.removeNamedItem "BATCHAUDITGUID"
'            xmlTempRequest.Attributes.removeNamedItem "APPLICATIONDATE"
'            objRateChangeDO.FindRateChangeAuditList xmlTempRequest, xmlRateChangeAuditListNode
'            Set xmlRateChangeAudit = xmlGetMandatoryNode(xmlRateChangeAuditListNode, "RATECHANGEBATCHAUDIT")
'            xmlCopyMandatoryAttribute xmlRateChangeAudit, xmlMPRequestElem, "OLDRATE"
'        End If
'        'BM0230 MDC 08/01/2003 - End
'
'        strResponse = objMortgageProductBO.GetInterestRateForRateChange(xmlMPRequestElem.xml)
'        errCheckXMLResponse strResponse, True
'
'        Set xmlMPResponse = New FreeThreadedDOMDocument40
'        xmlMPResponse.loadXML strResponse
'
'        Set xmlMPResponseNode = xmlMPResponse.firstChild
'
'' now  convert the elements returned back into attributes
'        'Set xmlMPResponse = xmlCreateAttributeBasedResponse(xmlMPResponse, True)
'        Set xmlMPResponseNode = xmlCreateAttributeBasedResponse(xmlMPResponseNode, True)
'        'SYS4340 end
'        xmlCopyMandatoryAttribute xmlMPResponseNode, xmlNode, "NEWRATE"
'' SYS4343 Attribute name different to  database field name
'        xmlCopyMandatoryAttributeValue xmlMPResponseNode, xmlNode, "NEWRATESTARTDATE", "NEWSTARTDATE"
'        xmlCopyMandatoryAttribute xmlMPResponseNode, xmlNode, "OLDRATE"
'        objRateChangeDO.CreateRateChangeAudit xmlNode, xmlTempResponse
'    Next
'
'    'SYS3018 MDC 06/12/2001
'    gobjContext.SetComplete
'    'SYS3018 End
'
'ProcessApplicationExit:
'    Set objBatchSchedBO = Nothing
'    Set xmlDoc = Nothing
'    Set xmlTempRequest = Nothing
'    Set xmlTempResponse = Nothing
'    Set xmlNode = Nothing
'    Set xmlMessage = Nothing
'    Set xmlNodeList = Nothing
'    Set xmlLoanCompNode = Nothing
'    Set xmlRateChangeNode = Nothing
'    'BM0230 MDC 08/01/2003
'    Set xmlRateChangeAudit = Nothing
'    Set xmlRateChangeAuditListNode = Nothing
'    'BM0230 MDC 08/01/2003 - End
'
'    If Err.Number <> 0 Then
'        Err.Raise Err.Number, Err.Source, Err.Description
'    End If
'    Exit Sub
'
'ProcessApplicationVbErr:
'    'SYS3541 MDC 20/12/2001
'    strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
'    strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
'    strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
'    strErrMsg = strErrMsg & vbCrLf & "Batch Audit Guid: " & strBatchAuditGUID
'    strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
'    strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
'    strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
'    App.LogEvent strErrMsg, vbLogEventTypeError
'    'SYS3541 End
'
'    gobjContext.SetAbort
'    GoTo ProcessApplicationExit

End Sub

Private Sub IRateChangeNTTxBO_RunApplicationRateChange(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    RunApplicationRateChange
' Author:       Mark Coleman
' Date:         05/09/2001
' Purpose:      Recosts quotation for a single application based on new base rate
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
' Notes:        This procedure is run once for each item related to the current
'               batch in the BatchAudit table.
'               Runs in a new transaction so it is isolated from the rest of
'               the batch. This allows the batch to continue if this item fails.
'------------------------------------------------------------------------

On Error GoTo RunApplicationRateChangeVbErr

Const cstrFunctionName As String = "IRateChangeNTTxBO_RunApplicationRateChange"

Dim objRateChangeDO As IRateChangeDO
Dim objAppQuoteBO As ApplicationQuoteBO

Dim xmlRateChange As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlAppQuoteNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTempNode As IXMLDOMNode
Dim xmlLCNode As IXMLDOMNode
Dim xmlRateChangeNode As IXMLDOMElement

Dim strAppNo As String
Dim strResponse As String
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strErrMsg As String 'SYS3541 MDC 20/09/2001
       
    'Get Request values ----------------------------------------------------------------------
    Set xmlRateChange = xmlGetMandatoryNode(vxmlRequestNode, "RATECHANGE") '.cloneNode(True)
    strBatchNumber = xmlGetMandatoryAttributeText(xmlRateChange, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlRateChange, "BATCHRUNNUMBER")
    
    'Check if batch has been cancelled -------------------------------------------------------
    If Not IsBatchCancelled(strBatchNumber, strBatchRunNumber) Then
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
        
        'Get Rate Change Batch Audit List ----------------------------------------------------
        Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        Set xmlRateChangeNode = xmlDoc.createElement("RATECHANGE")
        xmlTempRequest.appendChild xmlRateChangeNode
        xmlRateChangeNode.setAttribute "BATCHNUMBER", strBatchNumber
        xmlRateChangeNode.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
        'SYS3440 MDC 07/12/2001 - Only find loan components for this application
        xmlCopyMandatoryAttribute xmlRateChange, xmlRateChangeNode, "BATCHAUDITGUID"
        'SYS3440 End
        objRateChangeDO.FindRateChangeAuditList xmlRateChangeNode, xmlTempResponse
        
        'Get the application number
        Set xmlNodeList = xmlTempResponse.selectNodes("RATECHANGEBATCHAUDIT")
        If xmlNodeList Is Nothing Then
            errThrowError cstrFunctionName, oeMissingElement, "RATECHANGEBATCHAUDIT"
        End If
        Set xmlNode = xmlNodeList.Item(0)
        strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
        'Check if the application is locked --------------------------------------------------
        If GetApplicationLockStatus(strAppNo) = lsLocked Then
            'Lock found, so post message back to queue and retry later
            If DoRetry(vxmlRequestNode) Then
                'Indicate that retrying so batch schedule success/failure not updated at this time
                xmlSetAttributeValue vxmlResponseNode, "TYPE", "RETRY"
            End If
        Else
            'Recost the quotation ------------------------------------------------------------
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            Set xmlAppQuoteNode = xmlDoc.createElement("APPLICATIONQUOTE")
            xmlTempRequest.appendChild xmlAppQuoteNode
            
            'get key values
            Set xmlTempNode = xmlDoc.createElement("APPLICATIONNUMBER")
            xmlTempNode.Text = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
            xmlAppQuoteNode.appendChild xmlTempNode
             
            Set xmlTempNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
            xmlTempNode.Text = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
            xmlAppQuoteNode.appendChild xmlTempNode
             
            Set xmlTempNode = xmlDoc.createElement("QUOTATIONNUMBER")
            xmlTempNode.Text = xmlGetMandatoryAttributeText(xmlNode, "QUOTATIONNUMBER")
            xmlAppQuoteNode.appendChild xmlTempNode
             
            Set xmlTempNode = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
            xmlTempNode.Text = xmlGetMandatoryAttributeText(xmlNode, "MORTGAGESUBQUOTENUMBER")
            xmlAppQuoteNode.appendChild xmlTempNode
            
            Set xmlLCNode = xmlDoc.createElement("LOANCOMPONENTS")
            xmlAppQuoteNode.appendChild xmlLCNode
            For Each xmlNode In xmlNodeList
                'compile loancomponent list
                Set xmlTempNode = xmlDoc.createElement("LOANCOMPONENTSEQUENCENUMBER")
                xmlTempNode.Text = xmlGetMandatoryAttributeText(xmlNode, "LOANCOMPONENTSEQUENCENUMBER")
                xmlLCNode.appendChild xmlTempNode
            Next
        
            'BM0230 MDC 08/01/2003
            xmlCopyAttribute vxmlRequestNode, xmlTempRequest, "BASERATEAPPLIED"
            'BM0230 MDC 08/01/2003 - End
            
            Set objAppQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
            
            strResponse = objAppQuoteBO.RecostQuotationForRateChange(xmlTempRequest.xml)
            errCheckXMLResponse strResponse, True
            
            '#TASK - Printing
            
        End If
    Else  'JD BMIDS604
        'Batch is cancelled, raise an error to prevent the batch status being set as Complete
        errThrowError cstrFunctionName, 7025
    End If
    
    'SYS3018 MDC 06/12/2001
    gobjContext.SetComplete
    'SYS3018 End
    
RunApplicationRateChange_Exit:
    Set objRateChangeDO = Nothing
    Set objAppQuoteBO = Nothing
    Set xmlRateChange = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlDoc = Nothing
    Set xmlAppQuoteNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlTempNode = Nothing
    Set xmlLCNode = Nothing
    Set xmlRateChangeNode = Nothing
   
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source & ", " & TypeName(Me) & "." & cstrFunctionName, Err.Description
    End If
    Exit Sub

RunApplicationRateChangeVbErr:
    'SYS3541 MDC 20/12/2001
    strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & cstrFunctionName
    strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
    strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
    strErrMsg = strErrMsg & vbCrLf & "Application Number: " & strAppNo
    strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
    strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
    strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
    App.LogEvent strErrMsg, vbLogEventTypeError
    'SYS3541 End
    
    gobjContext.SetAbort
    GoTo RunApplicationRateChange_Exit

End Sub

Private Function IsBatchCancelled(ByVal strBatchNumber As String, _
                                    ByVal strBatchRunNumber As String) As Boolean

On Error GoTo IsBatchCancelled_Exit

Const cstrFunctionName As String = "IsBatchCancelled"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
Dim xmlBatchScheduleNode As IXMLDOMElement

Dim objBatchSched As BatchScheduleBO

Dim strValue As String
Dim strValueId As String
Dim strResponse As String
   
    'Get the current status of the batch schedule
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlTempRequest.setAttribute "OPERATION", "GetBatchSchedule"
    Set xmlBatchScheduleNode = xmlDoc.createElement("BATCHSCHEDULE")
    xmlBatchScheduleNode.setAttribute "BATCHNUMBER", strBatchNumber
    xmlBatchScheduleNode.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    xmlTempRequest.appendChild xmlBatchScheduleNode
    
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse
    Set xmlBatchScheduleNode = xmlGetMandatoryNode(xmlDoc.documentElement, "BATCHSCHEDULE")
    strValue = xmlGetMandatoryAttributeText(xmlBatchScheduleNode, "STATUS")
    
    'Get the combovalue for Cancelled
    strValueId = GetFirstComboValueId("BatchScheduleStatus", "CA")
    
    IsBatchCancelled = (strValue = strValueId)
'    If strValue = strValueId Then
'        'Batch Cancelled
'        IsBatchCancelled = True
'    Else
'        IsBatchCancelled = False
'    End If

IsBatchCancelled_Exit:
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set objBatchSched = Nothing
    Set xmlBatchScheduleNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Sub IRateChangeNTTxBO_SetBaseRateAppliedDate(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    IRateChangeNTTxBO_SetBaseRateAppliedDate
' Author:       Mark Coleman
' Date:         20/08/2001
' Purpose:      Update all base rates included in this batch with an applied date
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to append to
' Notes:
'               Update base records where
'               StartDate < System Date + Rate Change Look Ahead Days
'               and Applied date is null
'------------------------------------------------------------------------

On Error GoTo IRateChangeNTTxBO_SetBaseRateAppliedDateErr   'SYS3018 MDC 06/12/2001

'Const cstrFunctionName As String = "IRateChangeNTTxBO_SetBaseRateAppliedDate"

Dim objRateChangeDO As IRateChangeDO

    Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
    objRateChangeDO.SetBaseRateAppliedDate vxmlRequestNode, vxmlResponseNode

    'SYS3018 MDC 06/12/2001
    gobjContext.SetComplete
    'SYS3018 End
    
IRateChangeNTTxBO_SetBaseRateAppliedDateExit:
    Set objRateChangeDO = Nothing
    
    'SYS3018 MDC 06/12/2001
'    errCheckError cstrFunctionName, TypeName(Me)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    Exit Sub

IRateChangeNTTxBO_SetBaseRateAppliedDateErr:
    gobjContext.SetAbort
    GoTo IRateChangeNTTxBO_SetBaseRateAppliedDateExit
    'SYS3018 End

End Sub

Private Function GetApplicationLockStatus(ByVal strAppNo As String) As LOCKSTATUS

On Error GoTo GetApplicationLockStatus_Exit

Const cstrFunctionName As String = "GetApplicationLockStatus"

Dim objAppManBO As ApplicationManagerBO

Dim xmlTempRequest As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim lngErr As Long
Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Create Request
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONLOCK")
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = strAppNo
    xmlAppNode.appendChild xmlNode
    xmlTempRequest.appendChild xmlAppNode
    
    'Check if Application is locked
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strResponse = objAppManBO.GetLockData(xmlTempRequest.xml)
    
    'Check Response
    lngErr = errCheckXMLResponse(strResponse, False)
    If lngErr <> 0 Then
        lngErr = errGetOmigaErrorNumber(lngErr)
    End If
    If lngErr = oeRecordNotFound Then
        'No lock exists
        GetApplicationLockStatus = lsUnlocked
    ElseIf lngErr = 0 Then
        'Lock exists
        GetApplicationLockStatus = lsLocked
    Else
        'Unexpected error
        GetApplicationLockStatus = lsError
    End If
    
GetApplicationLockStatus_Exit:
    Set objAppManBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlAppNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function DoRetry(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean

On Error GoTo DoRetryExit

Const cstrFunctionName As String = "DoRetry"

Dim objBatchSched As BatchScheduleBO
Dim objMessageQueue As IomMessageQueue

Dim xmlBatch As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim lngValue As Long
Dim lngRetries As Long
Dim strResponse As String
Dim strQueueName As String
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strBatchAuditGUID As String

    'Get Request values
    Set xmlBatch = xmlGetMandatoryNode(vxmlRequestNode, "RATECHANGE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatch, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatch, "BATCHRUNNUMBER")
    strBatchAuditGUID = xmlGetMandatoryAttributeText(xmlBatch, "BATCHAUDITGUID")
  
    'Find maximum number of retries
    lngValue = GetGlobalParamAmount("RateChangeApplicationRetries")
  
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
  
    'Find current number of retries
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindBatchAuditList"
    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGUID
    xmlTempRequest.appendChild xmlNode
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse
    Set xmlNode = xmlGetMandatoryNode(xmlDoc.documentElement, "BATCHAUDIT")
    lngRetries = xmlGetMandatoryAttributeAsLong(xmlNode, "NUMBEROFRETRIES")
    
    If lngRetries < lngValue Then
        'Increment the number of retries count...
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "UpdateBatchAudit"
        Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
        xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
        xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
        xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGUID
        xmlSetAttributeValue xmlNode, "NUMBEROFRETRIES", CStr(lngRetries + 1)
        xmlTempRequest.appendChild xmlNode
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
        
        '... and retry by posting message back to the queue
        'Get Queue Name & setup Message Queue Request
        strQueueName = GetGlobalParamString("RateChangeQueueName")
        If Len(Trim$(strQueueName)) = 0 Then
            errThrowError cstrFunctionName, oeInvalidParameter, "RateChangeQueueName not found"
        End If
        'vxmlRequestNode.removeChild xmlBatchAudit   'Remove the Batch Audit node from the original request
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "SendToQueue"
        Set xmlNode = xmlDoc.createElement("MESSAGEQUEUE")
        xmlSetAttributeValue xmlNode, "QUEUENAME", strQueueName
        xmlSetAttributeValue xmlNode, "PROGID", gstrRATECHANGE & ".RateChangeBO"
        xmlSetAttributeValue xmlNode, "XML", vxmlRequestNode.xml
        xmlTempRequest.appendChild xmlNode
        'Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
        Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
        Set xmlTempResponse = objMessageQueue.SendToQueue(xmlTempRequest)
        DoRetry = True
        
    Else
        'Max retries exceeded so fail
        'BMIDS622 WriteRejectReport does nothing, so there is no point calling it
        'Call WriteRejectReport
        errThrowError cstrFunctionName, oeRCBatchLockingError
    End If

DoRetryExit:
    Set objBatchSched = Nothing
    Set objMessageQueue = Nothing
    Set xmlBatch = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

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

