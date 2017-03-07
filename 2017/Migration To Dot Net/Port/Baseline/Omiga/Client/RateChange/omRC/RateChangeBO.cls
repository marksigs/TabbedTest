VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RateChangeBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------------------------------
'Workfile:      RateChangeBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Rate Change Business Object.
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV/MC  20/08/2001  Created
'MC     19/11/2001  SYS3027 - Improve error handling
'MC     14/12/2001  SYS3505 - Add audit reporting
'MC     20/12/2001  SYS3541 - Write error info to event log
'PSC    25/02/2002  SYS4097 - Amend RunRateChangeBatch to get QueueName. Improve
'                             error handling
'GHun   14/03/2002  SYS4273 - Transactioning incorrect
'DRC    15/03/2002  SYS4214 - Error Message in RunRateChangeBatch
'DRC    15/03/2002  SYS4176 - Call to GetInterestRateForRateChange
'SA     25/03/2002  SYS4287 - Call to omBatch.PrintDocuments added.
'SA     28/03/2002  SYS4340 - Change request to call completebatch to include standard omiga request.
'GHun   04/04/2002  SYS4333 - Locked cases error detail in batch rate change
'GD     05/04/2002  SYS4334 - 'BATCHAUDIT field sometimes NULL' - add call to BatchScheduleBO.IncrementBatchScheduleFailure
'SA     12/04/2002  SYS4379 - Call to Print Docs - make sure request is correct.
'GHun   16/04/2002  SYS4368 - Errors in failed batch run, rerun
'CL     13/05/2002  SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'-------------------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements MessageQueueComponentVC2

Private gobjContext As ObjectContext
Private gobjDO As IRateChangeDO


'-------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    19/09/2002  BMIDS00501 - Improve error handling in CompleteBatch
'MDC    07/01/2003  BM0230 - Fix interest rate selection in failed batches
'GHun   19/02/2003  BM0367 - Amend RunRateChangeBatch to pass Delay and ServerTime to ProcessApplication
'GHun   02/04/2003  BM0438 - Amend WasBaseRateApplied to use GetBatchScheduleProgress
'GHun   12/08/2003  BMIDS622 - Restructure
'GHun   05/11/2003  BMIDS622 - Amend CompleteBatch to improve missing log folder handling
'GHun   12/11/2003  BMIDS622 - Amend CompleteBatch to make the reporting optional
'JD     22/11/2004  BMIDS604 - Amend CompleteBatch to include Cancelled status.
'-------------------------------------------------------------------------------------------------------

Private Sub FindApplications(ByVal vxmlRequestNode As IXMLDOMElement, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    FindApplications
' Author:       Mark Coleman
' Date:         05/09/2001
' Purpose:      Find all applications that are affected by a Rate Change
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------
        
'BMIDS622 No longer used, but cannot be deleted because OOSS will just replace it
'On Error GoTo FindApplicationsExit
'
'Dim xmlNode As IXMLDOMNode
'Dim xmlNodeList As IXMLDOMNodeList
'
'Const strFunctionName As String = "FindApplications"
'
'    If gobjDO Is Nothing Then
'        Set gobjDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
'    End If
'
'    'Find all applications for which a Rate Change is required
'    gobjDO.FindApplications vxmlRequestNode, vxmlResponseNode
'
'    'For Each application find the Loan Components
'    Set xmlNodeList = vxmlResponseNode.selectNodes("RATECHANGE")
'    For Each xmlNode In xmlNodeList
'        gobjDO.FindLoanComponentsForApplication xmlNode, xmlNode
'        gobjDO.FindBaseRateDataforLoanComponent xmlNode, xmlNode
'    Next
'
'FindApplicationsExit:
'    Set xmlNode = Nothing
'    Set xmlNodeList = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)
'
End Sub

Private Sub FindFailedApplications(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
''------------------------------------------------------------------------
'' Procedure:    FindFailedApplications
'' Author:       Mark Coleman
'' Date:         05/09/2001
'' Purpose:      Find all failed applications from a previously run
''               Rate Change batch job.
'' Input parameters:
''               vxmlRequestNode - XML Request
''               vxmlResponseNode - XML Response to be appended to
'' Output parameters:
''               N/A
''------------------------------------------------------------------------
''BM0230 MDC 09/01/2003
''On Error GoTo FindFailedApplications_Exit
''
''Dim strFunctionName As String
''strFunctionName = "FindFailedApplications"
''
''Dim objBatchSched As BatchScheduleBO
''Dim objRateChangeDO As IRateChangeDO
''
''Dim xmlTempRequest As IXMLDOMNode
''Dim xmlTempResponse As IXMLDOMNode
''Dim xmlRateChangeBatchAuditList As IXMLDOMNodeList
''Dim xmlApplicationList As IXMLDOMNodeList
''Dim xmlAppNode As IXMLDOMNode
''Dim xmlRateChangeNode As IXMLDOMNode
''Dim xmlApplicationHeaderElem As IXMLDOMElement
''Dim xmlLoanComponentElem As IXMLDOMElement
''Dim xmlNode As IXMLDOMNode
''Dim xmlTempResponseNode As IXMLDOMNode
''Dim xmlNodeList As IXMLDOMNodeList
''Dim xmlDoc As FreeThreadedDOMDocument40
''
''Dim strResponse As String
''Dim strValue As String
''Dim strValueId As String
''
''Dim strHeaderAppNumber As String
''Dim strHeaderAppFFNumber As String
''Dim strCurrentApplicationNumber As String
''Dim strCurrentApplicationFactFindNumber As String
''
''    Set xmlDoc = New FreeThreadedDOMDocument40
''
''    'Validate
''    'PSC 25/02/02 SYS4097
''    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
''    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
''    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
''
''    'Ensure we are looking for failed rate changes
''    strValueId = GetFirstComboValueId("BatchAuditStatus", "F")
''    xmlSetAttributeValue xmlNode, "STATUS", strValueId
''
''    'Create request
''    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
''    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindBatchAuditList"
''
''    'Call FindBatchAuditList to find failed items
''    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
''    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
''    xmlDoc.loadXML strResponse
''    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
''
''
''    'Get Batch Audit records found and retrieve corresponding RateChangeAudit records
''    Set xmlNodeList = xmlDoc.selectNodes(".//BATCHAUDIT")
''    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
''    Set xmlTempRequest = xmlDoc.createElement("RATECHANGEBATCHAUDIT")
''    Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
''
''    For Each xmlNode In xmlNodeList
''
''        xmlCopyMandatoryAttribute xmlNode, xmlTempRequest, "BATCHAUDITGUID"
''        objRateChangeDO.FindRateChangeAuditList xmlTempRequest, xmlTempResponse
''
''        ' need to convert RateChange Audit record into application & node hierarchy
''        Set xmlRateChangeBatchAuditList = xmlTempResponse.selectNodes(".//RATECHANGEBATCHAUDIT")
''        strHeaderAppNumber = ""
''        strHeaderAppFFNumber = ""
''        For Each xmlRateChangeNode In xmlRateChangeBatchAuditList
''            strCurrentApplicationNumber = xmlAssistEx.xmlGetAttributeText(xmlRateChangeNode, "APPLICATIONNUMBER")
''            strCurrentApplicationFactFindNumber = xmlAssistEx.xmlGetAttributeText(xmlRateChangeNode, "APPLICATIONFACTFINDNUMBER")
''            If (strCurrentApplicationNumber <> strHeaderAppNumber) Or _
''               (strCurrentApplicationFactFindNumber <> strHeaderAppFFNumber) Then
''               ' Create a new Header ELEMENT
''                strHeaderAppNumber = strCurrentApplicationNumber
''                strHeaderAppFFNumber = strCurrentApplicationFactFindNumber
''                Set xmlApplicationHeaderElem = xmlDoc.createElement("RATECHANGE")
''                xmlCopyAttribute xmlRateChangeNode, xmlApplicationHeaderElem, "APPLICATIONNUMBER"
''                xmlCopyAttribute xmlRateChangeNode, xmlApplicationHeaderElem, "APPLICATIONFACTFINDNUMBER"
''                xmlCopyAttribute xmlRateChangeNode, xmlApplicationHeaderElem, "QUOTATIONNUMBER"
''                xmlCopyAttribute xmlRateChangeNode, xmlApplicationHeaderElem, "APPLICATIONDATE"
''                xmlCopyAttribute xmlRateChangeNode, xmlApplicationHeaderElem, "MORTGAGESUBQUOTENUMBER"
''                vxmlResponseNode.appendChild xmlApplicationHeaderElem
''            End If
''            'Create a new LOANCOMPONENT Element
''            Set xmlLoanComponentElem = xmlDoc.createElement("LOANCOMPONENT")
''            xmlCopyAttribute xmlRateChangeNode, xmlLoanComponentElem, "LOANCOMPONENTSEQUENCENUMBER"
''            xmlApplicationHeaderElem.appendChild xmlLoanComponentElem
''        Next
''
''        'Now need to get any altered (baserate) data
''        'For Each application find the Loan Components
''        Set xmlApplicationList = vxmlResponseNode.selectNodes("RATECHANGE")
''        For Each xmlAppNode In xmlApplicationList
''           objRateChangeDO.FindBaseRateDataforLoanComponent xmlAppNode, xmlAppNode
''        Next
''
''    Next
''
''FindFailedApplications_Exit:
''    Set objBatchSched = Nothing
''    Set objRateChangeDO = Nothing
''    Set xmlTempRequest = Nothing
''    Set xmlNode = Nothing
''    Set xmlNodeList = Nothing
''    Set xmlDoc = Nothing
''    Set xmlNodeList = Nothing
''    Set xmlTempResponse = Nothing
''    Set objRateChangeDO = Nothing
''    Set xmlRateChangeBatchAuditList = Nothing
''    Set xmlApplicationHeaderElem = Nothing
''    Set xmlLoanComponentElem = Nothing
''    Set xmlApplicationList = Nothing
''    Set xmlAppNode = Nothing
''
''    errCheckError strFunctionName, TypeName(Me)
    
'BMIDS622 No longer used, but cannot be deleted because OOSS will just replace it
'On Error GoTo FindFailedApplications_Exit
'
'Dim strFunctionName As String
'strFunctionName = "FindFailedApplications"
'
'Dim xmlNode As IXMLDOMNode
'Dim xmlApplicationList As IXMLDOMNodeList
'Dim xmlAppNode As IXMLDOMNode
'
'    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
'    xmlCheckMandatoryAttribute xmlNode, "BATCHNUMBER"
'    xmlCheckMandatoryAttribute xmlNode, "BATCHRUNNUMBER"
'
'    If gobjDO Is Nothing Then
'        Set gobjDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
'    End If
'
'    'Find all failed applications for which a Rate Change is required
'    gobjDO.FindFailedApplications xmlNode, vxmlResponseNode
'
'    Set xmlApplicationList = vxmlResponseNode.selectNodes(".//RATECHANGE")
'    If xmlApplicationList.length = 0 Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'
'    'Now need to get any altered (baserate) data
'    'For Each application find the Loan Components
'    For Each xmlAppNode In xmlApplicationList
'       gobjDO.FindBaseRateDataforLoanComponent xmlAppNode, xmlAppNode
'    Next
'
'FindFailedApplications_Exit:
'    Set xmlNode = Nothing
'    Set xmlApplicationList = Nothing
'    Set xmlAppNode = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)
''BM0230 MDC 09/01/2003 - End

End Sub

Private Sub FindRateChangeAuditList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlReponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    FindRateChangeAuditList
' Author:       Mark Coleman
' Date:         05/09/2001
' Purpose:      Find RateChangeBatchAudit records
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------
    
On Error GoTo FindRateChangeAuditListExit
    
Const strFunctionName As String = "FindRateChangeAuditList"
Dim xmlBatchNode As IXMLDOMNode
    
    'Validate RATECHANGE element exists
    Set xmlBatchNode = xmlGetMandatoryNode(vxmlRequestNode, "RATECHANGE")
    
    'Validate BATCHAUDITGUID attribute is specified
    xmlCheckMandatoryAttribute xmlBatchNode, "BATCHAUDITGUID"
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
    End If
    
    gobjDO.FindRateChangeAuditList vxmlRequestNode, vxmlReponseNode
        
FindRateChangeAuditListExit:
    Set xmlBatchNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub RunApplicationRateChange(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
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
'------------------------------------------------------------------------

On Error GoTo RunApplicationRateChangeExit
    
Const cstrFunctionName As String = "RunApplicationRateChange"
    
Dim xmlRateChangeNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlFinishTrans As IXMLDOMNode
'Dim xmlTempResponse As IXMLDOMNode
Dim xmlBatchSched As IXMLDOMElement

Dim objRateChangeNTTxBO As IRateChangeNTTxBO
Dim objBatchSched As BatchScheduleBO

Dim strResponse As String
Dim blnInitialising As Boolean

'SYS3541 MDC 20/12/2001
Dim strErrMsg As String
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strBatchAuditGUID As String
'SYS3541 End
       
    blnInitialising = True
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    'Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    
    'Check for the Mandatory node
    Set xmlRateChangeNode = xmlGetMandatoryNode(vxmlRequestNode, "RATECHANGE")
    
    'SYS3541 MDC 20/12/2001
    strBatchNumber = xmlGetMandatoryAttributeText(xmlRateChangeNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlRateChangeNode, "BATCHRUNNUMBER")
    strBatchAuditGUID = xmlGetMandatoryAttributeText(xmlRateChangeNode, "BATCHAUDITGUID")
    'SYS3541 End
    
    Set xmlBatchSched = xmlDoc.createElement("BATCHSCHEDULE")
    xmlBatchSched.setAttribute "BATCHNUMBER", strBatchNumber
    xmlBatchSched.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    xmlBatchSched.setAttribute "BATCHAUDITGUID", strBatchAuditGUID
    blnInitialising = False
    
    Set objRateChangeNTTxBO = gobjContext.CreateInstance(App.Title & ".RateChangeNTTxBO")
    objRateChangeNTTxBO.RunApplicationRateChange vxmlRequestNode, vxmlResponseNode
    
    'SYS4333 GHun 04/04/2002 Do not increment batch success if response is RETRY
    If Not (xmlGetAttributeText(vxmlResponseNode, "TYPE") = "RETRY") Then
    
        'Increment Batch Success
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlTempRequest.setAttribute "OPERATION", "IncrementBatchScheduleSuccess"
        xmlTempRequest.appendChild xmlBatchSched
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        If errCheckXMLResponse(strResponse, False) <> 0 Then
            'Write Reject Report ???
            'BMIDS622 WriteRejectReport does nothing so there is no point in calling it
            'Call WriteRejectReport
        Else
            xmlDoc.loadXML strResponse
            Set xmlFinishTrans = xmlDoc.selectSingleNode(".//BATCHSCHEDULE")
        End If
        
    End If 'SYS4333
    
    'Check if final transaction
    'PSC 25/02/02 SYS4097 - Move to CompleteBatch
'    If Not xmlFinishTrans Is Nothing Then
'        If xmlGetAttributeAsBoolean(xmlFinishTrans, "LASTTRANSACTION", "0") Then
'            objRateChangeNTTxBO.SetBaseRateAppliedDate xmlTempRequest, xmlTempResponse
'        End If
'    End If
    
RunApplicationRateChangeExit:
    
    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & cstrFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Audit Guid: " & strBatchAuditGUID
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    'SYS3541 End
    
    'If error is DB failure or occurs during initialisation we cannot update the batch status
    'If Err.Number <> 0 And (Err = DB_FAIL Or blnInitialising) Then
    If Err.Number <> 0 Then
        If blnInitialising Then
        '   Raise the error back so that the message queue is informed and message moved to holding queue
           Err.Raise Err.Number, Err.Source, Err.Description
        Else
            'Increment Batch Schedule Failure
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            xmlTempRequest.setAttribute "OPERATION", "IncrementBatchScheduleFailure"
            xmlBatchSched.setAttribute "ERRORNUMBER", CStr(Err.Number)
            xmlBatchSched.setAttribute "ERRORSOURCE", Err.Source
            xmlBatchSched.setAttribute "ERRORDESCRIPTION", Err.Description
            xmlTempRequest.appendChild xmlBatchSched
            strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
            If errCheckXMLResponse(strResponse, False) <> 0 Then
                'Write Reject Report ???
                'BMIDS622 WriteRejectReport does nothing, so there is no point in calling it
                'Call WriteRejectReport
            Else
                xmlDoc.loadXML strResponse
                Set xmlFinishTrans = xmlDoc.selectSingleNode(".//BATCHSCHEDULE")
            End If
        End If
    End If
    
    'Check if final transaction
    If Not xmlFinishTrans Is Nothing Then
        If xmlGetAttributeAsBoolean(xmlFinishTrans, "LASTTRANSACTION", "0") Then
            'SYS4379 Need to pass in whole Omiga request string
            'Call CompleteBatch(xmlBatchSched, xmlTempResponse)
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            xmlTempRequest.appendChild xmlBatchSched
            'BMIDS622 GHun 13/08/2003
            'Call CompleteBatch(xmlTempRequest, xmlTempResponse)
            QueueBatchCompletion xmlTempRequest
            'BMIDS622 End
        End If
    End If
    
    Set xmlRateChangeNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlFinishTrans = Nothing
    'Set xmlTempResponse = Nothing
    Set xmlBatchSched = Nothing
    Set xmlDoc = Nothing
    Set objRateChangeNTTxBO = Nothing
    Set objBatchSched = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub CompleteBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    CompleteBatch
' Author:       Mark Coleman
' Date:         05/09/2001
' Purpose:      Action any clean up processing on completion of a batch job
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
'------------------------------------------------------------------------

'BMIDS00501 MDC 19/09/2002
'On Error GoTo CompleteBatch_Exit
On Error GoTo CompleteBatch_VbErr
'BMIDS00501 MDC 19/09/2002 - End

    Const cstrFunctionName As String = "CompleteBatch"
    
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlBatchSched As IXMLDOMElement
    Dim xmlBatchApp As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlDoc As FreeThreadedDOMDocument40
    'SYS3505 MDC 14/12/2001
    Dim xmlNode As IXMLDOMNode
    Dim xmlRCAuditNodeList As IXMLDOMNodeList
    Dim xmlRCAuditNode As IXMLDOMNode
    Dim intFileNo As Integer
    Dim intValue As Integer
    Dim strTextLine As String
    Dim strBatchLogFile As String
    Dim strBatchLogFilePath As String
    'SYS3505 End
    
    Dim objBatchSched As BatchScheduleBO
    Dim objRateChangeDO As IRateChangeDO
    Dim objRateChangeNTTxBO As IRateChangeNTTxBO
    
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strValue As String
    Dim strResponse As String
    'BMIDS00501 MDC 19/09/2002
    Dim lngRet As Long
    Dim strBatchStatusValueid As String
    'BMIDS00501 MDC 19/09/2002 - End
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    'BMIDS622
    Dim blnWriteLog As Boolean
    'BMIDS622 End

    'BMIDS00501 MDC 19/09/2002
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    'BMIDS00501 MDC 19/09/2002 - End

    'Validation
    'SYS4340 We are now sending in the whole Omiga4 request
    'Set xmlBatchSched = vxmlRequestNode 'xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "//BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    'Find List of Applications processed
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    Set xmlBatchApp = xmlDoc.createElement("BATCH")
    xmlBatchApp.setAttribute "BATCHNUMBER", strBatchNumber
    xmlBatchApp.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
    objRateChangeDO.FindBatchApplicationList xmlBatchApp, xmlTempResponse
    
    'Validate overall success by checking if any transactions errored
    'Get the combovalue for F - 'Failure'
    strValue = GetFirstComboValueId("BatchAuditStatus", "F")
    Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATION[@STATUS='" & strValue & "' or @STATUS='']")
    
    If xmlNodeList.length > 0 Then
        strValue = "CWE"
        'JD BMIDS604 check the error number on the failed applications. We are looking for a 'Cancelled' app
        Dim nCancelledErr As Long
        Dim xmlCancelledNodeList As IXMLDOMNodeList
        nCancelledErr = 7025 + vbObjectError + 512
        Set xmlCancelledNodeList = xmlTempResponse.selectNodes("APPLICATION[@ERRORNUMBER='" & CStr(nCancelledErr) & "']")
        If xmlCancelledNodeList.length > 0 Then
            strValue = "CA"
        End If
    Else
        strValue = "C"
    End If
    
    'BMIDS00501 MDC 19/09/2002
    Set xmlTempRequest = xmlDoc.createElement("RATECHANGE")
    xmlTempRequest.setAttribute "BATCHNUMBER", strBatchNumber
    xmlTempRequest.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    
    'Get the combovalue for status
    strBatchStatusValueid = GetFirstComboValueId("BatchScheduleStatus", strValue)
    'strValue = GetFirstComboValueId("BatchScheduleStatus", strValue)
    'Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    
    'Get the combovalue for S - 'Success'
    strValue = GetFirstComboValueId("BatchAuditStatus", "S")
    If xmlTempResponse.selectNodes("APPLICATION[@STATUS='" & strValue & "']").length > 0 Then
        'Set BaseRateAppliedDate as at least one item in the Batch succeeded
        Set objRateChangeNTTxBO = gobjContext.CreateInstance(App.Title & ".RateChangeNTTxBO")
        objRateChangeNTTxBO.SetBaseRateAppliedDate xmlTempRequest, xmlTempResponse
    End If
    
'    'Update the Batch Status
''    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "SetBatchStatus"
'    xmlTempRequest.appendChild xmlBatchSched
'    xmlSetAttributeValue xmlBatchSched, "STATUS", strBatchStatusValueid
'    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True, vxmlResponseNode
'
'    'SYS3505 MDC 14/12/2001. Complete audit report for rate change
'    'Get RateChangeBatchAudit records
'    Set xmlTempRequest = xmlDoc.createElement("RATECHANGE")
'    xmlSetAttributeValue xmlTempRequest, "BATCHNUMBER", strBatchNumber
'    xmlSetAttributeValue xmlTempRequest, "BATCHRUNNUMBER", strBatchRunNumber
    
    'BMIDS00501 MDC 19/09/2002 - End
    
    'BMIDS622 GHun 12/11/2003 Check whether to create a log file or not
    blnWriteLog = GetGlobalParamBoolean("RateChangeLogFile")
    
    If blnWriteLog Then
    'BMIDS622 End
    
        ' PSC 25/02/02 SYS4097 - Start
        
        On Error Resume Next
        objRateChangeDO.FindRateChangeAuditList xmlTempRequest, xmlTempResponse
        
        lngErrNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
        
        'BMIDS00501 MDC 19/09/2002
        'On Error GoTo CompleteBatch_Exit
        On Error GoTo CompleteBatch_VbErr
        'BMIDS00501 MDC 19/09/2002 - End
        
        If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
            Err.Raise lngErrNo, strSource, strDescription
        End If
        ' PSC 25/02/02 SYS4097 - End
        
        'Get the BatchLogFilePath GlobalParameter Value from GlobalParameter DB Table
        strBatchLogFilePath = GetGlobalParamString("BatchLogFilePath")
        
        'Checking for the Directory Existence to write the outputFile
        'If does not Exist then Create One
        
        'BMIDS622 Ignore errors when testing and creating the output folder
        On Error Resume Next
        'BMIDS622 End
        
        If Len(Dir(strBatchLogFilePath, vbDirectory)) = 0 Then
            MkDir strBatchLogFilePath
            'BMIDS622 Check if the directory creation succeeded
            If Len(Dir(strBatchLogFilePath, vbDirectory)) = 0 Then
                ' The output folder does not exist and cannot be created so the log cannot be written
                App.LogEvent "The BatchLogFilePath (" & strBatchLogFilePath & ") does not exist and cannot be created.", vbLogEventTypeError
                blnWriteLog = False
            End If
            'BMIDS622 End
        End If
        
        'BMIDS622 Restore the original error hanlder
        On Error GoTo CompleteBatch_VbErr
        
        If blnWriteLog Then
        'BMIDS622 End
        
            strBatchLogFile = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber & "_RC"
            
            intFileNo = FreeFile
            Open strBatchLogFile & ".log" For Append As #intFileNo
            Print #intFileNo, vbCrLf & "Batch Audit:" & vbCrLf
            Print #intFileNo, "BatchAuditGuid" & vbTab & vbTab & vbTab & vbTab & "Application" & vbTab _
                            & "Fact Find" & vbTab & "Quotation" & vbTab & "Sub Quote" _
                            & vbTab & "LC Count" & vbTab & "Status"
            Print #intFileNo, "------------------------------------------------------" _
                                & "------------------------------------------------------------------"
            
            Set xmlNodeList = xmlTempResponse.selectNodes(".//APPLICATION")
            For Each xmlNode In xmlNodeList
                strValue = xmlGetAttributeText(xmlNode, "BATCHAUDITGUID")
                strTextLine = strValue & vbTab
                
                Set xmlRCAuditNodeList = xmlTempResponse.selectNodes(".//RATECHANGEBATCHAUDIT[@BATCHAUDITGUID='" & strValue & "']")
                If xmlRCAuditNodeList.length > 0 Then
                    Set xmlRCAuditNode = xmlRCAuditNodeList.Item(0)
                    strValue = xmlGetAttributeText(xmlRCAuditNode, "APPLICATIONNUMBER")
                    strTextLine = strTextLine & strValue & vbTab
                    strValue = xmlGetAttributeText(xmlRCAuditNode, "APPLICATIONFACTFINDNUMBER")
                    strTextLine = strTextLine & strValue & vbTab & vbTab
                    strValue = xmlGetAttributeText(xmlRCAuditNode, "QUOTATIONNUMBER")
                    strTextLine = strTextLine & strValue & vbTab & vbTab
                    strValue = xmlGetAttributeText(xmlRCAuditNode, "MORTGAGESUBQUOTENUMBER")
                    strTextLine = strTextLine & strValue & vbTab & vbTab
                    strTextLine = strTextLine & CStr(xmlRCAuditNodeList.length) & vbTab & vbTab
                End If
                
                intValue = xmlGetAttributeAsInteger(xmlNode, "STATUS")
                strValue = GetComboText("BatchAuditStatus", intValue)
                strTextLine = strTextLine & strValue
                
                Print #intFileNo, strTextLine
            Next
            
            Print #intFileNo, vbCrLf & "--- End of Report ---"
            Close #intFileNo
            'SYS3505 End
        
    'BMIDS622
        End If
    End If
    'BMIDS622 End
    
    'BMIDS00501 MDC 19/09/2002
'    ' PSC 25/02/02 SYS4097
'    Set objRateChangeNTTxBO = gobjContext.CreateInstance(App.Title & ".RateChangeNTTxBO")
'    objRateChangeNTTxBO.SetBaseRateAppliedDate xmlTempRequest, xmlTempResponse
    'BMIDS00501 MDC 19/09/2002 - End
    
    'SA SYS4287 25/3/02
    'Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlBatchApp = xmlDoc.createElement("BATCH")
    xmlBatchApp.setAttribute "BATCHNUMBER", strBatchNumber
    xmlBatchApp.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.setAttribute "OPERATION", "PrintDocuments"
    xmlTempRequest.appendChild xmlBatchApp
    
    If objBatchSched Is Nothing Then
        Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    End If
    
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
   
    'BMIDS00501 MDC 19/09/2002
'    'Update the Batch Status
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlTempRequest.setAttribute "OPERATION", "SetBatchStatus"
    xmlTempRequest.appendChild xmlBatchSched
    xmlBatchSched.setAttribute "STATUS", strBatchStatusValueid
    'BMIDS622 Set the completion date and time
    xmlSetSysDateToNodeAttrib xmlBatchSched, "COMPLETIONDATETIME", True
    'BMIDS622 End
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        'SetBatchStatus failed so write event log message
        strValue = vbCrLf & "Error occurred in " & TypeName(Me) & ".SetBatchStatus" _
            & vbCrLf & "Number: " & CStr(Err.Number) & vbCrLf & "Source: " & Err.Source _
            & vbCrLf & "Description: " & Err.Description
        App.LogEvent strValue, vbLogEventTypeError
    End If
    'BMIDS00501 MDC 19/09/2002 - End
    
CompleteBatch_Exit:
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlBatchSched = Nothing
    Set xmlBatchApp = Nothing
    Set xmlNodeList = Nothing
    Set xmlDoc = Nothing
    Set objBatchSched = Nothing
    Set objRateChangeDO = Nothing
    'SYS3505 MDC 14/12/2001
    Set xmlNode = Nothing
    Set xmlRCAuditNodeList = Nothing
    Set xmlRCAuditNode = Nothing
    Set objRateChangeNTTxBO = Nothing
    'SYS3505 End
       
    'BMIDS00501 MDC 19/09/2002
    'No need to raise errors back as they can only go back to MessageQueueComponentVC2_OnMessage
    'errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CompleteBatch_VbErr:
    'An error has occurred processing CompleteBatch. Update the Batch Status to
    'Complete With Errors and write error info to BatchSchedule
    strValue = vbCrLf & "Error occurred in " & TypeName(Me) & "." & cstrFunctionName _
        & vbCrLf & "Number: " & CStr(Err.Number) & vbCrLf & "Source: " & Err.Source _
        & vbCrLf & "Description: " & Err.Description
    App.LogEvent strValue, vbLogEventTypeError
    
    'Get the combovalue for status Completed With Errors
    strValue = GetFirstComboValueId("BatchScheduleStatus", "CWE")
    
    'Update the Batch Status
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlTempRequest.setAttribute "OPERATION", "SetBatchStatus"
    xmlTempRequest.appendChild xmlBatchSched
    xmlBatchSched.setAttribute "STATUS", strValue
    xmlBatchSched.setAttribute "ERRORNUMBER", Err.Number
    xmlBatchSched.setAttribute "ERRORSOURCE", cstrFunctionName & ", " & Err.Source
    xmlBatchSched.setAttribute "ERRORDESCRIPTION", Err.Description
    
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, False)
    If lngRet <> 0 Then
        'SetBatchStatus failed so write event log message
        strValue = vbCrLf & "Error occurred in " & TypeName(Me) & ".SetBatchStatus" _
            & vbCrLf & "Number: " & CStr(Err.Number) & vbCrLf & "Source: " & Err.Source _
            & vbCrLf & "Description: " & Err.Description
        App.LogEvent strValue, vbLogEventTypeError
    End If
    
    GoTo CompleteBatch_Exit
    'BMIDS00501 MDC 19/09/2002 - End

End Sub

'BMIDS622 Replaced by updated version below
'Private Sub RunRateChangeBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
''------------------------------------------------------------------------
'' Procedure:    RunRateChangeBatch
'' Author:       Mark Coleman
'' Date:         05/09/2001
'' Purpose:      Initiate a Rate Change batch job
'' Input parameters:
''               vxmlRequestNode - XML Request
''               vxmlResponseNode - XML Response to be appended to
'' Output parameters:
''               N/A
'' Notes:        This method creates a BatchAudit record for each
''               application affected by a Rate Change. It also creates
''               RateChangeBatchAudit records for each Loan Component the
''               application quote has. It the posts a message to the
''               Rate Change message queue for processing.
''------------------------------------------------------------------------
'
'On Error GoTo RunRateChangeBatchExit
'
'Dim objBatchSchedBO As BatchScheduleBO
'Dim objMessageQueue As IomMessageQueue
'Dim objRateChangeNTTxBO As IRateChangeNTTxBO
'
'Dim xmlDoc As FreeThreadedDOMDocument40
'Dim xmlTempRequest As IXMLDOMNode
'Dim xmlTempResponse As IXMLDOMNode
'Dim xmlNode As IXMLDOMNode
'Dim xmlNodeList As IXMLDOMNodeList
'Dim xmlBatchNode As IXMLDOMNode
'Dim xmlBatchAuditNode As IXMLDOMNode
'Dim xmlBatchScheduleNode As IXMLDOMElement
'Dim xmlErrRequest As IXMLDOMNode    'SYS3027
'
'Dim strValue As String
'Dim strValueId As String
'Dim strResponse As String
'Dim strBatchNumber As String
'Dim strBatchRunNumber As String
'Dim lngAppCount As Long
'Dim blnExit As Boolean
'Dim strErrMsg As String 'SYS3541 MDC 20/12/2001
'Dim strQueueName As String
'Dim strBaseRateApplied As String    'BM0230 MDC 07/01/2003
'
'Dim strDescription As String
'Dim strApplicationNo As String
'Dim strApplicationFFNo As String
'Dim strQuotationNo As String
'Dim strMortgageSubQNo As String
'Dim strLoanComponentSeqNo As String
''GD 05/04/2002 SYS4334 - Start
''Extra Declarations
'Dim strErrDesc As String
'Dim strErrNumber As String
'Dim xmlTempNode As IXMLDOMElement
'Dim xmlTempBatchRequest As IXMLDOMElement
'Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
'Dim xmlBatchSchedule As IXMLDOMElement
''End Extra Declarations
'Const strFunctionName As String = "RunRateChangeBatch"
'
''BM0367
'Dim lRateChangeDelay    As Long
'Dim dteDBServer         As Date
'Dim lTimeDiff           As Long
'Dim xmlMQLDelay         As IXMLDOMElement
''BM0367
'
'    Set xmlDoc = New FreeThreadedDOMDocument40
'
'    'Validation
'    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
'    strBatchNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
'    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
'
'    'Create request to get Batch data
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetBatchSchedule"
'
'    ' PSC 25/02/02 SYS4097 - Start
'    Set xmlBatchScheduleNode = xmlDoc.createElement("BATCHSCHEDULE")
'    xmlCopyMandatoryAttribute xmlNode, xmlBatchScheduleNode, "BATCHNUMBER"
'    xmlCopyMandatoryAttribute xmlNode, xmlBatchScheduleNode, "BATCHRUNNUMBER"
'    ' PSC 25/02/02 SYS4097 - End
'    xmlTempRequest.appendChild xmlBatchScheduleNode
'
'
'    'Create and call the Batch Scheduler component
'    Set objBatchSchedBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
'    strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'    xmlDoc.loadXML strResponse
'    Set xmlTempResponse = xmlDoc.documentElement
'
'    ' PSC 25/02/02 SYS4097
'    Set xmlBatchScheduleNode = xmlGetMandatoryNode(xmlTempResponse, "BATCHSCHEDULE")
'
'    'Check that Batch has not been cancelled
'    ' PSC 25/02/02 SYS4097
'    strValue = xmlGetMandatoryAttributeText(xmlBatchScheduleNode, "STATUS")
'    'Get the combovalue for Cancelled
'    strValueId = GetFirstComboValueId("BatchScheduleStatus", "CA")
'    If strValue = strValueId Then
'        'Batch Cancelled
'        blnExit = True
'    End If
'
'    Dim lngErrorNo As Long
'    Dim strSource As String
'
'
'    If Not blnExit Then
'        'Find items to be processed
'        ' PSC 25/02/02 SYS4097
'        strValue = xmlGetAttributeText(xmlBatchScheduleNode, "FAILEDBATCHRUNNUMBER")
'
'
'        If Len(Trim$(strValue)) = 0 Then
'            'New Batch Job
'
'            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'            Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
'            Set xmlBatchNode = xmlDoc.createElement("BATCH")
'            xmlCopyAttribute xmlBatchScheduleNode, xmlBatchNode, "BATCHNUMBER"
'            xmlCopyAttribute xmlBatchScheduleNode, xmlBatchNode, "BATCHRUNNUMBER"
'            xmlTempRequest.appendChild xmlBatchNode
'
'            ' PSC 25/02/2002 SYS4097 - Start
'            On Error Resume Next
'            Call FindApplications(xmlTempRequest, xmlTempResponse)
'
'            lngErrorNo = Err.Number
'            strSource = Err.Source
'            strDescription = Err.Description
'            On Error GoTo RunRateChangeBatchExit
'
'            If lngErrorNo <> 0 Then
'                If errGetOmigaErrorNumber(lngErrorNo) = oeRecordNotFound Then
'                    'SYS4340 Send in Omiga4 request element too
'                    'CompleteBatch xmlBatchScheduleNode, vxmlResponseNode
'                    Dim xmlCompleteBatchRequest As IXMLDOMNode
'                    Set xmlCompleteBatchRequest = vxmlRequestNode.cloneNode(False)
'                    xmlCompleteBatchRequest.appendChild xmlBatchScheduleNode
'                    CompleteBatch xmlCompleteBatchRequest, vxmlResponseNode
'
'                    blnExit = True
'                Else
'                    Err.Raise lngErrorNo, strSource, strDescription
'                End If
'            End If
'            ' PSC 25/02/2002 SYS4097 - End
'        Else
'            'Re run Failed Batch
'
'            ' PSC 25/02/2002 SYS4097 - Start
'            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'            Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
'            Set xmlBatchAuditNode = xmlDoc.createElement("BATCHAUDIT")
'            xmlCopyAttribute xmlBatchScheduleNode, xmlBatchAuditNode, "BATCHNUMBER"
'            xmlSetAttributeValue xmlBatchAuditNode, "BATCHRUNNUMBER", strValue
'            xmlTempRequest.appendChild xmlBatchAuditNode
'
'            On Error Resume Next
'            Call FindFailedApplications(xmlTempRequest, xmlTempResponse)
'
'            lngErrorNo = Err.Number
'            strSource = Err.Source
'            strDescription = Err.Description
'            On Error GoTo RunRateChangeBatchExit
'
'            If lngErrorNo <> 0 Then
'                If errGetOmigaErrorNumber(lngErrorNo) = oeRecordNotFound Then
'                    errThrowError strFunctionName, 575
'                Else
'                    Err.Raise lngErrorNo, strSource, strDescription
'                End If
'            End If
'            ' PSC 25/02/2002 SYS4097 - End
'
'            'BM0230 MDC 07/01/2003
'            Set xmlTempRequest = xmlDoc.createElement("BATCHSCHEDULE")
'            xmlCopyMandatoryAttribute xmlBatchAuditNode, xmlTempRequest, "BATCHNUMBER"
'            'xmlCopyMandatoryAttribute xmlBatchAuditNode, xmlTempRequest, "BATCHRUNNUMBER"
'            strBaseRateApplied = CStr(Abs(CInt(WasBaseRateApplied(xmlTempRequest))))
'            'BM0230 MDC 07/01/2003 - End
'
'            ' GHun 16/04/2002 SYS4368 strBatchRunNumber should be left as is, and not be set to
'            ' strValue as this is the FAILEDBATCHRUNNUMBER and audit records are then created
'            ' against the wrong run number
'            'strBatchRunNumber = strValue
'        End If
'
'        If Not blnExit Then
'            'Update the Batch Schedule
'            Set xmlNodeList = xmlTempResponse.selectNodes(".//RATECHANGE")
'            lngAppCount = xmlNodeList.length
'            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'            xmlSetAttributeValue xmlTempRequest, "OPERATION", "UpdateBatchSchedule"
'            Set xmlBatchScheduleNode = xmlDoc.createElement("BATCHSCHEDULE")
'            xmlSetAttributeValue xmlBatchScheduleNode, "BATCHNUMBER", strBatchNumber
'            xmlSetAttributeValue xmlBatchScheduleNode, "BATCHRUNNUMBER", strBatchRunNumber
'            xmlSetAttributeValue xmlBatchScheduleNode, "RETRY", "1"
'            xmlSetAttributeValue xmlBatchScheduleNode, "TOTALNUMBEROFRECORDS", CStr(lngAppCount)
'            xmlTempRequest.appendChild xmlBatchScheduleNode
'            strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
'            errCheckXMLResponse strResponse, True
'
'            Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
'            'BM0230 MDC 07/01/2003
'            If Len(strBaseRateApplied) > 0 Then
'                xmlSetAttributeValue xmlTempRequest, "BASERATEAPPLIED", strBaseRateApplied
'            End If
'            'BM0230 MDC 07/01/2003 - End
'
'            Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
'            Set xmlBatchNode = xmlGetMandatoryNode(xmlTempRequest, "BATCH")
'
'            ' PSC 25/02/02 SYS4097 - Start
'            strQueueName = GetMandatoryGlobalParamString("RateChangeQueueName")
'            If Len(Trim$(strQueueName)) = 0 Then
'                errThrowError strFunctionName, oeInvalidParameter, "RateChangeQueueName not found"
'            End If
'
'            xmlSetAttributeValue xmlBatchNode, "QUEUENAME", strQueueName
'            ' PSC 25/02/02 SYS4097 - End
'
'            'BM0367
'            'Get the RateChange batch delay and database Server Time once before the loop,
'            'so they dont have to be queried for every application
'            lRateChangeDelay = GetRateChangeBatchDelay()
'            dteDBServer = GetDatabaseServerDate()
'
'            'Work out the difference in time between the database and application server
'            lTimeDiff = DateDiff("s", Now(), dteDBServer)
'            'Convert the delay to seconds and add it to the time difference
'            lTimeDiff = lTimeDiff + (lRateChangeDelay * 60)
'
'            Set xmlMQLDelay = xmlDoc.createElement("MQLDELAY")
'            xmlMQLDelay.setAttribute "DELAY", lTimeDiff
'            'BM0367 End
'
'            'Process each application
'            Set objRateChangeNTTxBO = gobjContext.CreateInstance(App.Title & ".RateChangeNTTxBO")
'            For Each xmlNode In xmlNodeList
'                'Call NTTxBO to create BatchAudit, RateChangeBatchAudit records and
'                'send a corresponding message to the queue for each application
'                On Error Resume Next
'                xmlBatchNode.appendChild xmlNode
'                'BM0230 MDC 08/01/2003
'                If Len(Trim$(strValue)) > 0 Then
'                    xmlSetAttributeValue xmlNode, "FAILEDBATCHNUMBER", strValue
'                End If
'                'BM0230 MDC 08/01/2003 - End
'
'                'BM0367
'                xmlNode.appendChild xmlMQLDelay.cloneNode(True)
'                'BM0367
'
'                objRateChangeNTTxBO.ProcessApplication xmlTempRequest, xmlTempResponse
'                If Err.Number <> 0 Then
'                    'SYS3027 MDC 19/11/2001 - Improve error handling
'                    'Update Batch Audit with error information
'                    'GD 05/04/2002 SYS4334 - Start
'                    strErrDesc = Err.Description
'                    strErrNumber = Err.Number
'                    'GD 05/04/2002 SYS4334 - End
'                    strValue = Err.Source
'                    If Len(Trim$(strValue)) = 0 Or strValue = App.Title Then
'                        strValue = App.Title & ".ProcessApplication"
'                    End If
'                    Set xmlErrRequest = vxmlRequestNode.cloneNode(False)
'                    xmlSetAttributeValue xmlErrRequest, "OPERATION", "CreateBatchAudit"
'                    Set xmlBatchAuditNode = xmlDoc.createElement("BATCHAUDIT")
'                    xmlSetAttributeValue xmlBatchAuditNode, "ERRORNUMBER", CStr(Err.Number)
'                    xmlSetAttributeValue xmlBatchAuditNode, "ERRORSOURCE", strValue
'                    ' DRC 15/03/2002 SYS4214 - Start
'                    strApplicationNo = xmlGetAttributeText(xmlNode, "APPLICATIONNUMBER")
'                    strApplicationFFNo = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
'                    strQuotationNo = xmlGetAttributeText(xmlNode, "QUOTATIONNUMBER")
'                    strMortgageSubQNo = xmlGetAttributeText(xmlNode, "MORTGAGESUBQUOTENUMBER")
'                    strLoanComponentSeqNo = xmlGetAttributeText(xmlNode, "LOANCOMPONENTSEQUENCENUMBER")
'
'                    strDescription = "Application No: " & strApplicationNo & _
'                                     " Application Fact Find No: " & strApplicationFFNo & _
'                                     " Quotation No: " & strQuotationNo & _
'                                     " Mortgage Subquote No: " & strMortgageSubQNo & _
'                                     " LoanComponent No: " & strLoanComponentSeqNo
'                    ' DRC 01/03/2002 SYS4214 - End
'                    xmlSetAttributeValue xmlBatchAuditNode, "ERRORDESCRIPTION", strDescription
'                    Err.Clear
'                    On Error GoTo RunRateChangeBatchExit
'                    xmlSetAttributeValue xmlBatchAuditNode, "BATCHNUMBER", strBatchNumber
'                    xmlSetAttributeValue xmlBatchAuditNode, "BATCHRUNNUMBER", strBatchRunNumber
'                    xmlErrRequest.appendChild xmlBatchAuditNode
'                    strResponse = objBatchSchedBO.omBatchRequest(xmlErrRequest.xml)
'                    errCheckXMLResponse strResponse, True
'                    'SYS3027 - End
'                    'GD 05/04/2002 SYS4334 - Start
'                    'Error has occured in objRateChangeNTTxBO.ProcessApplication
'                    'so call IncrementBatchScheduleFailure
'                    ' Update the batch failure count
'
'                    'Load the string response from CreatebatchAudit into an xml so we can extract the BatchAuditGUID
'                    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
'                    'Pick off REQUEST only
'                    Set xmlTempBatchRequest = vxmlRequestNode.cloneNode(False)
'                    'Extract BATCHAUDIT node from call to CreateBatchAudit
'                    Set xmlTempNode = xmlGetMandatoryNode(xmlTempResponseDoc, ".//BATCHAUDIT")
'                    'Build a BatchSchedule Node
'                    Set xmlBatchSchedule = xmlTempResponseDoc.createElement("BATCHSCHEDULE")
'                    'Change the OPERATION
'                    xmlSetAttributeValue xmlTempBatchRequest, "OPERATION", "IncrementBatchScheduleFailure"
'                    xmlCopyMandatoryAttribute xmlTempNode, xmlBatchSchedule, "BATCHAUDITGUID"
'                    xmlSetAttributeValue xmlBatchSchedule, "ERRORNUMBER", strErrNumber
'                    xmlSetAttributeValue xmlBatchSchedule, "ERRORSOURCE", strValue
'                    xmlSetAttributeValue xmlBatchSchedule, "ERRORDESCRIPTION", strErrDesc & ":" & strDescription
'                    xmlSetAttributeValue xmlBatchSchedule, "BATCHNUMBER", strBatchNumber
'                    xmlSetAttributeValue xmlBatchSchedule, "BATCHRUNNUMBER", strBatchRunNumber
'                    xmlTempBatchRequest.appendChild xmlBatchSchedule
'
'                    strResponse = objBatchSchedBO.omBatchRequest(xmlTempBatchRequest.xml)
'
'                    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
'                    errCheckXMLResponseNode xmlTempResponseDoc.documentElement, vxmlResponseNode, True
'
'                    'GD 05/04/2002 SYS4334 - End
'
'                End If
'                'Remove node just processed from the request
'                xmlBatchNode.removeChild xmlNode
'            Next
'            On Error GoTo RunRateChangeBatchExit
'        End If
'    End If
'
'RunRateChangeBatchExit:
'
'    Set objBatchSchedBO = Nothing
'    Set xmlDoc = Nothing
'    Set xmlTempRequest = Nothing
'    Set xmlTempResponse = Nothing
'    Set xmlNode = Nothing
'    Set xmlNodeList = Nothing
'    Set xmlBatchNode = Nothing
'    Set xmlBatchAuditNode = Nothing
'    Set xmlBatchScheduleNode = Nothing
'    Set objMessageQueue = Nothing
'    Set objRateChangeNTTxBO = Nothing
'    Set xmlErrRequest = Nothing 'SYS3027
'    'GD 05/04/2002 SYS4334 - Start
'    Set xmlTempNode = Nothing
'    Set xmlTempBatchRequest = Nothing
'    Set xmlTempResponseDoc = Nothing
'    Set xmlBatchSchedule = Nothing
'    'GD 05/04/2002 SYS4334 - End
'    'BM0367
'    Set xmlMQLDelay = Nothing
'    'BM0367 End
'
'    'SYS3541 MDC 20/12/2001
'    If Err.Number <> 0 Then
'        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
'        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
'        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
'        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
'        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
'        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
'        App.LogEvent strErrMsg, vbLogEventTypeError
'    End If
'    'SYS3541 End
'
'    errCheckError strFunctionName, TypeName(Me)
'
'End Sub
'BMIDS622 End

Private Sub RunRateChangeBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    RunRateChangeBatch
' Author:       Mark Coleman, restructured by GHun
' Date:         05/09/2001
' Purpose:      Initiate a Rate Change batch job
' Input parameters:
'               vxmlRequestNode - XML Request
'               vxmlResponseNode - XML Response to be appended to
' Output parameters:
'               N/A
' Notes:        Calls a stored proc to create a BatchAudit record for each
'               application affected by a Rate Change and to create
'               RateChangeBatchAudit records for each Loan Component the
'               application quote has. It the posts a message to the
'               Rate Change message queue for processing.
'------------------------------------------------------------------------
    
    Const cstrFunctionName  As String = "RunRateChangeBatch"
    
    Dim objBatchSchedBO     As BatchScheduleBO
    Dim objRateChangeDO     As IRateChangeDO
    Dim objMessageQueue     As IomMessageQueue
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlNodeList         As IXMLDOMNodeList
    Dim xmlNode             As IXMLDOMNode
    Dim xmlTempResponse     As IXMLDOMNode
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlBatchSchedule    As IXMLDOMElement
    Dim xmlMessage          As IXMLDOMElement
    Dim xmlElement          As IXMLDOMElement
    Dim strValue            As String
    Dim strCancelledValueId As String
    Dim strResponse         As String
    Dim strBatchNumber      As String
    Dim strBatchRunNumber   As String
    Dim strBatchAuditGUID   As String
    Dim blnBatchCancelled   As Boolean
    Dim strErrMsg           As String
    Dim strQueueName        As String
    Dim strBaseRateApplied  As String
    Dim intMessageQueueType As Integer

On Error GoTo RunRateChangeBatchExit

    'Validation
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")

    'Create request to get Batch data
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.setAttribute "OPERATION", "GetBatchSchedule"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlBatchSchedule = xmlDoc.createElement("BATCHSCHEDULE")
    xmlBatchSchedule.setAttribute "BATCHNUMBER", strBatchNumber
    xmlBatchSchedule.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    xmlTempRequest.appendChild xmlBatchSchedule
  
    'Create and call the Batch Scheduler component to get the BatchSchedule
    Set objBatchSchedBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse
    Set xmlTempResponse = xmlDoc.documentElement
    
    Set xmlBatchSchedule = xmlGetMandatoryNode(xmlTempResponse, "BATCHSCHEDULE")
    
    'Check that Batch has not been cancelled
    strValue = xmlGetMandatoryAttributeText(xmlBatchSchedule, "STATUS")
    'Get the combo value for Cancelled
    strCancelledValueId = GetFirstComboValueId("BatchScheduleStatus", "CA")
    If strValue = strCancelledValueId Then
        'Batch Cancelled
        blnBatchCancelled = True
    End If
    
    If Not blnBatchCancelled Then
    
        'Check if this is a rerun of a failed batch
        strValue = xmlGetAttributeText(xmlBatchSchedule, "FAILEDBATCHRUNNUMBER")
        If Len(Trim$(strValue)) > 0 Then
            'Check if the base rate has already been applied
            Set xmlTempRequest = xmlDoc.createElement("BATCHSCHEDULEPROGRESS")
            xmlTempRequest.setAttribute "BATCHNUMBER", strBatchNumber
            strBaseRateApplied = CStr(Abs(CInt(WasBaseRateApplied(xmlTempRequest))))
        End If
        
        'Prepare the request and call PrepareRateChange, which calls a stored proc to
        'find the applications for rate change, and creates all the BatchAudit and
        'RateChangeBatchAudit records in one go
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlTempRequest.appendChild xmlBatchSchedule.cloneNode(True)
        Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
        Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
        
        objRateChangeDO.PrepareRateChange xmlTempRequest, xmlTempResponse
        errCheckXMLResponseNode xmlTempResponse, vxmlResponseNode, True

        'Select the returned BatchAudit nodes (one per application)
        Set xmlNodeList = xmlTempResponse.selectNodes("BATCH/BATCHAUDIT")
        
        If xmlNodeList.length = 0 Then
            
            'No Apps were found so complete the batch
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            xmlTempRequest.appendChild xmlBatchSchedule
            CompleteBatch xmlTempRequest, xmlTempResponse
        
        Else
        
            'Get MessageQueueType global parameter once outside of the for loop
            intMessageQueueType = GetGlobalParamAmount("MessageQueueType")
            If intMessageQueueType = 0 Then
                errThrowError cstrFunctionName, oeInvalidParameter, "MessageQueueType not found"
            End If
        
            'Get the RateChange queue name global parameter
            strQueueName = GetMandatoryGlobalParamString("RateChangeQueueName")
            If Len(Trim$(strQueueName)) = 0 Then
                errThrowError cstrFunctionName, oeInvalidParameter, "RateChangeQueueName not found"
            End If
        
            Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
        
            'Send a message to the message queue for each BatchAudit record returned
            For Each xmlNode In xmlNodeList
            
                strBatchAuditGUID = Trim$(xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID"))
                'Check if curly brackets need to be added as the GUID is represented in string form
                If Len(strBatchAuditGUID) = 36 Then
                    strBatchAuditGUID = "{" & strBatchAuditGUID & "}"
                End If
    
                'Setup the XML to be sent as a message
                Set xmlMessage = vxmlRequestNode.cloneNode(False)
                xmlMessage.setAttribute "OPERATION", "RunApplicationRateChange"
                Set xmlElement = xmlDoc.createElement("RATECHANGE")
                xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
                xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
                xmlElement.setAttribute "BATCHAUDITGUID", strBatchAuditGUID
                If Len(strBaseRateApplied) > 0 Then
                    xmlElement.setAttribute "BASERATEAPPLIED", strBaseRateApplied
                End If
                xmlMessage.appendChild xmlElement
            
                'Setup the Message Queue Request
                Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlTempRequest, "OPERATION", "SendToQueue"
                Set xmlElement = xmlDoc.createElement("MESSAGEQUEUE")
                xmlElement.setAttribute "QUEUENAME", strQueueName
                xmlElement.setAttribute "QUEUETYPE", intMessageQueueType
                xmlElement.setAttribute "PROGID", gstrRATECHANGE & ".RateChangeBO"
                xmlElement.setAttribute "XML", xmlMessage.xml
                xmlTempRequest.appendChild xmlElement
                
                'Send the message to the queue
                Set xmlTempResponse = objMessageQueue.SendToQueue(xmlTempRequest)
            
            Next
            
        End If
    End If
    
RunRateChangeBatchExit:

    Set objBatchSchedBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlBatchSchedule = Nothing
    Set objMessageQueue = Nothing
    Set objRateChangeDO = Nothing
    Set xmlMessage = Nothing
    Set xmlElement = Nothing
    
    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & cstrFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Batch Run Number: " & strBatchRunNumber
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    'SYS3541 End
       
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Public Function RateChangeBaseRequest(ByVal vstrXmlIn As String) As String

On Error GoTo RCBaseRequestVbErr

Dim xmlInDoc As FreeThreadedDOMDocument40
Dim xmlOutDoc As FreeThreadedDOMDocument40
Dim xmlRequestNode As IXMLDOMNode
Dim xmlOperationNode As IXMLDOMNode
Dim xmlOperationNodeList As IXMLDOMNodeList
Dim xmlResponseElem As IXMLDOMElement
Dim xmlResponseNode As IXMLDOMNode
Dim xmlAttrib As IXMLDOMAttribute

'Const strFunctionName As String = "RCBaseRequest"
    
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    xmlOutDoc.async = False
    
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    
    xmlInDoc.loadXML vstrXmlIn
    
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        DoRequest xmlRequestNode, xmlResponseNode
    Else
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttrib In xmlRequestNode.Attributes
                xmlOperationNode.Attributes.setNamedItem xmlAttrib.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseNode
        Next
    End If

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    RateChangeBaseRequest = xmlOutDoc.xml
    
RCBaseRequestExit:

'    Set gobjTxBO = Nothing
    Set gobjDO = Nothing
    
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlAttrib = Nothing
    
    Exit Function

RCBaseRequestVbErr:
    
    RateChangeBaseRequest = errCreateErrorResponse()
    
    GoTo RCBaseRequestExit
    
End Function

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    'Error Handler
    On Error GoTo DoRequestExit
    
    'Constant
    Const cstrFunctionName As String = "DoRequest"
        
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase$(strOperation)
    
    Select Case strOperation
    
        Case "RUNRATECHANGEBATCH"
            RunRateChangeBatch vxmlRequestNode, vxmlResponseNode
        
        Case "RUNAPPLICATIONRATECHANGE"
            RunApplicationRateChange vxmlRequestNode, vxmlResponseNode
      
        Case "FINDAPPLICATIONS"
            FindApplications vxmlRequestNode, vxmlResponseNode
'
'        Case "FINDFAILEDAPPLICATIONS"
'            FindFailedApplications vxmlRequestNode, vxmlResponseNode
'
'        Case "CREATERATECHANGEAUDIT"
'            CreateRateChangeAudit vxmlRequestNode, vxmlResponseNode
'
'        Case "FINDRATECHANGEAUDITLIST"
'            FindRateChangeAuditList vxmlRequestNode, vxmlResponseNode
'
'        Case "SETBASERATEAPPLIEDDATE"
'            SetBaseRateAppliedDate vxmlRequestNode, vxmlResponseNode
    
        'BMIDS622 GHun 12/08/2003
        Case "COMPLETEBATCH"
            CompleteBatch vxmlRequestNode, vxmlResponseNode
        'BMIDS622 End
    
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Function MessageQueueComponentVC2_OnMessage(ByVal in_xmlConfig As String, ByVal in_xmlData As String) As Long

On Error GoTo MessageQueueComponentVC2_OnMessageVbErr

'Const cstrFunctionName As String = "MessageQueueComponentVC2_OnMessage"

Dim strErrMsg As String
Dim strResponse As String

    strResponse = RateChangeBaseRequest(in_xmlData)
    
    'Check the response. If error, abort transaction and return suitable code to MQ
    errCheckXMLResponse strResponse, True
    
    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS  'Complete so remove message from queue
    
MessageQueueComponentVC2_OnMessage_Exit:

    Exit Function
    
MessageQueueComponentVC2_OnMessageVbErr:
    strErrMsg = "Error processing message." & vbCrLf & "Number: " & CStr(Err.Number) & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    App.LogEvent strErrMsg, vbLogEventTypeError
    
  'SYS4273 GHun 14/03/2002 Transactioning incorrect
    'gobjContext.SetAbort
    'MessageQueueComponentVC2_OnMessage = MESSQ_RESP_RETRY_MOVE_MESSAGE '3  'Fail and move message to another queue
    gobjContext.SetComplete
    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS  'Complete so remove message from queue
    'SYS4273 End
        
    GoTo MessageQueueComponentVC2_OnMessage_Exit

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

'BM0230 MDC 07/01/2003
Private Function WasBaseRateApplied(ByVal vxmlBatch As IXMLDOMNode) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    WasBaseRateApplied
' author:       Mark Coleman
' date:         07/01/2003
' description:
' pass:         vxmlBatch  XML Input request specifying the batch run to check
' return:       Boolean to indicate if the Base Rate was applied in the requested batch run
'------------------------------------------------------------------------------------------
On Error GoTo WasBaseRateApplied_Exit

Const cstrFunctionName As String = "WasBaseRateApplied"
    
Dim objBatchScheduleBO As BatchScheduleBO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlBatchRequest As IXMLDOMElement
Dim xmlTempResponse As IXMLDOMNode
Dim xmlBatchSchedule As IXMLDOMElement

Dim strResponse As String
'Dim lngSuccesses As Long
'Dim lngTotalRecords As Long
Dim blnReturn As Boolean

    'Create Request
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlBatchRequest = xmlDoc.createElement("REQUEST")
    'BM0438 TOTALNUMBEROFSUCCESSES is now stored on BatchScheduleProgress
    'xmlSetAttributeValue xmlBatchRequest, "OPERATION", "GetBatchSchedule"
    xmlSetAttributeValue xmlBatchRequest, "OPERATION", "GetBatchScheduleProgress"
    'BM0438 End
    
    xmlBatchRequest.appendChild vxmlBatch
    
    Set objBatchScheduleBO = gobjContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    
    'Make call to GetBatchSchedule
    strResponse = objBatchScheduleBO.omBatchRequest(xmlBatchRequest.xml)
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, xmlTempResponse, True
    
    'Determine if at least one item in the batch succeeded
    '(if so the baserateapplieddate would have been set)
    'BM0438
    'If xmlDoc.documentElement.selectNodes("BATCHSCHEDULE[@TOTALNUMBEROFSUCCESSES > 0]").length > 0 Then
    If xmlDoc.documentElement.selectNodes("BATCHSCHEDULEPROGRESS[@TOTALNUMBEROFSUCCESSES > 0]").length > 0 Then
    'BM0438 End
'    Set xmlBatchSchedule = xmlGetMandatoryNode(xmlDoc.documentElement, "BATCHSCHEDULE")
'    If xmlGetAttributeAsLong(xmlBatchSchedule, "TOTALNUMBEROFSUCCESSES") > 0 Then
        blnReturn = True
    End If
    
    WasBaseRateApplied = blnReturn
    
WasBaseRateApplied_Exit:
    Set xmlDoc = Nothing
    Set xmlBatchRequest = Nothing
    Set objBatchScheduleBO = Nothing
    Set xmlTempResponse = Nothing
    Set xmlBatchSchedule = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)

End Function
'BM0230 MDC 07/01/2003 - End

'BMIDS622 GHun 18/08/2003 No longer used
''BM0367
'Private Function GetRateChangeBatchDelay() As Long
'' header ----------------------------------------------------------------------------------
'' procedure:    GetRateChangeBatchDelay
'' author:       GHun
'' date:         20/02/2003
'' description:  Read the number of minutes for a RateChange batch to delay from the registry
''               The delay is used to reduce concurrency to avoid deadlocks and timeouts
'' pass:         nothing
'' return:       Long value expressed in minutes
''------------------------------------------------------------------------------------------
'    Const cstrRegKey        As String = "HKLM\SOFTWARE\Omiga4\System Configuration\RateChangeBatchDelay"
'    Const clDefaultDelay    As Long = 10
'    Dim objWshShell         As Object
'    Dim strRateChangeDelay  As String
'    Dim lRateChangeDelay    As Long
'
''Ignore any errors that occur in this function
'On Error Resume Next
'
'    Set objWshShell = CreateObject("WScript.Shell")
'    strRateChangeDelay = Trim(objWshShell.RegRead(cstrRegKey))
'    If Len(strRateChangeDelay) = 0 Then
'        lRateChangeDelay = clDefaultDelay
'    Else
'        If IsNumeric(strRateChangeDelay) Then
'            lRateChangeDelay = Abs(CLng(strRateChangeDelay))
'        Else
'            lRateChangeDelay = clDefaultDelay
'        End If
'    End If
'
'    GetRateChangeBatchDelay = lRateChangeDelay
'
'    Set objWshShell = Nothing
'End Function
''BM0367 End

'BMIDS622 GHun 18/08/2003 No longer used
''BM0367
'Private Function GetDatabaseServerDate() As Date
'' header ----------------------------------------------------------------------------------
'' procedure:    GetDatebaseServerDate
'' author:       GHun
'' date:         20/02/2003
'' description:  Call RateChangeDO to get the current date and time from the database
'' pass:         nothing
'' return:       Date
''------------------------------------------------------------------------------------------
'    Const cstrFunctionName  As String = "GetDatabaseServerDate"
'    Dim objRateChangeDO     As IRateChangeDO
'    Dim xmlServer           As IXMLDOMElement
'    Dim strDate             As String
'    Dim xmlDoc              As FreeThreadedDOMDocument40
'    Dim xmlResponse         As IXMLDOMElement
'
'On Error GoTo GetDatabaseServerDateExit
'
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    Set xmlResponse = xmlDoc.createElement("RESPONSE")
'
'    Set objRateChangeDO = gobjContext.CreateInstance(App.Title & ".RateChangeDO")
'    objRateChangeDO.GetDatabaseServerDate xmlResponse
'    Set xmlServer = xmlGetMandatoryNode(xmlResponse, "SERVER")
'    'User App Server Date (Now()) if DB Server date is not returned
'    GetDatabaseServerDate = xmlGetAttributeAsDate(xmlServer, "DATE", Now())
'
'GetDatabaseServerDateExit:
'
'    Set objRateChangeDO = Nothing
'    Set xmlServer = Nothing
'    Set xmlDoc = Nothing
'    Set xmlResponse = Nothing
'
'    errCheckError cstrFunctionName, TypeName(Me)
'End Function
''BM0367 End

'BMIDS622 GHun 12/08/2003
Private Sub QueueBatchCompletion(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    QueueBatchCompletion
' author:       GHun
' date:         12/08/2003
' description:  Send a message to the message queue to call CompleteBatch so that it runs in a separate transaction
' pass:         vxmlRequestNode XML Request to be passed on to CompleteBatch
' return:       nothing
'------------------------------------------------------------------------------------------
    Const cstrFunctionName  As String = "QueueBatchCompletion"
    Dim objMessageQueue     As IomMessageQueue
    Dim strQueueName        As String
    Dim xmlMessage          As IXMLDOMElement
    Dim xmlMQRequest        As IXMLDOMElement
    Dim xmlMQResponse       As IXMLDOMNode
    Dim xmlElement          As IXMLDOMElement
    Dim xmlDoc              As FreeThreadedDOMDocument40

On Error GoTo QueueBatchCompletionExit

    'Get queue name & setup message queue request
    strQueueName = GetGlobalParamString("RateChangeQueueName")
    If Len(Trim$(strQueueName)) = 0 Then
        errThrowError cstrFunctionName, oeInvalidParameter, "RateChangeQueueName not found"
    End If

    'Set up XML to be sent as a message
    Set xmlMessage = vxmlRequestNode.cloneNode(True)
    xmlSetAttributeValue xmlMessage, "OPERATION", "CompleteBatch"

    'Setup message queue request
    Set xmlMQRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlMQRequest, "OPERATION", "SendToQueue"
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlElement = xmlDoc.createElement("MESSAGEQUEUE")
    xmlElement.setAttribute "QUEUENAME", strQueueName
    xmlElement.setAttribute "PROGID", gstrRATECHANGE & ".RateChangeBO"
    xmlElement.setAttribute "XML", xmlMessage.xml
    xmlMQRequest.appendChild xmlElement
    
    'Send the message to the queue
    Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
    Set xmlMQResponse = objMessageQueue.SendToQueue(xmlMQRequest)

QueueBatchCompletionExit:
    Set objMessageQueue = Nothing
    Set xmlMessage = Nothing
    Set xmlMQRequest = Nothing
    Set xmlMQResponse = Nothing
    Set xmlElement = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BMIDS622 End

'BMIDS622 DoClientRequest should be a sub, because it does not return anything. However OOSS
' merges the code incorrectly if it is changed, so the return type has been set to boolean
Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As Boolean
On Error GoTo DoClientRequestExit
    
    Const cstrFunctionName As String = "DoClientRequest"

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        cstrFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    errCheckError cstrFunctionName, TypeName(Me)

End Function
