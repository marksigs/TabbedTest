VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "BatchScheduleTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BatchScheduleTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Dependencies:  BatchScheduleDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     02/04/01    SYS2201: Created
'MV     01/05/01    Added New Method called IncrementBatchAuditRetries
'MV     11/05/01    Changed RunBatch and LaunchBatch methods to Amend the USERID ,UNITID,MACHINEID,OPERATION
'                   in the Request
'MDC    12/09/2001  SYS2250 - Rate Change Batch
'MC     05/12/01    SYS3018 - Fixes to Batch Process for SQL Server.
'DM     06/12/2001 SYS3185 Add CreateBatchException
'MC     20/12/2001  SYS3541 - Write error info to event log
'SA     18/01/2002  SYS3327 - Removed Stage 1 in LaunchBatch as it is unnecessary.
'SA     24/01/2002  SYS3820 - CompleteRunBatch should set status to new value "In Progress"
'                             Database trigger set up to update to complete.
'GHun   26/03/2002  SYS4305 - Batch status stays as launched even after batch completed
'STB    10/04/2002  SYS4357 - DateTime datatypes are converted to American format by the DOM.
'GHun   15/04/2002  SYS4368 - FAILEDBATCHRUNNUMBER should be used instead of FAILEDBATCHNUMBER
'GHun   22/04/2002  SYS4441 - Valuer payments - Batch status remains as Running
'DJP    07/05/2002  SYS4534 - Don't write Now (Date datatype) to an attribute value.
'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'GHun   24/03/2003  BM0438 - Updated LaunchBatch, IncrementBatchScheduleSuccess/Failure, UpdateBatchSchedule
'GHun   09/09/2003  BMIDS622 RateChange restructure and clean up
'------------------------------------------------------------------------------------------
'Mars History:
'
'Prog   Date        Description
'GHun   22/02/2006  MAR1090 Fix rescheduling of batches
'------------------------------------------------------------------------------------------

Option Explicit

Implements IBatchScheduleTXBO

Implements ObjectControl
Private gobjContext As ObjectContext
Private objDO As IBatchScheduleDO

Private Function CallBatchProcessByName(ByVal intProgramType As Integer, ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode
    
    'Error Handler
    On Error GoTo CallBatchProcessByNameExit
    
    'Cosntant
    Const strFunctionName As String = "CallBatchProcessByName"
    
    'Declare Variables
    Dim blnPhase2 As Boolean
    Dim strBatchProgramObjectCall As String
    Dim strBatchProgramMethodCall As String
    'Dim strBatchProgramOperationCall  As String
    Dim strReponse  As String
    Dim strErrMsg As String 'SYS3541 MDC 20/12/2001
    
    Dim objBatchProgramObjectCall As Object
    
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlDoc As FreeThreadedDOMDocument40
    
'Stage1:

    'checking for the Phase2 component
    blnPhase2 = IsValidationType("BatchProgramObjectCall", intProgramType, "2")
    
    strBatchProgramObjectCall = GetComboText("BatchProgramObjectCall", intProgramType)
    strBatchProgramMethodCall = GetComboText("BatchProgramMethodCall", intProgramType)
    
    'If it is Phase2 Component then
    If blnPhase2 Then
        Set xmlAttrib = vxmlRequestNode.ownerDocument.createAttribute("OPERATION")
        xmlAttrib.Value = GetComboText("BatchProgramOperationCall", intProgramType)
        vxmlRequestNode.Attributes.setNamedItem xmlAttrib
        Set xmlAttrib = Nothing
    End If
 
'Stage2:

    'SYS3018 MDC 04/12/2001 -   Call Batch Process asynchronously allowing current
    '                           transaction to complete (releasing database locks)
    'Save the ProgId, MethodName and Request XML
    g_strProgId = strBatchProgramObjectCall
    g_strMethodName = strBatchProgramMethodCall
    g_strRequest = vxmlRequestNode.xml
    
    'Set a timer to run subroutine CallBatchProcessAsync in ½ second...
    g_lngTimerId = SetTimer(0, 0, 500, AddressOf CallBatchProcessAsync)
    If g_lngTimerId = 0 Then
        'Error setting timer
        errThrowError strFunctionName, oeInternalError, "Unable to call batch process asynchronously"
    End If
    
    '... and continue
    strReponse = "<RESPONSE TYPE='SUCCESS'/>"
'    Set objBatchProgramObjectCall = gobjContext.CreateInstance(strBatchProgramObjectCall)
'
'    strReponse = CallByName(objBatchProgramObjectCall, strBatchProgramMethodCall, VbMethod, vxmlRequestNode.xml)
'
    'SYS3018 End
    
    ' PSC 02/05/01 - Start
    Set xmlDoc = xmlLoad(strReponse, strFunctionName)
    Set CallBatchProcessByName = xmlDoc.documentElement
    ' PSC 02/05/01 - End

CallBatchProcessByNameExit:

    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Calling: " & strBatchProgramObjectCall
        strErrMsg = strErrMsg & vbCrLf & "Method: " & strBatchProgramMethodCall
        strErrMsg = strErrMsg & vbCrLf & "Request: " & vxmlRequestNode.xml
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    'SYS3541 End

    Set objBatchProgramObjectCall = Nothing
    Set xmlAttrib = Nothing
    Set xmlDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Function
    
End Function

Private Sub IBatchScheduleTXBO_CompleteRunBatch(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    IBatchScheduleTXBO_CompleteRunBatch
' author:       Mark Coleman
' date:         06/12/2001
' description:  When Batch has been run either update status to 'Complete' or if a
'               regular batch job, launch the next one.
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------

On Error GoTo CompleteRunBatchExit

Const strFunctionName As String = "IBatchScheduleTXBO_CompleteRunBatch"
    
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlBatchNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlNewWorkingDate As IXMLDOMNode

Dim cintValueIds As Collection

Dim strBatchNumber As String
Dim strBatchStatus As String
Dim intBatchFrequency As String
Dim strExecutionDateTime As String
Dim dtNextExecutionDateTime As Date
Dim intBatchStatus As Integer
Dim intIndex As Integer

Dim xmlSystemDateRequest As IXMLDOMElement
Dim xmlTempNode1 As IXMLDOMNode
Dim xmlTempNode2 As IXMLDOMNode
Dim xmlNextWorkingDate As IXMLDOMNode
Dim strBatchProgramChannelId  As String

Dim objSystemDatesBO  As Object
Dim strNextWorkingDate  As String   'MAR1090 GHun
Dim xmlOutputDoc As FreeThreadedDOMDocument40   'MAR1090 GHun

    'Create Temp Response Element
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    xmlSetAttributeValue xmlTempResponse, "TYPE", "SUCCESS"

    Set xmlBatchNode = xmlGetMandatoryNode(vxmlRequestNode, ".//BATCH")
    strBatchNumber = xmlGetAttributeText(xmlBatchNode, "BATCHNUMBER")

    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    'Get Batch Details
    objDO.GetBatch vxmlRequestNode, xmlTempResponse
    errCheckXMLResponse xmlTempResponse.xml, True
    
    Set xmlBatchNode = xmlGetMandatoryNode(xmlTempResponse, ".//BATCH")
    strBatchStatus = xmlGetAttributeText(xmlBatchNode, "BATCHSTATUS")
    intBatchFrequency = xmlGetAttributeAsInteger(xmlBatchNode, "FREQUENCY")
    strExecutionDateTime = xmlGetAttributeText(xmlBatchNode, "BATCHSCHEDULEEXECUTIONDATETIME")  'MAR1090 GHun
    
    Set cintValueIds = New Collection
    GetValueIdsForValidationType "BatchScheduleStatus", "R", cintValueIds
    
    ' If the Batch is not cancelled then
    If Not IsValidationType("BatchStatus", CInt(strBatchStatus), "CA") Then
    
        'If the RunBatchFrequency is not Adhoc then
        If Not IsValidationType("BatchFrequency", intBatchFrequency, "A") Then  'MAR1090 GHun

            'MAR1090 GHun Don't reschedule reruns of failed batches
            If Len(xmlGetAttributeText(xmlBatchNode, "FAILEDBATCHRUNNUMBER", vbNullString)) = 0 Then
            
                'If the RunBatch Frequency is on Every Working Day then
                If IsValidationType("BatchFrequency", intBatchFrequency, "WD") Then
            'MAR1090 End
                    
                    strBatchProgramChannelId = GetGlobalParamString("BatchProgramChannelId")
    
                    'Preparing Request Element
                    Set xmlSystemDateRequest = xmlDoc.createElement("REQUEST")
                    xmlDoc.appendChild xmlSystemDateRequest
    
                    Set xmlTempNode1 = xmlDoc.createElement("SYSTEMDATE")
                    xmlSystemDateRequest.appendChild xmlTempNode1
    
                    Set xmlTempNode2 = xmlDoc.createElement("DATE")
                    xmlTempNode2.Text = strExecutionDateTime
                    xmlTempNode1.appendChild xmlTempNode2
    
                    Set xmlTempNode2 = xmlDoc.createElement("DIRECTION")
                    xmlTempNode2.Text = "+"
                    xmlTempNode1.appendChild xmlTempNode2
    
                    Set xmlTempNode2 = xmlDoc.createElement("OFFSET")
                    xmlTempNode2.Text = "0"
                    xmlTempNode1.appendChild xmlTempNode2
    
                    Set xmlTempNode2 = xmlDoc.createElement("CHANNELID")    'MAR1090 GHun
                    xmlTempNode2.Text = strBatchProgramChannelId
                    xmlTempNode1.appendChild xmlTempNode2
    
                    'Create an Instance of ombase
                    Set objSystemDatesBO = gobjContext.CreateInstance("omBase.SystemDatesBO")
    
                    'Find next Working Day
                    'MAR1090 GHun
                    'Set xmlNextWorkingDate = objSystemDatesBO.FindWorkingDay(xmlSystemDateRequest.xml)
                    strNextWorkingDate = objSystemDatesBO.FindWorkingDay(xmlSystemDateRequest.xml)
                    Set xmlOutputDoc = New FreeThreadedDOMDocument40
                    xmlOutputDoc.loadXML strNextWorkingDate
                    errCheckXMLResponse strNextWorkingDate, True
                    Set xmlNewWorkingDate = xmlOutputDoc.selectSingleNode(".//SYSTEMDATE")
                    'MAR1090 End
                    
                    dtNextExecutionDateTime = xmlGetNodeAsDate(xmlNewWorkingDate, "DATE")
                      
                Else
                    'MAR1090 GHun
                    'dtNextExecutionDateTime = CDate(strExecutionDateTime) + Val(GetValidationTypeForValueID("BatchFrequency", strBatchFrequency))
                    If IsValidationType("BatchFrequency", intBatchFrequency, "D") Then
                        dtNextExecutionDateTime = DateAdd("d", 1, CDate(strExecutionDateTime))
                    ElseIf IsValidationType("BatchFrequency", intBatchFrequency, "W") Then
                        dtNextExecutionDateTime = DateAdd("ww", 1, CDate(strExecutionDateTime))
                    ElseIf IsValidationType("BatchFrequency", intBatchFrequency, "M") Then
                        dtNextExecutionDateTime = DateAdd("m", 1, CDate(strExecutionDateTime))
                    ElseIf IsValidationType("BatchFrequency", intBatchFrequency, "Y") Then
                        dtNextExecutionDateTime = DateAdd("yyyy", 1, CDate(strExecutionDateTime))
                    End If
                    'MAR1090 End
                End If
    
                Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    
                Set xmlBatchNode = xmlDoc.createElement("BATCH")
                xmlSetAttributeValue xmlBatchNode, "BATCHNUMBER", strBatchNumber
                xmlSetAttributeValue xmlBatchNode, "EXECUTIONDATETIME", CStr(dtNextExecutionDateTime)
                xmlTempRequest.appendChild xmlBatchNode
    
                IBatchScheduleTXBO_LaunchBatch xmlTempRequest, vxmlResponseNode
                errCheckXMLResponse vxmlResponseNode.xml, True
            End If  'MAR1090 GHun
            
        Else
        
            'SYS4305 Only update the status if the batch is not complete
            If Not IsValidationType("BatchStatus", CInt(strBatchStatus), "C") Then
            
                'SYS3820 Update to New Status "In Progress"
                'Update Batch with the new status to Complete
                'GetValueIdsForValidationType "BatchScheduleStatus", "C", cintValueIds
                
                'SYS4305 "In Progress" is a BatchStatus, not a BatchScheduleStatus
                'GetValueIdsForValidationType "BatchScheduleStatus", "I", cintValueIds
                GetValueIdsForValidationType "BatchStatus", "I", cintValueIds
    
                For intIndex = 1 To cintValueIds.Count
                    intBatchStatus = CInt(cintValueIds.Item(intIndex))
                Next
    
                'Create Response Node
                Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
                xmlSetAttributeValue xmlTempResponse, "TYPE", "SUCCESS"
    
                'Create Request Node
                Set xmlBatchNode = xmlDoc.createElement("BATCH")
                xmlSetAttributeValue xmlBatchNode, "BATCHNUMBER", strBatchNumber
                xmlSetAttributeValue xmlBatchNode, "STATUS", intBatchStatus
    
                'Update Batch with new Status
                objDO.UpdateBatch xmlBatchNode, xmlTempResponse
                errCheckXMLResponse xmlTempResponse.xml, True
    
                Set objDO = Nothing
            End If

        End If

    End If

CompleteRunBatchExit:
    Set xmlDoc = Nothing
    Set xmlNewWorkingDate = Nothing
    Set xmlBatchNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set cintValueIds = Nothing
    Set xmlSystemDateRequest = Nothing
    Set xmlTempNode1 = Nothing
    Set xmlTempNode2 = Nothing
    Set xmlNextWorkingDate = Nothing
    Set objSystemDatesBO = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IBatchScheduleTxBO_CreateBatchSchedule(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreateBatchScheduleVbErr
    
    Const strFunctionName As String = "IBatchScheduleTxBO_CreateBatchSchedule"
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    objDO.CreateBatchSchedule vxmlRequestParentNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
CreateBatchScheduleExit:
    
    Set objDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateBatchScheduleVbErr:

    gobjContext.SetAbort
    
    GoTo CreateBatchScheduleExit
End Sub

'BM0438
Private Sub IBatchScheduleTXBO_IncrementBatchScheduleFailure(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "IBatchScheduleTXBO_IncrementBatchScheduleFailure"
    Dim objBatchScheduleDO As IBatchScheduleDO

On Error GoTo IncrementBatchScheduleFailureVbErr

    Set objBatchScheduleDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    
    objBatchScheduleDO.IncrementBatchScheduleFailure vxmlRequestParentNode, vxmlResponseNode
    gobjContext.SetComplete
    
IncrementBatchScheduleFailureExit:
    Set objBatchScheduleDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

IncrementBatchScheduleFailureVbErr:
    gobjContext.SetAbort
    GoTo IncrementBatchScheduleFailureExit
End Sub
'BM0438 End

'BM0438
Private Sub IBatchScheduleTXBO_IncrementBatchScheduleSuccess(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "IBatchScheduleTXBO_IncrementBatchScheduleSuccess"
    Dim objBatchScheduleDO As IBatchScheduleDO

On Error GoTo IncrementBatchScheduleSuccessVbErr

    Set objBatchScheduleDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    
    objBatchScheduleDO.IncrementBatchScheduleSuccess vxmlRequestParentNode, vxmlResponseNode
    gobjContext.SetComplete
    
IncrementBatchScheduleSuccessExit:
    Set objBatchScheduleDO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

IncrementBatchScheduleSuccessVbErr:
    gobjContext.SetAbort
    GoTo IncrementBatchScheduleSuccessExit
End Sub
'BM0438 End

Private Sub IBatchScheduleTXBO_IncrementBatchAuditRetries(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
  
    On Error GoTo IncrementBatchAuditRetriesVbErr
    
    Const strFunctionName As String = "IBatchScheduleTXBO_IncrementBatchAuditRetries"
        
    'Declare Variables
    Dim lngTotalNumberOfRetries As Long
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchAuditGUID As String
    
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseElem  As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    'SYS4441 2 new variables required
    Dim xmlBatchAudit As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    
    'SYS4441 Use xmlBatchAudit variable rather than overwriting vxmlRequestParent
    'Set vxmlRequestParentNode = vxmlRequestParentNode.selectSingleNode(".//BATCHAUDIT")
    Set xmlBatchAudit = vxmlRequestParentNode.selectSingleNode(".//BATCHAUDIT")
        
    'Store into Local Variables
    strBatchNumber = xmlGetAttributeText(xmlBatchAudit, "BATCHNUMBER")
    strBatchRunNumber = xmlGetAttributeText(xmlBatchAudit, "BATCHRUNNUMBER")
    strBatchAuditGUID = xmlGetAttributeText(xmlBatchAudit, "BATCHAUDITGUID")
    
'Stage1:
    
    'Preparing Request Element
    'SYS4441 Clone the REQUEST node from the node passed in
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlElement = xmlTempDoc.createElement("BATCHAUDIT")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    xmlElement.setAttribute "BATCHAUDITGUID", strBatchAuditGUID
    
    'SYS4441 Append xmlElement as a child of xmlTempElement
    xmlTempElement.appendChild xmlElement
    
    'Preparing Response Element
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlTempDoc.appendChild xmlTempResponseElem
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    'Call FindBatchAuditList
    'SYS4441 Pass in xmlTempElement rather than xmlElement as FindBatchAuditList expects
    'the BATCHAUDIT node to be a child node
    objDO.FindBatchAuditList xmlTempElement, xmlTempResponseElem
    If xmlTempResponseElem.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    'SYS4441 Retrieve the existing number of retries
    lngTotalNumberOfRetries = xmlGetAttributeAsLong(xmlTempResponseElem.selectSingleNode("BATCHAUDIT"), "NUMBEROFRETRIES")
    
    Set xmlTempResponseElem = Nothing
    Set xmlTempDoc = Nothing
    
'Stage2:
    
    'Increment NoOfRetries by 1
    lngTotalNumberOfRetries = lngTotalNumberOfRetries + 1
    
'Stage3:
    
    'Update Batchaudit Table with the new No Of retries
    
    'Preparing Request Element
    xmlElement.setAttribute "NUMBEROFRETRIES", lngTotalNumberOfRetries
    
    'Preparing Response Element
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlTempDoc.appendChild xmlTempResponseElem
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    objDO.UpdateBatchAudit xmlElement, xmlTempResponseElem
    errCheckXMLResponseNode xmlTempResponseElem, , True
    
    Set xmlTempDoc = Nothing
    Set xmlTempResponseElem = Nothing
    
    
IncrementBatchAuditRetriesExit:

    Set xmlTempDoc = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set objDO = Nothing
    Set xmlBatchAudit = Nothing
    Set xmlTempElement = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
IncrementBatchAuditRetriesVbErr:

    gobjContext.SetAbort

    GoTo IncrementBatchAuditRetriesExit
End Sub

Private Sub IBatchScheduleTXBO_RunBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Variables
    Dim strBatchNumber As String
    Dim strBatchRunNumber  As String
    Dim strBatchScheduleStatus   As String
    Dim strProgramType As String
    Dim strPaymentJobType As String
    
    Dim intBatchScheduleStatus As Integer
    Dim intIndex As Integer
    
    Dim cintValueIds As Collection
    
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlResponseNode  As IXMLDOMNode
    Dim xmlTempElement  As IXMLDOMElement
    Dim xmlTempBatchNode As IXMLDOMNode
    
    'SYS3018 MDC 06/12/2001 - Remove code not required
    'Dim strBatchStatus As String
    'Dim strBatchFrequency As String
    'Dim strExecutionDateTime As String
    'Dim strBatchFrequencyComboValue  As String
    'Dim strBatchProgramMethodCall  As String
    'Dim strBatchProgramObjectCall As String
    'Dim strResponse As String
    'Dim obj As Object
    'Dim dtNextExecutionDateTime As Date
    'Dim xmlNewWorkingDate As IXMLDOMNode
    'Dim xmlTempElem As IXMLDOMElement
    'Dim intBatchStatus As Integer
    'SYS3018 End
    
    'Error Handler
    On Error GoTo RunBatchExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleTXBO_RunBatch"
    
    'Store into Local Variables
    Set xmlTempBatchNode = vxmlRequestParentNode.selectSingleNode(".//BATCH")
    strBatchNumber = xmlGetAttributeText(xmlTempBatchNode, "BATCHNUMBER")
    
'Stage1:
    
    'Create Temp Response Element
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    'xmlTempDoc.appendChild xmlTempResponseElem
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    'Get Batch Details
    objDO.GetBatch xmlTempBatchNode, xmlTempResponseElem
    errCheckXMLResponse xmlTempResponseElem.xml, True
    
    'Load Batch Node from the Response
    Set xmlNode = xmlTempResponseElem.selectSingleNode(".//BATCH")
    
    'Store into local Variables
    strProgramType = xmlGetAttributeText(xmlNode, "PROGRAMTYPE")
    'MV - 02/05/01
    
    'SYS3018 MDC 06/12/2001 - Remove code not required
'    strBatchStatus = xmlGetAttributeText(xmlNode, "BATCHSTATUS")
'    strBatchFrequency = xmlGetAttributeText(xmlNode, "FREQUENCY")
'    strExecutionDateTime = xmlGetAttributeText(xmlNode, "BATCHEXECUTIONDATETIME")
    'SYS3018 End
    
    strPaymentJobType = xmlGetAttributeText(xmlNode, "PAYMENTJOBTYPE")
    strBatchRunNumber = xmlGetAttributeText(xmlNode, "BATCHRUNNUMBER")
    strBatchScheduleStatus = xmlGetAttributeText(xmlNode, "BATCHSCHEDULESTATUS")
    
'Stage2:
    
    'If the BatchSchedule.Status is not CANCELLED then
     If Not IsValidationType("BatchScheduleStatus", CInt(strBatchScheduleStatus), "CA") Then
        
        'Update Batch Schedule Status
        
        'Preparing Request Element
        Set cintValueIds = New Collection
        GetValueIdsForValidationType "BatchScheduleStatus", "R", cintValueIds
        For intIndex = 1 To cintValueIds.Count
            intBatchScheduleStatus = CInt(cintValueIds.Item(intIndex))
        Next
        
        'Create Response Node
         Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
        xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
        'xmlTempDoc.appendChild xmlTempResponseElem
    
        'Create Request Node
        Set xmlElement = xmlTempDoc.createElement("BATCHSCHEDULE")
        xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
        xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
        xmlElement.setAttribute "STATUS", intBatchScheduleStatus
        
        'Call UpdateBatchSchedule with New Status
        objDO.UpdateBatchSchedule xmlElement, xmlTempResponseElem
        errCheckXMLResponse xmlTempResponseElem.xml, True
        
    End If
    
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlElement = xmlTempDoc.createElement("BATCH")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    
    If CInt(strProgramType) = 20 Then
        'PaymentProcessing
        xmlElement.setAttribute "PAYMENTJOBTYPE", strPaymentJobType
    End If
    
    xmlTempElement.appendChild xmlElement
    
    Set xmlResponseNode = CallBatchProcessByName(CInt(strProgramType), xmlTempElement)
    errCheckXMLResponse xmlResponseNode.xml, True
    
'Stage3:
    
RunBatchExit:
    
    Set xmlTempDoc = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set objDO = Nothing
    
    'SYS3018 MDC 06/12/2001 - Dereference all objects
    Set cintValueIds = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempElement = Nothing
    Set xmlTempBatchNode = Nothing
    'SYS3018 End
    
    gobjContext.SetComplete
    
    'SYS3018 MDC 04/12/2001
    If Err.Number <> 0 Then
        'Write to Event Log
        App.LogEvent vbCr & "Error: " & Err.Number & vbCr & "Source: " & Err.Source _
                    & vbCr & "Description: " & Err.Description, vbLogEventTypeError
    End If
    'SYS3018 End
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IBatchScheduleTXBO_SetBatchStatus(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo SetBatchStatusVbErr
    
    Const strFunctionName As String = "IBatchScheduleTXBO_SetBatchStatus"
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    objDO.SetBatchStatus vxmlRequestParentNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
SetBatchStatusExit:
    
    Set objDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
SetBatchStatusVbErr:

    gobjContext.SetAbort

    GoTo SetBatchStatusExit
    
End Sub

Private Sub IBatchScheduleTXBO_LaunchBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Variables
    Dim strBatchNumber As String
    Dim strExecuteBatchDateTime  As String
    Dim strFailedBatchRunNumber  As String
    'Dim intBatchStatus As Integer
    Dim strMasterQueueName As String
   
    Dim dtToDay As Date
    
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseElem  As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    'Dim xmlNode As IXMLDOMNode
    Dim xmlElement1 As IXMLDOMElement
    Dim xmlBatchRequestNode As IXMLDOMNode
    Dim objMessageQueue As IomMessageQueue
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    
    'Error Handler
On Error GoTo LaunchBatchVBErr
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleTXBO_LaunchBatch"
    
    'Store into Local Variables
    Set xmlBatchRequestNode = xmlGetMandatoryNode(vxmlRequestParentNode, "BATCH")
    strBatchNumber = xmlGetAttributeText(xmlBatchRequestNode, "BATCHNUMBER")
    
    If (xmlGetAttributeText(xmlBatchRequestNode, "EXECUTIONDATETIME") = "") Or (IsNull(xmlGetAttributeText(xmlBatchRequestNode, "EXECUTIONDATETIME"))) Then
        strExecuteBatchDateTime = Now()
    Else
        strExecuteBatchDateTime = CSafeDate(xmlGetAttributeText(xmlBatchRequestNode, "EXECUTIONDATETIME"))
    End If
    
    'SYS4368 Changed FAILEDBATCHNUMBER to FAILEDBATCHRUNNUMBER
    'strFailedBatchNumber = xmlGetAttributeText(xmlBatchRequestNode, "FAILEDBATCHNUMBER")
    strFailedBatchRunNumber = xmlGetAttributeText(xmlBatchRequestNode, "FAILEDBATCHRUNNUMBER")

'SYS3327 SA Stage 1 is unnecessary. All it is doing is checking the status of the batch.
'If launched from Supervisor, this method cannot be called unless the status is correct (ie "Created")
'in other words - the status checking is done at GUI level.
'Also, the GetBatch method uses the view GetBatch which assumes
'BatchSchedule records have been created - they haven't - that happens
' in stage 2!!
'Stage1:
    
    'Get Batch Info
    
'    'Preparing Request Element
'    Set xmlElement = xmlTempDoc.createElement("BATCH")
'    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
'
'    'Call Get Batch
'    If objDO Is Nothing Then
'        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
'    End If
'
'    objDO.GetBatch xmlElement, vxmlResponseNode
'    If vxmlResponseNode.hasChildNodes = False Then
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'
'    'load the BatchSchedule Node into TempVariable
'    Set xmlNode = vxmlResponseNode.selectSingleNode(".//BATCH")
'
'    'Store into Local Variables
'    intBatchStatus = xmlGetAttributeText(xmlNode, "BATCHSTATUS")
'
'    'If the Status is Cancelled then throw an Error
'    If IsValidationType("BatchStatus", intBatchStatus, "C") Then
'         errThrowError strFunctionName, 306
'    End If
    
'Stage2:

    'Create Batch Schedule with Status as Waiting
    dtToDay = Now()

    'Preparing Request Element
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    
    Set xmlElement = xmlTempDoc.createElement("BATCHSCHEDULE")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    xmlElement.setAttribute "STATUS", "10"
    
    'DJP SYS4534 Don't write Now (Date datatype) to an attribute value.
    xmlElement.setAttribute "CREATIONDATETIME", CStr(dtToDay)
    
    xmlElement.setAttribute "EXECUTIONDATETIME", strExecuteBatchDateTime
    
    'BM0438 the totals are now on the BatchScheduleProgress table
    'xmlElement.setAttribute "TOTALNUMBEROFRECORDS", "0"
    'xmlElement.setAttribute "TOTALNUMBEROFFAILURES", "0"
    'xmlElement.setAttribute "TOTALNUMBEROFSUCCESSES", "0"
    'BM0438 End
    
    'SYS4368 Changed FAILEDBATCHNUMBER to FAILEDBATCHRUNNUMBER
    'xmlElement.setAttribute "FAILEDBATCHNUMBER", strFailedBatchNumber
    xmlElement.setAttribute "FAILEDBATCHRUNNUMBER", strFailedBatchRunNumber
    
    xmlTempElement.appendChild xmlElement
    
    'Create TempResponseElement
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempDoc.appendChild xmlTempResponseElem
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Call CreateBatchSchedule
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    objDO.CreateBatchSchedule xmlTempElement, xmlTempResponseElem
    errCheckXMLResponse xmlTempResponseElem.xml, True
    'strBatchRunNumber = xmlGetAttributeText(xmlTempResponseElem, "BATCHRUNNUMBER")
        
    'BM0438 Create the corresponding BatchScheduleProgress entry
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    Set xmlElement = xmlTempDoc.createElement("BATCHSCHEDULEPROGRESS")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    'xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    xmlElement.setAttribute "TOTALNUMBEROFRECORDS", "0"
    xmlElement.setAttribute "TOTALNUMBEROFFAILURES", "0"
    xmlElement.setAttribute "TOTALNUMBEROFSUCCESSES", "0"
    xmlTempElement.appendChild xmlElement
    
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    
    objDO.CreateBatchScheduleProgress xmlTempElement, xmlTempResponseElem
    errCheckXMLResponse xmlTempResponseElem.xml, True
    'BM0438 End
    
    Set xmlTempResponseElem = Nothing
    'Set xmlTempDoc = Nothing
    
'Stage3:

    'Get the MASTERQUEUENAME GlobalParameter Value from GlobalParameter DB Table
    strMasterQueueName = GetGlobalParamString("MasterQueueName")
    
'Stage4:

    'Send to Queue
    
    'Preparing Request Element
    Set xmlElement = xmlTempDoc.createElement("REQUEST")
    
    Set xmlElement1 = xmlTempDoc.createElement("MESSAGEQUEUE")
    xmlElement1.setAttribute "QUEUENAME", strMasterQueueName
    xmlElement1.setAttribute "PROGID", "omBatch.BatchScheduleBO"
    
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    xmlTempElement.Attributes.getNamedItem("OPERATION").Text = "RUNBATCH"
    Set xmlTempNode = vxmlRequestParentNode.selectSingleNode(".//BATCH")
    xmlTempElement.appendChild xmlTempNode
    
    xmlElement1.setAttribute "XML", xmlTempElement.xml
    xmlElement1.setAttribute "EXECUTEAFTERDATE", strExecuteBatchDateTime    'MAR1090 GHun
    xmlElement.appendChild xmlElement1
    
    'Call SendToQueue
    Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
    Set xmlResponse = objMessageQueue.SendToQueue(xmlElement)
    errCheckXMLResponse xmlResponse.xml, True
    
'Stage5:
    
    'Create TempResponseElement
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    'BM0438 Not required
    'xmlTempDoc.appendChild xmlTempResponseElem
    'BM0438 End
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Preparing Request Element
    Set xmlElement = xmlTempDoc.createElement("BATCH")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    xmlElement.setAttribute "STATUS", "20"
    
    'call Update Batch Schedule with the new Status
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    objDO.UpdateBatch xmlElement, xmlTempResponseElem
    errCheckXMLResponse xmlTempResponseElem.xml, True
    
    gobjContext.SetComplete
    
LaunchBatchExit:
    
    Set xmlTempDoc = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlTempElement = Nothing
    'Set xmlNode = Nothing
    Set xmlElement1 = Nothing
    Set objMessageQueue = Nothing
    Set xmlBatchRequestNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlResponse = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
LaunchBatchVBErr:
    gobjContext.SetAbort
    'BM0438 Log errors to the event log
    App.LogEvent "Error " & Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description, vbLogEventTypeInformation
    
    GoTo LaunchBatchExit
End Sub

Private Sub IBatchScheduleTXBO_UpdateBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateBatchVbErr
    
    Const strFunctionName As String = "IBatchScheduleTXBO_UpdateBatch"
    
    Dim xmlUpdateBatchNode As IXMLDOMNode
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    Set xmlUpdateBatchNode = vxmlRequestParentNode.selectSingleNode("BATCH")
        
    objDO.UpdateBatch xmlUpdateBatchNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
UpdateBatchExit:
    
    Set objDO = Nothing
    Set xmlUpdateBatchNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateBatchVbErr:

    gobjContext.SetAbort

    GoTo UpdateBatchExit
End Sub

Private Sub IBatchScheduleTxBO_UpdateBatchAudit(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateBatchAuditVbErr
    
    Const strFunctionName As String = "IBatchScheduleTxBO_UpdateBatchAudit"
    
    Dim xmlUpdateBatchAuditNode As IXMLDOMNode
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    Set xmlUpdateBatchAuditNode = vxmlRequestParentNode.selectSingleNode("BATCHAUDIT")
        
    objDO.UpdateBatchAudit xmlUpdateBatchAuditNode, vxmlResponseNode
    
    gobjContext.SetComplete
    
UpdateBatchAuditExit:
    
    Set objDO = Nothing
    Set xmlUpdateBatchAuditNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateBatchAuditVbErr:

    gobjContext.SetAbort

    GoTo UpdateBatchAuditExit
End Sub

Private Function IBatchScheduleTxBO_CreateBatchAudit(ByVal vxmlRequestParentNode As IXMLDOMNode) As IXMLDOMNode
    
    'Error Handler
    On Error GoTo CreateBatchAuditVbErr
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleTxBO_CreateBatchAudit"
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    Set IBatchScheduleTxBO_CreateBatchAudit = objDO.CreateBatchAudit(vxmlRequestParentNode)
    
    gobjContext.SetComplete
    
CreateBatchAuditExit:
    
    Set objDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Function
    
CreateBatchAuditVbErr:

    gobjContext.SetAbort
    
    GoTo CreateBatchAuditExit
End Function

Private Sub IBatchScheduleTxBO_UpdateBatchSchedule(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Const strFunctionName As String = "IBatchScheduleTxBO_UpdateBatchSchedule"
    Dim xmlBatchSchedule As IXMLDOMElement
    'BM0438
    Const cstrTNORecords         As String = "TOTALNUMBEROFRECORDS"
    Const cstrTNOSuccesses       As String = "TOTALNUMBEROFSUCCESSES"
    Const cstrTNOFailures        As String = "TOTALNUMBEROFFAILURES"
    Dim xmlBatchScheduleProgress As IXMLDOMElement
    Dim blnHasTNORecords         As Boolean
    Dim blnHasTNOSuccesses       As Boolean
    Dim blnHasTNOFailures        As Boolean
    'BM0438 End
    
On Error GoTo UpdateBatchScheduleVbErr
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & "." & "BatchScheduleDO")
    End If
    
    Set xmlBatchSchedule = vxmlRequestParentNode.selectSingleNode("BATCHSCHEDULE")
    
    'BM0438 updates to totals must be made on the BatchScheduleProgress table
    blnHasTNORecords = Not IsNull(xmlBatchSchedule.getAttribute(cstrTNORecords))
    blnHasTNOSuccesses = Not IsNull(xmlBatchSchedule.getAttribute(cstrTNOSuccesses))
    blnHasTNOFailures = Not IsNull(xmlBatchSchedule.getAttribute(cstrTNOFailures))
    
    'Check if any TotalNumberOf* attributes are in the request
    If (blnHasTNORecords Or blnHasTNOSuccesses Or blnHasTNOFailures) Then
        'Create a new BatchScheduleProgress Node
        Set xmlBatchScheduleProgress = vxmlRequestParentNode.ownerDocument.createElement("BATCHSCHEDULEPROGRESS")
        'Copy the BatchNumber and BatchRunNumber to the new request
        xmlBatchScheduleProgress.setAttribute "BATCHNUMBER", xmlBatchSchedule.getAttribute("BATCHNUMBER")
        xmlBatchScheduleProgress.setAttribute "BATCHRUNNUMBER", xmlBatchSchedule.getAttribute("BATCHRUNNUMBER")
        
        'Move the TotalNumberOf* attributes to the BatchScheduleProgress request
        If blnHasTNORecords Then
            xmlBatchScheduleProgress.setAttribute cstrTNORecords, xmlBatchSchedule.getAttribute(cstrTNORecords)
            xmlBatchSchedule.removeAttribute cstrTNORecords
        End If
        If blnHasTNOSuccesses Then
            xmlBatchScheduleProgress.setAttribute cstrTNOSuccesses, xmlBatchSchedule.getAttribute(cstrTNOSuccesses)
            xmlBatchSchedule.removeAttribute cstrTNOSuccesses
        End If
        If blnHasTNOFailures Then
            xmlBatchScheduleProgress.setAttribute cstrTNOFailures, xmlBatchSchedule.getAttribute(cstrTNOFailures)
            xmlBatchSchedule.removeAttribute cstrTNOFailures
        End If
        
        objDO.UpdateBatchScheduleProgress xmlBatchScheduleProgress, vxmlResponseNode
    End If
    
    'If there are any attributes left, besides the 2 that make up the key, then update
    If xmlBatchSchedule.Attributes.length > 2 Then
        objDO.UpdateBatchSchedule xmlBatchSchedule, vxmlResponseNode
    End If
    'BM0438 End
    
    gobjContext.SetComplete
    
UpdateBatchScheduleExit:
    
    Set objDO = Nothing
    Set xmlBatchSchedule = Nothing
    Set xmlBatchScheduleProgress = Nothing 'BM0438
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
UpdateBatchScheduleVbErr:
    gobjContext.SetAbort
    GoTo UpdateBatchScheduleExit
End Sub

Private Sub IBatchScheduleTXBO_CreateBatchException(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo CreateBatchExceptionVbErr
    Const strFunctionName As String = "IBatchScheduleTXBO_CreateBatchException"
    
    If objDO Is Nothing Then
        Set objDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
  
    objDO.CreateBatchException vxmlRequestParentNode, vxmlResponseNode
  
  
    gobjContext.SetComplete
CreateBatchExceptionExit:
    
    Set objDO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
Exit Sub
CreateBatchExceptionVbErr:
    ' Not a fatal error to fail to write to Batch Exception.
    GoTo CreateBatchExceptionExit
    
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

