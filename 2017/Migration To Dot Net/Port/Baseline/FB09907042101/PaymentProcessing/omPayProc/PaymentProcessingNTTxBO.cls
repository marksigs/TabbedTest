VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "PaymentProcessingNTTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      PaymentProcessingNTTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   PaymentProcessing Transactional Business Object.
'               Starts a New Transaction

'Dependencies:
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MC     06/12/01    SYS3018 Fixes to Batch Process for SQL Server.
'MC     07/12/01    SYS3440 Various Fixes to Batch Process.
'DM     07/01/02    SYS3185 Implement CompleteInterfacing
'RF     25/02/02    SYS4161 Incorrect validation in InterfacePaidPayment
'RF     26/02/02    SYS3555 Update ApplicationStage in InterfacePaidPayment
'RF     27/02/02    SYS4193 Add new workaround method TemporaryCompleteInterfacing
'DM     15/03/02    SYS4274 CompleteInterfacing changes.
'DM     27/03/02    SYS4313 CompleteInterfacing changes
'DM     09/04/02    SYS4377 Testing for nothing from the CBA
'DM     11/04/02    SYS4385 Create the request correctly
'JLD    22/04/02    SYS3552 send all payment types to be interfaces except ROF and Cancel Balance
'SA     21/05/02    SYS4678 Completions processing - Doretry method losing BatchSchedule node.
'SA     22/05/02    SYS4515 CompleteInterface method looking for status of "ERROR" instead of "FAILURE" so setting payment status incorrectly
'SG     28/05/02    SYS4767 MSMS to Core integration
'SA     31/05/02    MSMS0085 Changes to HexagonCreatePaymentRecord & HexagonProcessNewPayment
'STB    07/06/02    SYS4747 DoRetry functions correctly and raises an error after the maximum # retries.
'STB    07/06/02    SYS4076 Allow applications to be locked by the current (batch) UserID.
'STB    10/06/02    SYS4674 Fixes for multi-loan components.
'STB    14/06/02    SYS4894 Do not attempt to lock the application if it is already locked.
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        AQR         Description
'PSC    14/11/2005  MAR520      Amend CompleteInterfacing to store returned values
'PSC    25/11/2005  MAR140      Add GetPostCompletionCaseTasksList
'GHun   20/04/2006  MAR1651     Commented out obsolete function SetMinimumApplicationStage
'JD     19/05/2006  MAR1398     InterfacePaidPayment - check global parameter PPCalculateAtCompletion before calling alpha+ calcs
'AW     23/08/2006  EP977       InterfacePaidPayment - Unlock applicatiions here because CompleteInterfacing no longer called
'                                                      in Epsom Phase 1.
'--------------------------------------------------------------------------------------------------------------
'Epsom History:
'
'Prog   Date        Description
'PSC    13/03/2007  EP2_1347 Treat valuation refunds the same as incentives
'--------------------------------------------------------------------------------------------------------------

Option Explicit

Implements IPaymentProcessingNTTxBO
Implements ObjectControl

Private gobjContext As ObjectContext

Private Enum LOCKSTATUS
    lsLocked
    lsUnlocked
    lsError
End Enum

'SG 28/05/02 SYS4767 START
'CHAPS file formatting codes.
Private Const CHAPS_HEADER_PREFIX   As String = "H"
Private Const CHAPS_TRAILER_PREFIX  As String = "T"
Private Const CHAPS_PAYMENT_PREFIX  As String = "P"
Private Const CHAPS_SEP             As String = "="
Private Const CHAPS_RECORD_SUFFIX   As String = "#"

'CHAPS Format masks.
Private Const CHAPS_TEXT_FORMAT  As String = "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"

'CHAPS record lengths (including CrLf).
Private Const CHAPS_HEADER_LENGTH   As Long = 13
'BM0173 MDC 10/12/2002
'Private Const CHAPS_TRAILER_LENGTH  As Long = 21
'Private Const CHAPS_PAYMENT_LENGTH  As Long = 973
Private Const CHAPS_TRAILER_LENGTH  As Long = 24
Private Const CHAPS_PAYMENT_LENGTH  As Long = 976
'BM0173 MDC 10/12/2002 - End

'Hexagon record lengths (including CrLf).
Private Const HEXAGON_HEADER_LENGTH  As Long = 37
Private Const HEXAGON_PAYMENT_LENGTH As Long = 112

'SG 20/05/02 MSMS0083
Private mstrCHAPSPaymentBody As String
Private mstrHexagonPaymentBody As String
'SG 28/05/02 SYS4767 END

'BM0339 MDC 05/03/2003
Private Enum CompleteInterfacingStages
    stgGeneral
    stgCreateTask
    stgPrintLetter
    stgUpdatePayment
    stgUpdateFee
    stgUnlock
End Enum
'BM0339 MDC 05/03/2003 - End

Private Sub IPaymentProcessingNTTxBO_CompleteInterfacing( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'DM     07/01/02    SYS3185 - Implement CompleteInterfacing
'DM     27/03/02    SYS4313 CompleteInterfacing changes
'------------------------------------------------------------------------------------------
On Error GoTo CompleteInterfacingVbErr

    Const strFunctionName As String = "IPaymentProcessingNTTxBO_CompleteInterfacing"

    ' PSC 20/09/2002 BMIDS00480
    'Dim objCBRequestNode As IXMLDOMNode
    Dim objSystemStatusNode As IXMLDOMNode
    ' PSC 20/09/2002 BMIDS00480
    'Dim objResultStatus As IXMLDOMNode
    Dim objRespXML As IXMLDOMNode ' An XML response node
    ' PSC 20/09/2002 BMIDS00480 - Start
    'Dim objApplicationManagerBO As ApplicationManagerBO
    'Dim strRespFromUnLock As String
    'Dim objXMLRequest As IXMLDOMNode
    'Dim objApplicationNumberNode As IXMLDOMNode
    ' PSC 20/09/2002 BMIDS00480 - End
    Dim objCBRequestHeaderNode As IXMLDOMNode
    ' PSC 20/09/2002 BMIDS00480
    'Dim objMessageQueueListener As Object ' Late bound because we have not got access to the type library

    Dim strApplicationNumber As String
    Dim strResponse As String
    Dim blnWaitingResponse As Boolean   ' PSC 20/09/2002 BMIDS00480
    
    'BM0339 MDC 05/03/2003
    Dim strMessage As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strPaymentSequenceNumber As String
    Dim intStage As CompleteInterfacingStages
    
    intStage = stgGeneral
    'BM0339 MDC 05/03/2003 - End
    
    ' DM 15/03/02 SYS4274
    ' PSC 20/09/2002 BMIDS00480 - Start
    'Set objCBRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST")
    Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "HEADER")
    ' PSC 20/09/2002 BMIDS00480 - End
    
    ' Needed in a number of places so get it now
    strApplicationNumber = xmlGetAttributeText(objCBRequestHeaderNode, "APPLICATIONNUMBER")
    
    ' PSC 20/09/2002 BMIDS00480
    ' UpdateAdminSubLedgerCodes vxmlRequestNode
    
    ' PSC 19/09/2002 SYS4184 - Uncomment code
    If GetGlobalParamBoolean("GenerateCompletionsStatusTask") Then
        intStage = stgCreateTask    'BM0339 MDC 05/03/2003
        
        ' This needs to be fixed.
        CreateCompletionStatusTask vxmlRequestNode
    
        intStage = stgGeneral    'BM0339 MDC 05/03/2003
    
    End If
    
    ' Check the transaction within the Admin System was handled successfully technically
    ' PSC 20/09/2002 BMIDS00480 - Start
    Set objSystemStatusNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/RESPONSE")
    If xmlGetAttributeText(objSystemStatusNode, "TYPE") = "SUCCESS" Then 'No technical error
    
        ' Log all business errors including warnings
        ' DM 15/03/02 SYS4274
'        Set objResultStatus = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/STATUS")
'        If xmlGetAttributeText(objResultStatus, "VALUE") <> "SUCCESS" Then
'            ' PSC 19/09/2002 SYS4863 Add call to UpdateFeeCompletionIndicators
'            UpdateFeeCompletionIndicators vxmlRequestNode, "E"
'            LogBusinessExceptions vxmlRequestNode
'
'        End If
'        'SYS4515 Need to check for "FAILURE" not "ERROR"
'        'If xmlGetAttributeText(objResultStatus, "VALUE") <> "ERROR" Then
'        If xmlGetAttributeText(objResultStatus, "VALUE") <> "FAILURE" Then
         ' No fatal business error exists
         
        intStage = stgPrintLetter    'BM0339 MDC 05/03/2003
         
        PrintCompletionLetter vxmlRequestNode, objRespXML
         
        intStage = stgUpdatePayment    'BM0339 MDC 05/03/2003
        
        ' PSC 14/11/2005 MAR520
        UpdateAccountNo vxmlRequestNode
         
         ' Update payment details
         UpdateDisbursementStatus vxmlRequestNode, "I"
         
        intStage = stgUpdateFee    'BM0339 MDC 05/03/2003
         
         ' PSC 19/09/2002 SYS4863 Add call to UpdateFeeCompletionIndicators
         UpdateFeeCompletionIndicators vxmlRequestNode, "S"

    Else
        intStage = stgUpdatePayment    'BM0339 MDC 05/03/2003
            
            ' One or more fatal business errors have been found within the CBA
        UpdateDisbursementStatus vxmlRequestNode, "IF"
        
        intStage = stgUpdateFee    'BM0339 MDC 05/03/2003
        
        ' PSC 19/09/2002 SYS4863 Add call to UpdateFeeCompletionIndicators
        UpdateFeeCompletionIndicators vxmlRequestNode, "E"

    End If
    
    'BM0339 MDC 05/03/2003 - Move unlocking to a separate function
    ' PSC 20/09/2002 BMIDS00480 - Start
    ' Check there are no disbursements waiting for a reply before unlocking
    ' the application
'    blnWaitingResponse = IsADisbursementAwaitingResponse(strApplicationNumber)
'
'    If blnWaitingResponse = False Then
'        UnlockApplicationOnly strApplicationNumber, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")   'BM0339 MDC 06/03/2003
'    End If
    ' PSC 20/09/2002 BMIDS00480 - End
    'BM0339 MDC 05/03/2003 - End
        
    intStage = stgGeneral    'BM0339 MDC 05/03/2003
        
'    Else
'        ' A technical error occured within the CBA
'
'        UpdateDisbursementStatus vxmlRequestNode, "IF"
'
'        ' PSC 19/09/2002 SYS4863 Add call to UpdateFeeCompletionIndicators
'        UpdateFeeCompletionIndicators vxmlRequestNode, "E"
'
'        ' Need to stop the Queue into completeinterfacing
'        Dim strAdminSystemQueueName As String
'        Dim intMessageQueueType As Integer
'
'        strAdminSystemQueueName = GetGlobalParamString("AdminSystemQueueName")
'
'        Set objXMLRequest = Nothing
'        ' Need to create our own request node because MQL does not conform to Omiga standards.
'        Set objXMLRequest = vxmlRequestNode.ownerDocument.createElement("REQUEST")
'
'        Dim xmlQueuelistNode As IXMLDOMNode
'        Dim xmlQueueNode As IXMLDOMNode
'        Dim xmlQueueName As IXMLDOMNode
'        Dim xmlQueueTask As IXMLDOMNode
'        Dim xmlQueueType As IXMLDOMNode
'
'        xmlSetAttributeValue objXMLRequest, "ACTION", "UPDATE"
'        Set xmlQueuelistNode = vxmlRequestNode.ownerDocument.createElement("QUEUELIST")
'        Set xmlQueueNode = vxmlRequestNode.ownerDocument.createElement("QUEUE")
'        Set xmlQueueName = vxmlRequestNode.ownerDocument.createElement("NAME")
'        Set xmlQueueTask = vxmlRequestNode.ownerDocument.createElement("TASK")
'        Set xmlQueueType = vxmlRequestNode.ownerDocument.createElement("TYPE")
'
'        objXMLRequest.appendChild xmlQueuelistNode
'        xmlQueuelistNode.appendChild xmlQueueNode
'
'        xmlQueueNode.appendChild xmlQueueName
'        xmlQueueNode.appendChild xmlQueueType
'        xmlQueueNode.appendChild xmlQueueTask
'
'        xmlQueueName.Text = strAdminSystemQueueName
'        xmlQueueTask.Text = "STOP"
'        xmlQueueType.Text = "OMMQ1" ' Needs to be retrieved from a parameter. Or get MQL fixed
'                                    ' so this is not a required field.
'
'        Set objMessageQueueListener = CreateObject("MessageQueueListener.MessageQueueListener1.1")
'
'        strResponse = objMessageQueueListener.Configure(objXMLRequest.xml)
'        errCheckXMLResponse strResponse, True
'
'    End If
    ' PSC 20/09/2002 BMIDS00480 - End
    
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
   
    gobjContext.SetComplete

CompleteInterfacing_Exit:

    ' PSC 20/09/2002 BMIDS00480
    'Set objCBRequestNode = Nothing
    Set objRespXML = Nothing

    Set objSystemStatusNode = Nothing
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    'Set objResultStatus = Nothing
    'Set objApplicationManagerBO = Nothing
    'Set objXMLRequest = Nothing
    'Set objApplicationNumberNode = Nothing
    ' PSC 20/09/2002 BMIDS00480 - End
    Set objCBRequestHeaderNode = Nothing
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    'Set xmlQueuelistNode = Nothing
    'Set xmlQueueNode = Nothing
    'Set xmlQueueName = Nothing
    'Set xmlQueueTask = Nothing
    'Set xmlQueueType = Nothing
    'Set objMessageQueueListener = Nothing
    ' PSC 20/09/2002 BMIDS00480 - End

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub

CompleteInterfacingVbErr:
    
    
    'BM0339 MDC 05/03/2003 - Log errors instead of raising them
    If Not objCBRequestHeaderNode Is Nothing Then
        strBatchNumber = xmlGetAttributeText(objCBRequestHeaderNode, "BATCHNUMBER")
        strBatchRunNumber = xmlGetAttributeText(objCBRequestHeaderNode, "BATCHRUNNUMBER")
        strPaymentSequenceNumber = xmlGetAttributeText(objCBRequestHeaderNode, "PAYMENTSEQUENCENUMBER")
    End If
    strMessage = "Warning: An error occurred in " & strFunctionName
    Select Case intStage
    Case stgCreateTask
        strMessage = strMessage & " while creating Completion Status Task"
    Case stgPrintLetter
        strMessage = strMessage & " while printing completions letter"
    Case stgUpdatePayment
        strMessage = strMessage & " while updating payment status"
    Case stgUpdateFee
        strMessage = strMessage & " while updating fee completion indicators"
    Case stgUnlock
        strMessage = strMessage & " while unlocking the application"
    Case Else
    End Select
    
    strMessage = strMessage & vbCrLf & "Application: " & strApplicationNumber
    strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & strPaymentSequenceNumber
    strMessage = strMessage & vbCrLf & "Number: " & Err.Number
    strMessage = strMessage & vbCrLf & "Description: " & Err.Description
    strMessage = strMessage & vbCrLf & "Source: " & Err.Source
        
    Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
    
    'Clear the error and return success
    Err.Clear
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    'BM0339 MDC 05/03/2003 - End
    
    gobjContext.SetAbort
    GoTo CompleteInterfacing_Exit

End Sub

'BMIDS628 INR Moved this to om4PayProcGlobals to allow it to be used elsewhere
'BM0339 MDC 05/03/2003
'Private Sub LogWarningMessage(ByVal strBatchNumber As String, ByVal strBatchRunNumber As String, _
'                                                                    ByVal strMessage As String)
'
'On Error GoTo LogWarningMessageExit
'
'Dim strBatchLogFilePath As String
'Dim intFileNo As Integer
'Dim blnFileOpen As Boolean
'Dim strBatchLogFile As String
'
'    'Write an warning entry in the event log ====================================================
'    App.LogEvent vbCrLf & strMessage, vbLogEventTypeWarning
'
'
'    'Append an item to the batch run warnings log ===============================================
'    strBatchLogFilePath = GetGlobalParamString("BatchLogFilePath")
'
'    'Checking for the Directory Existence to write the outputFile
'    'If does not Exist then Create One
'    If Dir(strBatchLogFilePath, vbDirectory) = "" Then
'        MkDir strBatchLogFilePath
'    End If
'
'    strBatchLogFile = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber & "_Warnings"
'
'    intFileNo = FreeFile
'    Open strBatchLogFile & ".log" For Append As #intFileNo
'    blnFileOpen = True
'
'    Print #intFileNo, ""
'    Print #intFileNo, strMessage
'    Print #intFileNo, "------------------------------------------------------"
'
'LogWarningMessageExit:
'    If blnFileOpen Then
'        Close #intFileNo
'    End If
'
'End Sub
'BM0339 MDC 05/03/2003 - End

Private Sub IPaymentProcessingNTTxBO_GetLockData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:Returns the lock information for an application.
' pass:Application number, User id and unit
' return:   XML containing the lock information or an error indicating record not found.
'           Will raise an error for errors other than record not found.
' history:
' DM This method is implemented so we can get the Lock data for an Application
' in its own transaction. Otherwise we get a blocking situation.
'------------------------------------------------------------------------------------------
On Error GoTo GetLockDataVbErr
    Const strFunctionName As String = "IPaymentProcessingNTTxBO_GetLockData"

    Dim objAppManagerBO As ApplicationManagerBO
    Dim nodeNewRequest As IXMLDOMNode
    Dim elemTemp As IXMLDOMElement
    Dim nodeAppLock As IXMLDOMNode
    Dim nodeCBRequest As IXMLDOMNode
    Dim nodeCBRequestHeader As IXMLDOMNode
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim strApplicationNumber As String
    Dim strResponse As String
    
    ' Need to get the application number
    ' PSC 20/09/2002 BMIDS00480
    Set nodeCBRequest = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST")
    Set nodeCBRequestHeader = xmlGetMandatoryNode(nodeCBRequest, "HEADER")
    strApplicationNumber = xmlGetAttributeText(nodeCBRequestHeader, "APPLICATIONNUMBER")

    Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set elemTemp = nodeNewRequest.ownerDocument.createElement("APPLICATIONLOCK")
    Set nodeAppLock = nodeNewRequest.appendChild(elemTemp)
    Set elemTemp = nodeNewRequest.ownerDocument.createElement("APPLICATIONNUMBER")
    elemTemp.Text = strApplicationNumber
    nodeAppLock.appendChild elemTemp
         
    ' call omApp
    Set objAppManagerBO = gobjContext.CreateInstance("omApp.ApplicationManagerBO")
    
    strResponse = objAppManagerBO.GetLockData(nodeNewRequest.xml)
    
    Set xmlResponse = xmlLoad(strResponse, strFunctionName)
    
    xmlCopyAttribute xmlResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, "TYPE"
    
    Dim xmlChildNode As IXMLDOMNode
    For Each xmlChildNode In xmlResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next

    gobjContext.SetComplete

GetLockData_Exit:

    Set objAppManagerBO = Nothing
    Set nodeNewRequest = Nothing
    Set elemTemp = Nothing
    Set nodeAppLock = Nothing
    Set nodeCBRequest = Nothing
    Set nodeCBRequestHeader = Nothing
    Set xmlResponse = Nothing
    Set xmlChildNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub

GetLockDataVbErr:
    
    gobjContext.SetAbort
   
    GoTo GetLockData_Exit

End Sub

' PSC 25/11/2005 MAR140 - Start
Private Function IPaymentProcessingNTTxBO_GetPostCompletionCaseTasksList(ByVal vxmlRequest As String) As String
    On Error GoTo GetPostCompletionCaseTasksListVbErr
    
    Dim objPayProcBO As PaymentProcessingBO
    
    Set objPayProcBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")
    IPaymentProcessingNTTxBO_GetPostCompletionCaseTasksList = objPayProcBO.omPayProcRequest(vxmlRequest)

GetPostCompletionCaseTasksListExit:
    
    Set objPayProcBO = Nothing
    Exit Function
    
GetPostCompletionCaseTasksListVbErr:
    
    IPaymentProcessingNTTxBO_GetPostCompletionCaseTasksList = errCreateErrorResponse()
    GoTo GetPostCompletionCaseTasksListExit
  
End Function
' PSC 25/11/2005 MAR140 - End

Private Sub IPaymentProcessingNTTxBO_InterfacePaidPayment( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     26/02/02    SYS3555 Update ApplicationStage in InterfacePaidPayment.
'RF     27/02/02    SYS4193 Add new workaround method TemporaryCompleteInterfacing.
'------------------------------------------------------------------------------------------
On Error GoTo InterfacePaidPaymentVbErr

    Const strFunctionName As String = "IPaymentProcessingNTTxBO_InterfacePaidPayment"
    
    Dim objPayProcBO As PaymentProcessingBO
    Dim objMessageQueue As IomMessageQueue
    Dim objAdmin As AdminInterfaceBO
    Dim objBatchSched As BatchScheduleBO
    
    ' PSC 20/09/2002 BMIDS00480
    Dim objPayProcDO As IPAymentProcessingDO
    
    ' PSC 25/11/2005 MAR140
    Dim objPayProcNTTxBO As IPaymentProcessingNTTxBO
    
    Dim xmlBatchSched As IXMLDOMNode
    Dim xmlBatchAudit As IXMLDOMNode
    Dim xmlPaymentRequest As IXMLDOMNode
    Dim xmlPaymentRecord As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode
    Dim xmlQueueRequest As IXMLDOMNode
    Dim xmlQueueNode As IXMLDOMNode
    Dim nodeTemp As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeDisb As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode              ' PSC 20/09/2002 BMIDS00480
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strValue As String
    Dim strValueId As String
    Dim blnContinue As Boolean              ' PSC 20/09/2002 BMIDS00480
    
    Dim strResponse As String
    Dim strBatchNumber As String
    Dim strBatchAuditGuid As String
    Dim strQueueName As String
    Dim lngValue As Long
    Dim intLockStatus As LOCKSTATUS
    
    
    'SYS4076 - Applications can be already locked by the current batch user.
    Dim sUserId As String
    Dim sUnitID As String
    Dim sLockUserID As String
    Dim sLockUnitID As String
    Dim xmlAppLock As IXMLDOMNode
    
    ' PSC 15/08/2002 BMIDS00330
    Dim xmlRequestNode As IXMLDOMNode
        
    'BM0339 MDC 06/03/2003
    Dim strBatchRunNumber As String
    Dim strMessage As String
    'BM0339 MDC 06/03/2003 - End
    
    ' PSC 13/03/2007 EP2_1347 - Start
    Dim blnIsReturnOfFunds As Boolean
    Dim blnIsCancelBalance As Boolean
    Dim blnIsIncentive As Boolean
    Dim blnIsValuationRefund As Boolean
    Dim blnIsInitialAdvance As Boolean
    ' PSC 13/03/2007 EP2_1347 - End
    
    'Get Request values
    'SYS4747 - A copy/clone of the original request nodes are taken as both IsBatchCancelled and
    'the call into GetInterfacingDetails seems to be clearing the BATCHSCHEDULE and PAYMENTRECORD
    'elements.
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE").cloneNode(True)
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    'BMIDS628
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    
    Set xmlBatchAudit = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT").cloneNode(True)
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlBatchAudit, "BATCHAUDITGUID")
    
    Set xmlPaymentRequest = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD").cloneNode(True)
    strAppNo = xmlGetMandatoryAttributeText(xmlPaymentRequest, "APPLICATIONNUMBER")
    xmlCheckMandatoryAttribute xmlPaymentRequest, "PAYMENTSEQUENCENUMBER"
    'SYS4747 - End.
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'SYS4076 - the DOM should be set to synchronous operations.
    xmlDoc.async = False
    
    If Not IsBatchCancelled(xmlBatchSched) Then
    
        ' Call GetInterfacingDetails to get the Payment Details
        Set objPayProcBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetInterfacingDetails"
        xmlTempRequest.appendChild xmlPaymentRequest
        Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
        strResponse = objPayProcBO.omPayProcRequest(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
        xmlDoc.loadXML strResponse
        Set xmlTempResponse = xmlDoc.documentElement
        Set xmlPaymentRecord = xmlGetMandatoryNode(xmlTempResponse, "PAYMENTRECORD")
        
        intLockStatus = GetApplicationLockStatus(strAppNo, strResponse)
        
        'SYS4674 - Application may not have a lock element.
        'SYS4076 - If the UserID and the UnitID on the lock are the same, allow batch to continue.
        'Read the lock information into a DOM document and get the application lock information.
        xmlDoc.loadXML strResponse
        Set xmlAppLock = xmlGetNode(xmlDoc, "//APPLICATIONLOCK")
        
        'Get the batch requesters User and UnitID if available.
        sUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
        sUnitID = xmlGetAttributeText(vxmlRequestNode, "UNITID")
        
        'Get the 'lockers' User and UnitID if available.
        If Not xmlAppLock Is Nothing Then
            sLockUserID = xmlGetNodeText(xmlAppLock, "USERID")
            sLockUnitID = xmlGetNodeText(xmlAppLock, "UNITID")
        End If
        'SYS4674 - End
        
        'If the application is unlocked, or is locked by the same user.
        If (intLockStatus = lsUnlocked) Or ((intLockStatus = lsLocked And sUserId = sLockUserID And sUnitID = sLockUnitID)) Then
            'No applicable lock found, so lock application/customers
            strAppFFNo = GetFactFindNumberForApplication(strAppNo)
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            
            'SYS4894 - Do not attempt to lock the application if it is already locked
            If intLockStatus = lsUnlocked Then
                'BM0339 MDC 05/03/2003
                'Call LockApplication(xmlTempRequest, strAppNo, strAppFFNo)
                Call LockApplicationOnly(xmlTempRequest, strAppNo, strAppFFNo)
                'BM0339 MDC 05/03/2003 - End
            End If
        
            'Validate that Payment Method meets Batch criteria
            If ValidatePayment(xmlPaymentRecord, strBatchNumber, strBatchRunNumber, "P") Then
            
                'Validate Payment Type is 'Initial Advance'
                Set nodeDisb = xmlGetMandatoryNode(xmlPaymentRecord, "DISBURSEMENTPAYMENT")
                                
                strValue = xmlGetAttributeText(nodeDisb, "PAYMENTTYPE")
                
                ' PSC 13/03/2007 EP2_1347 - Start
                blnIsReturnOfFunds = IsValidationType("PaymentType", CSafeInt(strValue), "N")
                blnIsCancelBalance = IsValidationType("PaymentType", CSafeInt(strValue), "NCB")
                blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strValue), "C")
                blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strValue), "VALREFUND")
                blnIsInitialAdvance = IsValidationType("PaymentType", CSafeInt(strValue), "I")
                ' PSC 13/03/2007 EP2_1347 - End
                
                ' PSC 20/09/2002 BMIDS00480 - Allow ROF to be processed
                ' If strValue <> strROFValueId And strValue <> strCBValueId Then
                ' PSC 13/03/2007 EP2_1347
                If blnIsCancelBalance = False Then
                
                    ' PSC 20/09/2002 BMIDS00480 - Start
                    '------------------------------------------------------------------------------------------
                    ' Check that if this is not an initial advance that the application
                    ' already has an interfaced initial advance
                    '------------------------------------------------------------------------------------------
                    
                    Dim lngErrNo As Long
                    Dim strDescription As String
                    Dim strSource As String
                    
                    blnContinue = True
                    
                    ' PSC 13/03/2007 EP2_1347
                    If blnIsInitialAdvance = False Then
                    
                         ' Create request to see if interfaced initial advances exist
                         Set nodeTemp = vxmlRequestNode.ownerDocument.createElement("INTERFACEDINITIALADVANCE")
                         xmlSetAttributeValue nodeTemp, "APPLICATIONNUMBER", strAppNo
                    
                         Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
                     
                         On Error Resume Next
                         
                         Set xmlNode = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
                         
                         objPayProcDO.FindInterfacedInitialAdvances nodeTemp, xmlNode
                         
                         lngErrNo = Err.Number
                         strDescription = Err.Description
                         strSource = Err.Source
                         
                         On Error GoTo InterfacePaidPaymentVbErr
                         
                         ' If no initial advances found retry
                         If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
                            If DoRetry(vxmlRequestNode, 7007) Then
                                'Indicate that retrying so batch schedule success/failure not updated at this time
                                xmlSetAttributeValue vxmlResponseNode, "TYPE", "RETRY"
                                blnContinue = False
                            End If
                         ElseIf lngErrNo <> 0 Then
                             Err.Raise lngErrNo, strSource, strDescription
                         End If
                    End If
                    ' PSC 20/09/2002 BMIDS00480 - End
                    
                    '------------------------------------------------------------------------------------------
                    ' Validate Application Stage at least 'Complete'
                    '------------------------------------------------------------------------------------------
                    
                    If blnContinue = True Then  ' PSC 20/09/2002 BMIDS00480
                        'RF 26/02/02 SYS3555 Uncommented call
                        ' PSC 19/09/2002 SYS4184 - Start
                        'SetMinimumApplicationStage strAppNo, strAppFFNo, "CP"
                        UpdateApplicationStage vxmlRequestNode.cloneNode(False), strAppNo, strAppFFNo
                        ' PSC 19/09/2002 SYS4184 - End
                        
                        'Calculate initial payment and the monthly repayment and update Loan Component Payment records
                        
                        ' PSC 20/09/2002 BMIDS00480 - Start
                        ' Omly do this if not return of funds
                        ' PSC 15/08/2002 BMIDS00330
                        ' Setup a copy of the request to be used for further calls
                        ' PSC 08/10/2002 BMIDS00595 - Don't calc if incentive release
                        ' PSC 13/03/2007 EP2_1347
                        If blnIsReturnOfFunds = False And blnIsIncentive = False And blnIsValuationRefund = False Then
                            'JD MAR1398 check global parameter
                            If GetGlobalParamBoolean("PPCalculateAtCompletion") = True Then
                                Set xmlRequestNode = vxmlRequestNode.cloneNode(False)
                                Call CalculateMonthlyRepayments(strAppNo, strAppFFNo, xmlRequestNode, xmlPaymentRecord)
                            End If
                        End If
                        ' PSC 20/09/2002 BMIDS00480 - End
                        
                        'Generate Completion file and process
                        '#TASK - Interfacing
                        
                        Set nodeTemp = xmlBatchAudit.cloneNode(True)
                        xmlCopyMandatoryAttribute xmlBatchSched, nodeTemp, "BATCHNUMBER"
                        xmlCopyMandatoryAttribute xmlBatchSched, nodeTemp, "BATCHRUNNUMBER"
                        Set objAdmin = gobjContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
                        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                        xmlSetAttributeValue xmlTempRequest, "OPERATION", "ProcessCompletionsInterface"
                        xmlTempRequest.appendChild xmlPaymentRequest
                        xmlTempRequest.appendChild nodeTemp
                        strResponse = objAdmin.Request(xmlTempRequest.xml)
                        errCheckXMLResponse strResponse, True
                        
                        ' PSC 20/09/2002 BMIDS00480 - Start
                        ' Set payment status to Awaiting Interface Reponse
                        UpdatePaymentStatus xmlPaymentRequest, "R"
                        ' PSC 20/09/2002 BMIDS00480 - End
                        
                        ' PSC 16/11/2005 MAR140 - Start
                        Set objPayProcNTTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingNTTxBO")
                        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPostCompletionCaseTasksList"
                        Set nodeTemp = vxmlRequestNode.ownerDocument.createElement("APPLICATION")
                        xmlTempRequest.appendChild nodeTemp
                        xmlSetAttributeValue nodeTemp, "APPLICATIONNUMBER", strAppNo
                        xmlSetAttributeValue nodeTemp, "APPLICATIONFACTFINDNUMBER", strAppFFNo
                        xmlSetAttributeValue nodeTemp, "STAGEID", GetGlobalParamString("TMCompletionsStageId")
                        xmlCopyAttribute nodeDisb, nodeTemp, "COMPLETIONDATE"
                        
                        objPayProcNTTxBO.GetPostCompletionCaseTasksList xmlTempRequest.xml
                        ' PSC 16/11/2005 MAR140 - End
                        
                        'AW 23/08/2006 EP977
                        'Temporary fix for Phase 1 -  Because CompleteInterfacing is no longer called, apps will remain locked.
                        Call UnlockApplicationOnly(strAppNo, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID"))
                        
                        
                    End If ' PSC 20/09/2002 BMIDS00480
                Else
                    'Return of Funds or Cancel Balance payment type.
                    'Write Reject Report
                    
                    'INR BMIDS628 call LogWarningMessage instead of raising an error
'                    Call WriteRejectReport
'                    errThrowError strFunctionName, oePPBatchNotInitAdvance
                    strMessage = "Batch Transaction Error: InterfacePaidPayment: Payment is not an Initial Advance"
                    strMessage = strMessage & vbCrLf & "Application: " & strAppNo
                    strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                                xmlGetMandatoryAttributeText(xmlPaymentRequest, "PAYMENTSEQUENCENUMBER")
                    strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
                    Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
                    
                    'INR BMIDS628 Indicate that there has been a failure
                    xmlSetAttributeValue vxmlResponseNode, "TYPE", "FAIL"

                End If
                
            Else
                'BM0339 MDC 06/03/2003
                'Write Reject Report
                'Call WriteRejectReport
                'errThrowError strFunctionName, oePPBatchPaymentPayeeInvalid
                'BMIDS628 LogWarningMessage is called from validate function
                'where we have access to the reason for failure
'                strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
'                strMessage = "Warning: InterfacePaidPayment: Payment or Payee is Invalid."
'                strMessage = strMessage & vbCrLf & "Application: " & strAppNo
'                strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
'                            xmlGetMandatoryAttributeText(xmlPaymentRequest, "PAYMENTSEQUENCENUMBER")
'                strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
'                Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
                'BM0339 MDC 06/03/2003 - End
                'INR BMIDS628 Unlock the Application
                Call UnlockApplicationOnly(strAppNo, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID"))
                    
                'INR BMIDS628 Indicate that there has been a failure
                xmlSetAttributeValue vxmlResponseNode, "TYPE", "FAIL"

            End If
                        
        'Else it must be locked by someone else.
        ElseIf intLockStatus <> lsError Then
            'Lock found, so post message back to queue and retry later
            'SYS4678 Add Batchschedule node back on as it's lost it somehow!
            vxmlRequestNode.appendChild xmlBatchSched
            If DoRetry(vxmlRequestNode) Then
                'Indicate that retrying so batch schedule success/failure not updated at this time
                xmlSetAttributeValue vxmlResponseNode, "TYPE", "RETRY"
            End If
            
        Else
            'Unexpected error has occurred
            'Write to reject report
            'INR BMIDS628 call LogWarningMessage instead of raising an error
'            Call WriteRejectReport
'            'Raise error
'            errCheckXMLResponse strResponse, True, vxmlResponseNode
            strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
            strMessage = "Batch Transaction Error: InterfacePaidPayment: Unexpected error has occurred"
            strMessage = strMessage & vbCrLf & "Application: " & strAppNo
            strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                        xmlGetMandatoryAttributeText(xmlPaymentRequest, "PAYMENTSEQUENCENUMBER")
            strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
            Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
                    
            'INR BMIDS628 Indicate that there has been a failure
            xmlSetAttributeValue vxmlResponseNode, "TYPE", "FAIL"

        End If
        'SYS4076 - End.
    Else
        'JD BMIDS604 Batch is cancelled, raise an error to prevent the batch status being set as Complete
        errThrowError strFunctionName, 7025
    End If

    'SYS3018 MDC 05/12/2001
    gobjContext.SetComplete
    'SYS3018 End

InterfacePaidPayment_Exit:

    Set objPayProcBO = Nothing
    Set objMessageQueue = Nothing
    Set xmlBatchSched = Nothing
    Set xmlBatchAudit = Nothing
    Set xmlPaymentRequest = Nothing
    Set xmlPaymentRecord = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlQueueRequest = Nothing
    Set xmlQueueNode = Nothing
    Set nodeDisb = Nothing
    Set nodeTemp = Nothing
    Set xmlDoc = Nothing
    Set objPayProcDO = Nothing      ' PSC 20/09/2002 BMIDS00480
    Set xmlNode = Nothing           ' PSC 20/09/2002 BMIDS00480
    Set objPayProcNTTxBO = Nothing  ' PSC 25/11/2005 MAR140

    'SYS4076 - release application lock node.
    Set xmlAppLock = Nothing
    
    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError strFunctionName, TypeName(Me)
    End If
    
Exit Sub

InterfacePaidPaymentVbErr:
    
    gobjContext.SetAbort
    
    GoTo InterfacePaidPayment_Exit

End Sub

Private Sub TemporaryCompleteInterfacing( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal vstrApplicationNumber As String, _
    ByVal vstrPAYMENTSEQUENCENUMBER As String)
' header ----------------------------------------------------------------------------------
' description:
'   For use until the return journey from the CBA works, after which it should be removed.
'   Responsibilities:
'   1) Update disbursement
'   2) IncrementBatchSchedule
'   3) SetBatchStatus
' pass:
' return:
' history:
'RF     27/02/02    Created for SYS4193
'------------------------------------------------------------------------------------------
On Error GoTo TemporaryCompleteInterfacingVbErr

    Const cstrFunctionName As String = "TemporaryCompleteInterfacing"
    
    '------------------------------------------------------------------------------------------
    ' Update disbursement
    ' This uses code copied and tweaked from the UpdateDisbursement method
    '------------------------------------------------------------------------------------------
        
    Dim nodeRequest As IXMLDOMNode
    Dim elemPaymentRecord As IXMLDOMElement
    Dim elemDisbursementPayment As IXMLDOMElement
    Dim strValueId As String
    Dim objPaymentProcessingBO As PaymentProcessingBO
    Dim strResponse As String
    Dim nodeTemp As IXMLDOMNode
    
    strValueId = GetFirstComboValueId("PaymentStatus", "I")
    
    Set nodeRequest = xmlGetRequestNode(vxmlRequestNode)
    
    xmlSetAttributeValue nodeRequest, "OPERATION", "UPDATEDISBURSEMENT"
    
    Set elemPaymentRecord = vxmlRequestNode.ownerDocument.createElement("PAYMENTRECORD")
    xmlSetAttributeValue elemPaymentRecord, "APPLICATIONNUMBER", vstrApplicationNumber
    xmlSetAttributeValue elemPaymentRecord, "PAYMENTSEQUENCENUMBER", vstrPAYMENTSEQUENCENUMBER
    Set nodeTemp = nodeRequest.appendChild(elemPaymentRecord)
    
    Set elemDisbursementPayment = _
        vxmlRequestNode.ownerDocument.createElement("DISBURSEMENTPAYMENT")
    xmlSetAttributeValue elemDisbursementPayment, "PAYMENTSTATUS", strValueId
    nodeTemp.appendChild elemDisbursementPayment
    
    Set objPaymentProcessingBO = gobjContext.CreateInstance("omPayProc.PaymentProcessingBO")
    
    strResponse = objPaymentProcessingBO.omPayProcRequest(nodeRequest.xml)
    errCheckXMLResponse strResponse, True
        
    gobjContext.SetComplete
    
TemporaryCompleteInterfacing_Exit:
    
    Set nodeRequest = Nothing
    Set elemPaymentRecord = Nothing
    Set elemDisbursementPayment = Nothing
    Set nodeTemp = Nothing
    
    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError cstrFunctionName, TypeName(Me)
    End If

Exit Sub

TemporaryCompleteInterfacingVbErr:
    
    gobjContext.SetAbort
    
    GoTo TemporaryCompleteInterfacing_Exit
    
End Sub


Private Sub IPaymentProcessingNTTxBO_PaySanctionedPayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo PaySanctionedPaymentVbErr

Const strFunctionName As String = "IPaymentProcessingNTTxBO_PaySanctionedPayment"

Dim objPayProcBO As PaymentProcessingBO
Dim objMessageQueue As IomMessageQueue

Dim xmlBatchSched As IXMLDOMNode
Dim xmlBatchAudit As IXMLDOMNode
Dim xmlPayment As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode

'SG 28/05/02 SYS4767 START
'SG 25/03/02 SYS4202
Dim xmlBACSResponse As IXMLDOMNode
'SG 28/05/02 SYS4767 END

Dim xmlQueueRequest As IXMLDOMNode
Dim xmlQueueNode As IXMLDOMNode
Dim xmlDisb As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strAppNo As String
Dim strValue As String
Dim strValueId As String
Dim strResponse As String
Dim strBatchNumber As String
'Dim strRetries As String
Dim lngValue As Long
Dim strQueueName As String
Dim intLockStatus As LOCKSTATUS

' PSC 20/09/2002 BMIDS00480 - Start
Dim strPaymentType As String
Dim strROFValueId As String
' PSC 20/09/2002 BMIDS00480 - End

'SG 28/05/02 SYS4767 START
'STB: MSMS0025 - CHAPS/Hexagon file format.
Dim sCHAPSInterface As String
Dim xmlCHAPSResponse As IXMLDOMNode
Dim objGlobalParam As GlobalParameterDO
'SG 28/05/02 SYS4767 END

    'SYS4076 - Applications can be already locked by the current batch user.
    Dim sUserId As String
    Dim sUnitID As String
    Dim sLockUserID As String
    Dim sLockUnitID As String
    Dim xmlAppLock As IXMLDOMNode
    'BMIDS628 INR 07/03/2003
    Dim strBatchRunNumber As String
    Dim strMessage As String
    Dim strBatchAuditGuid As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'SYS4076 - the DOM should be set to synchronous operations.
    xmlDoc.async = False
    
    'Get Request values
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE").cloneNode(True)
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    
    'BMIDS628 Get BatchAuditGuid & BatchRunNumber for error message
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    Set xmlBatchAudit = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlBatchAudit, "BATCHAUDITGUID")
'    strRetries = xmlGetMandatoryAttributeText(xmlBatchAudit, "NUMBEROFRETRIES")
    
    If Not IsBatchCancelled(xmlBatchSched) Then
        'Get the Payment Details
        Set objPayProcBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")
        Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetInterfacingDetails"
        strResponse = objPayProcBO.omPayProcRequest(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
        xmlDoc.loadXML strResponse
        Set xmlTempResponse = xmlDoc.documentElement
        Set xmlPayment = xmlGetMandatoryNode(xmlTempResponse, "PAYMENTRECORD")
        strAppNo = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
        
        'Check if Application is locked
        intLockStatus = GetApplicationLockStatus(strAppNo, strResponse)
        
        'SYS4674 - Application may not have a lock element.
        'SYS4076 - If the UserID and the UnitID on the lock are the same, allow batch to continue.
        'Read the lock information into a DOM document and get the application lock information.
        xmlDoc.loadXML strResponse
        Set xmlAppLock = xmlGetNode(xmlDoc, "//APPLICATIONLOCK")
        
        'Get the batch requesters User and UnitID if available.
        sUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
        sUnitID = xmlGetAttributeText(vxmlRequestNode, "UNITID")
        
        'Get the 'lockers' User and UnitID if available.
        If Not xmlAppLock Is Nothing Then
            sLockUserID = xmlGetNodeText(xmlAppLock, "USERID")
            sLockUnitID = xmlGetNodeText(xmlAppLock, "UNITID")
        End If
        
        'If the application is unlocked, or is locked by the same user.
        If (intLockStatus = lsUnlocked) Or ((intLockStatus = lsLocked And sUserId = sLockUserID And sUnitID = sLockUnitID)) Then
            'No lock found, so lock application/customers
            'Get the ApplicationFactFindNumber for this application
            strValue = GetFactFindNumberForApplication(strAppNo)
            Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
            
            'SYS4894 - Do not attempt to lock the application if it is already locked
            If intLockStatus = lsUnlocked Then
                'BM0339 MDC 05/03/2003 - Lock only application
                'Call LockApplication(xmlTempRequest, strAppNo, strValue)
                Call LockApplicationOnly(xmlTempRequest, strAppNo, strValue)
                'BM0339 MDC 05/03/2003 - End
            End If
            
            ' PSC 20/09/2002 BMIDS00480 -Start
            strROFValueId = GetFirstComboValueId("PaymentType", "N")
            Set xmlDisb = xmlGetMandatoryNode(xmlPayment, "DISBURSEMENTPAYMENT")
            strPaymentType = xmlGetMandatoryAttributeText(xmlDisb, "PAYMENTTYPE")
            
            ' For Return Of Funds we just need to set the status to paid else we
            ' need to validate the inputs and produce the specified output
            If strPaymentType = strROFValueId Then
                Call UpdatePaymentStatus(xmlDisb, "P")
                'BM0339 MDC 06/03/2003
                'Unlock the Application
                 Call UnlockApplicationOnly(strAppNo, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID"))
                'BM0339 MDC 06/03/2003 - End
            Else
                'Validate that Payment and Payee meet Batch criteria
                If ValidatePayment(xmlPayment, strBatchNumber, strBatchRunNumber, "S", True) Then

                    'Evaluate Payment Output
                    strValueId = xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTMETHOD")
                    
                    ' PSC 16/11/2002 BMIDS00968 - Start
                    If IsValidationType("PaymentMethod", strValueId, "B") Then
                        'BACS
                        'SG 28/05/02 SYS4767 START
                        'SG 25/03/02 SYS4202
                        Set xmlBACSResponse = BACSPayment(xmlPayment)
                        If errCheckXMLResponseNode(xmlBACSResponse.selectSingleNode("RESPONSE"), , True) <> 0 Then
                            Call WriteRejectReport
                        End If
                        'SG 25/03/02 SYS4202
                        'SG 28/05/02 SYS4767 END
                    ElseIf IsValidationType("PaymentMethod", strValueId, "YC") Then
                        'CHAPS/TT
                        'SG 28/05/02 SYS4767 START
                        'STB: MSMS0025 - Produce CHAPS or HEXAGON file format.
                        Set objGlobalParam = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
                        sCHAPSInterface = objGlobalParam.GetCurrentParameterByType("PaymentFileInterfaceFormat", "STRING")
    
                        'It'll be either CHAPS or HEXAGON. Call appropriate method on PaymentProcessingBO
                        If sCHAPSInterface = "CHAPS" Then
                            ProcessCHAPSPayment xmlPayment, vxmlRequestNode.selectSingleNode("//@UNITID").Text, xmlCHAPSResponse
                        Else
                            ProcessHexagonPayment xmlPayment, xmlCHAPSResponse
                        End If
                        
                        'Checck XML error resopnse and write reject report if required.
                        If errCheckXMLResponseNode(xmlCHAPSResponse, , True) <> 0 Then
                            Call WriteRejectReport
                        End If
                        'STB: MSMS0025 - End.
                        'SG 28/05/02 SYS4767 END
                    ElseIf IsValidationType("PaymentMethod", strValueId, "CH") Then
                        'Cheque
                        '#TASK - Printing
                    ElseIf IsValidationType("PaymentMethod", strValueId, "TR") Then
                        'Internal Transfer
                        '#TASK - Printing
                    End If
                    ' PSC 16/11/2002 BMIDS00968 - End
                                
                    'Update the Payment Status
                    ' PSC 20/09/2002 BMIDS00480 - Moved further up
                    ' Set xmlDisb = xmlGetMandatoryNode(xmlPayment, "DISBURSEMENTPAYMENT")
                    Call UpdatePaymentStatus(xmlDisb, "P")
                    
                    'BM0339 MDC 06/03/2003
                    'Unlock the Application
                    Call UnlockApplicationOnly(strAppNo, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID"))
                    'BM0339 MDC 06/03/2003 - End
                    
                Else
                    'Didn't Validate the payment
                    'INR BMIDS628 LogWarningMessage is called from within validate
                    'instead of raising an error, where we have access to the error info
'                    'Write Reject Report
'                    Call WriteRejectReport
'                    errThrowError strFunctionName, oePPBatchPaymentPayeeInvalid
                    
                    'INR BMIDS628 Unlock the Application
                    Call UnlockApplicationOnly(strAppNo, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID"))
                    
                    'INR BMIDS628 Indicate that there has been a failure
                    xmlSetAttributeValue vxmlResponseNode, "TYPE", "FAIL"

                End If
                
            End If
            ' PSC 20/09/2002 BMIDS00480 - End
        'Else it must be locked by someone else.
        ElseIf intLockStatus <> lsError Then
            'Lock found, so post message back to queue and retry later
            If DoRetry(vxmlRequestNode) Then
                'Indicate that retrying so batch schedule success/failure not updated at this time
                xmlSetAttributeValue vxmlResponseNode, "TYPE", "RETRY"
            End If

        Else
            'Unexpected error has occurred
            'Write to reject report
            'INR BMIDS628 call LogWarningMessage instead of raising an error
'            Call WriteRejectReport
'            'Raise error
'            errCheckXMLResponse strResponse, True, vxmlResponseNode
            strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
            strMessage = "Batch Transaction Error: InterfacePaidPayment, Unexpected error has occurred"
            strMessage = strMessage & vbCrLf & "Application: " & _
                        xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
            strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                        xmlGetMandatoryAttributeText(xmlPayment, "PAYMENTSEQUENCENUMBER")
            strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
            Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
                    
            'INR BMIDS628 Indicate that there has been a failure
            xmlSetAttributeValue vxmlResponseNode, "TYPE", "FAIL"

        End If
        'SYS4076 - End.
    Else
        'JD BMIDS604 Batch is cancelled, raise an error to prevent the batch status being set as Complete
        errThrowError strFunctionName, 7025
    End If
    
    'SYS3018 MDC 05/12/2001
    gobjContext.SetComplete
    'SYS3018 End
    
PaySanctionedPayment_Exit:
    Set xmlBatchAudit = Nothing
    Set objPayProcBO = Nothing
    Set objMessageQueue = Nothing
    Set xmlBatchSched = Nothing
    Set xmlPayment = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlQueueRequest = Nothing
    Set xmlQueueNode = Nothing
    Set xmlDisb = Nothing
    Set xmlDoc = Nothing
    
    'SG 28/05/02 SYS4767 START
    'STB: MSMS0025.
    Set objGlobalParam = Nothing
    Set xmlCHAPSResponse = Nothing
    'SG 28/05/02 SYS4767 END
    
    'SYS4076 - release application lock node.
    Set xmlAppLock = Nothing
    
    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError strFunctionName, TypeName(Me)
    End If
    
    Exit Sub

PaySanctionedPaymentVbErr:
    gobjContext.SetAbort
    GoTo PaySanctionedPayment_Exit
    
End Sub

Private Function ValidatePayment( _
    ByVal vnodePaymentRecord As IXMLDOMNode, _
    ByVal vstrBatchNumber As String, _
    ByVal vstrBatchRunNumber As String, _
    ByVal vstrRequiredPaymentStatus As String, _
    Optional ByVal vblnValidatePayeeHistory As Boolean = False) As Boolean

' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'RF     25/02/02    SYS4161
'MDC    02/01/2003  BM0215
'MDC    06/03/2003  BM0339
'------------------------------------------------------------------------------------------
On Error GoTo ValidatePayment_Exit

    Const cstrFunctionName As String = "ValidatePayment"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim nodeTempResponse As IXMLDOMNode
    Dim nodeBatchParams As IXMLDOMNode
    Dim nodeDisbursementPayment As IXMLDOMNode
    Dim nodePayeeHistory As IXMLDOMNode
    
    Dim objPayProcDO As IPAymentProcessingDO
    
    Dim strPaymentMethod As String
    Dim strComboValueId As String
    Dim blnCheque  As Boolean
    Dim blnIsValid As Boolean
    Dim strDisbursementPaymentStatus As String
    'INR BMIDS628
    Dim strMessage As String
    
    blnIsValid = True
    blnCheque = False
    
    '------------------------------------------------------------------------------------------
    'Check Payment has correct status
    '------------------------------------------------------------------------------------------
    
    'Get the Disbursement Payment node
    Set nodeDisbursementPayment = _
        xmlGetMandatoryNode(vnodePaymentRecord, "DISBURSEMENTPAYMENT")
   
    'Find required Payment Status value id
    
    strDisbursementPaymentStatus = _
        xmlGetMandatoryAttributeText(nodeDisbursementPayment, "PAYMENTSTATUS")
    
    strComboValueId = GetFirstComboValueId("PaymentStatus", vstrRequiredPaymentStatus)
    
    If strDisbursementPaymentStatus <> strComboValueId Then
        ' allow for "Interface Failed"
        strComboValueId = GetFirstComboValueId("PaymentStatus", "IF")
        If strDisbursementPaymentStatus <> strComboValueId Then
            blnIsValid = False
        End If
    End If
        
    If blnIsValid And vstrRequiredPaymentStatus = "S" Then

        '------------------------------------------------------------------------------------------
        'Check Payment Method is valid for this Batch
        '------------------------------------------------------------------------------------------
    
        'Get method of payment
        strPaymentMethod = xmlGetAttributeText(vnodePaymentRecord, "PAYMENTMETHOD")
        
        'Check that a Payment Method has been specified
        If Len(strPaymentMethod) = 0 Then
            blnIsValid = False
        End If
        
        'Get the Payment Processing batch parameters
        
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        Set nodeTempResponse = xmlDoc.createElement("RESPONSE")
        
        Set nodeBatchParams = xmlDoc.createElement("BATCHPARAMETERS")
        xmlSetAttributeValue nodeBatchParams, "BATCHNUMBER", vstrBatchNumber
        
        Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
        objPayProcDO.GetBatchParameters nodeBatchParams, nodeTempResponse
        Set nodeBatchParams = xmlGetMandatoryNode(nodeTempResponse, "PAYPROCBATCHPARAMS")
        
        If blnIsValid Then
            'Check for BACS
            'BM0215 MDC 02/01/2003
'            strComboValueId = GetFirstComboValueId("PaymentMethod", "B")
'            If (strPaymentMethod = strComboValueId) And Not _
'                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "BACS") Then
            If IsValidationType("PaymentMethod", strPaymentMethod, "B") And Not _
                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "BACS") Then
            'BM0215 MDC 02/01/2003 - End
                blnIsValid = False
            End If
        End If
        
        If blnIsValid Then
            'Check for CHAPS/TT
            'BM0215 MDC 02/01/2003
'            strComboValueId = GetFirstComboValueId("PaymentMethod", "YC")
'            If (strPaymentMethod = strComboValueId) And Not _
'                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "CHAPSORTT") Then
            If IsValidationType("PaymentMethod", strPaymentMethod, "YC") And Not _
                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "CHAPSORTT") Then
            'BM0215 MDC 02/01/2003 - End
                blnIsValid = False
            End If
        End If
        
        If blnIsValid Then
            'Check for Cheque!
            'BM0215 MDC 02/01/2003
            If IsValidationType("PaymentMethod", strPaymentMethod, "CH") Then
                'This payment is a to be made by cheque
                If Not xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "CHEQUE") Then
                    blnIsValid = False
                Else
                    blnCheque = True
                End If
            End If
'            strComboValueId = GetFirstComboValueId("PaymentMethod", "CH")
'            If (strPaymentMethod = strComboValueId) And Not _
'                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "CHEQUE") Then
'                blnIsValid = False
'            Else
'                blnCheque = True
'            End If
            'BM0215 MDC 02/01/2003 - End
        End If
        
        If blnIsValid Then
            'Check for Internal Transfer
            'BM0215 MDC 02/01/2003
'            strComboValueId = GetFirstComboValueId("PaymentMethod", "TR")
'            If (strPaymentMethod = strComboValueId) And Not _
'                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "INTERNALTRANSFER") Then
            If IsValidationType("PaymentMethod", strPaymentMethod, "TR") And Not _
                xmlGetMandatoryAttributeAsBoolean(nodeBatchParams, "INTERNALTRANSFER") Then
            'BM0215 MDC 02/01/2003 - End
                blnIsValid = False
            End If
        End If
    
        'BM0339 MDC 06/03/2003
        If Not blnIsValid Then
            'INR BMIDS628 call LogWarningMessage instead of raising an error
'            errThrowError cstrFunctionName, oePPBatchInvalidPaymentMethod
            strMessage = "Batch Transaction Error: ValidatePayment: Payment method invalid"
            strMessage = strMessage & vbCrLf & "Application: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "APPLICATIONNUMBER")
            strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "PAYMENTSEQUENCENUMBER")
'            strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
            Call LogWarningMessage(vstrBatchNumber, vstrBatchRunNumber, strMessage)

        End If
        'BM0339 MDC 06/03/2003 - End
        
        '
        ' Require Account details for non Cheque Payments
        '
        'If blnIsValid And vblnValidatePayeeHistory And Not blnCheque Then
        If vblnValidatePayeeHistory And Not blnCheque Then
            '
            ' Non Cheque - Account Details required
            '
            Set nodePayeeHistory = xmlGetNode(vnodePaymentRecord, "PAYEEHISTORY")
            'BMIDS628 Be more specific about what details are missing
            Dim strMissingDetail As String
            '
            ' Validate Payee
            '
            If nodePayeeHistory Is Nothing Then
               blnIsValid = False
            ElseIf Len(xmlGetAttributeText(nodePayeeHistory, "BANKSORTCODE")) = 0 Then
               strMissingDetail = " BANKSORTCODE "
               blnIsValid = False
            ElseIf Len(xmlGetAttributeText(nodePayeeHistory, "ACCOUNTNUMBER")) = 0 Then
               strMissingDetail = " ACCOUNTNUMBER "
               blnIsValid = False
            ElseIf Len(xmlGetAttributeText(nodePayeeHistory, "BANKNAME")) = 0 Then
               strMissingDetail = " BANKNAME "
               blnIsValid = False
            End If
            
            'BM0339 MDC 06/03/2003
            If Not blnIsValid Then
                
                'INR BMIDS628 call LogWarningMessage instead of raising an error
    '            errThrowError cstrFunctionName, oePPBatchMissingPayeeDetails
                strMessage = "Batch Transaction Error: ValidatePayment: Payee bank details missing "
                strMessage = strMessage & strMissingDetail
                strMessage = strMessage & vbCrLf & "Application: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "APPLICATIONNUMBER")
                strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "PAYMENTSEQUENCENUMBER")
'                strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
                Call LogWarningMessage(vstrBatchNumber, vstrBatchRunNumber, strMessage)

            End If
            'BM0339 MDC 06/03/2003 - End
            
        End If
    ElseIf vstrRequiredPaymentStatus = "S" Then
        'INR BMIDS628 call LogWarningMessage instead of raising an error
'            errThrowError cstrFunctionName, oePPBatchIncorrectPaymentStatus
        strMessage = "Batch Transaction Error: ValidatePayment: Payment Status incorrect"
            strMessage = strMessage & vbCrLf & "Application: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "APPLICATIONNUMBER")
            strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & _
                        xmlGetMandatoryAttributeText(vnodePaymentRecord, "PAYMENTSEQUENCENUMBER")
'        strMessage = strMessage & vbCrLf & "BatchAuditGuid: " & strBatchAuditGuid
        Call LogWarningMessage(vstrBatchNumber, vstrBatchRunNumber, strMessage)

    End If
    
    ValidatePayment = blnIsValid
    
ValidatePayment_Exit:

    Set xmlDoc = Nothing
    Set nodeTempResponse = Nothing
    Set nodeBatchParams = Nothing
    Set nodeDisbursementPayment = Nothing
    Set nodePayeeHistory = Nothing
    Set objPayProcDO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Function

'MAR1651 GHun No longer used
'Private Sub SetMinimumApplicationStage( _
'    ByVal strAppNo As String, ByVal strAppFFNo As String, ByVal strStage As String)
'' header ----------------------------------------------------------------------------------
'' description:
''   Ensure Application Stage is at least the stage specified
'' pass:
''   Application Number, Application Fact Find Number, Stage Validation Type
'' return:
''   n/a
'' history:
'' MC    10/04/2001  Created.
''------------------------------------------------------------------------------------------
'On Error GoTo SetMinimumApplicationStage_Exit
'
'    Const strFunctionName As String = "SetMinimumApplicationStage"
'
'    'BMIDS697 GHun 04/02/2004
'    'Dim objTmBO As OmTmBO
'    Dim objTmNoTxBO As omTmNoTxBO
'    'BMIDS697 End
'
'    Dim objAppBO As ApplicationBO
'
'    Dim xmlDoc As FreeThreadedDOMDocument40
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim xmlGetStageRequest As IXMLDOMNode
'    Dim xmlAppStageNode As IXMLDOMNode
'    Dim xmlAppNode As IXMLDOMNode
'    Dim xmlStageNode As IXMLDOMNode
'    Dim xmlNode As IXMLDOMNode
'
'    Dim strRequiredStageId As String
'    Dim strCurrentStageId As String
'    Dim strResponse As String
'
'    'BMIDS697 GHun 04/02/2004
'    'Set objTmBO = gobjContext.CreateInstance("omTm.omTmBO")
'    Set objTmNoTxBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmNoTxBO")
'    'BMIDS697 End
'
'    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    xmlDoc.validateOnParse = False
'    xmlDoc.setProperty "NewParser", True
'
'    'Build request
'    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'    Set xmlAppNode = xmlDoc.createElement("APPLICATION")
'    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
'    xmlNode.Text = strAppNo
'    xmlAppNode.appendChild xmlNode
'    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
'    xmlNode.Text = strAppFFNo
'    xmlAppNode.appendChild xmlNode
'    xmlTempRequest.appendChild xmlAppNode
'
'    'Get Required Stage ----------------------------------------------------------------
'    strRequiredStageId = GetFirstComboValueId("ApplicationStage", strStage)
'
'    'Build Update Stage Request ---------------------------------------------------------
'    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToNextStage"
'    xmlSetAttributeValue xmlTempRequest, "USERAUTHORITYLEVEL", "99"
'    'Application
'    Set xmlAppNode = xmlDoc.createElement("APPLICATION")
'    xmlSetAttributeValue xmlAppNode, "APPLICATIONNUMBER", strAppNo
'    xmlSetAttributeValue xmlAppNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
'    xmlTempRequest.appendChild xmlAppNode
'    'Current Stage
'    Set xmlStageNode = xmlDoc.createElement("CURRENTSTAGE")
'    xmlSetAttributeValue xmlStageNode, "SOURCEAPPLICATION", "Omiga"
'    xmlSetAttributeValue xmlStageNode, "CASEID", strAppNo
'    xmlSetAttributeValue xmlStageNode, "ACTIVITYID", "10"
'    xmlTempRequest.appendChild xmlStageNode
'
'    'Get the current application stage --------------------------------------------------
'    Set xmlGetStageRequest = xmlDoc.createElement("REQUEST")
'    Set xmlAppStageNode = xmlDoc.createElement("APPLICATIONSTAGE")
'    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
'    xmlNode.Text = strAppNo
'    xmlAppStageNode.appendChild xmlNode
'    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
'    xmlNode.Text = strAppFFNo
'    xmlAppStageNode.appendChild xmlNode
'    xmlGetStageRequest.appendChild xmlAppStageNode
'    strResponse = objAppBO.GetCurrentApplicationStage(xmlGetStageRequest.xml)
'    errCheckXMLResponse strResponse, True
'    xmlDoc.loadXML strResponse
'    Set xmlAppNode = xmlGetMandatoryNode(xmlDoc.documentElement, ".//APPLICATIONSTAGE")
'    Set xmlNode = xmlGetMandatoryNode(xmlAppNode, "STAGENUMBER")
'    strCurrentStageId = xmlNode.Text 'Get Current Stage
'
'
'    'Update Stage until required stage is reached ----------------------------------------
'    Do Until CSafeLng(strCurrentStageId) >= CSafeLng(strRequiredStageId)
'
'        'Move to next stage
'        xmlSetAttributeValue xmlStageNode, "STAGEID", strCurrentStageId
'
'        'BMIDS697 GHun 04/02/2004
'        'strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
'        strResponse = objTmNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
'        'BMIDS697 End
'
'        errCheckXMLResponse strResponse, True
'
'        'Get the updated current application stage
'        strResponse = objAppBO.GetCurrentApplicationStage(xmlGetStageRequest.xml)
'        errCheckXMLResponse strResponse, True
'        xmlDoc.loadXML strResponse
'        Set xmlAppNode = xmlGetMandatoryNode(xmlDoc.documentElement, ".//APPLICATIONSTAGE")
'        Set xmlNode = xmlGetMandatoryNode(xmlAppNode, "STAGENUMBER")
'        strCurrentStageId = xmlNode.Text 'Get Current Stage
'
'    Loop
'
'SetMinimumApplicationStage_Exit:
'    Set objTmNoTxBO = Nothing
'    Set xmlDoc = Nothing
'    Set xmlTempRequest = Nothing
'    Set xmlAppNode = Nothing
'    Set xmlNode = Nothing
'    Set xmlGetStageRequest = Nothing
'    Set xmlAppStageNode = Nothing
'
'    errCheckError strFunctionName, TypeName(Me)
'
'End Sub
'MAR1651 End

' PSC 15/08/2002 BMIDS00330 - Pass vxmlRequestNode, vstrAppFFNo and vxmlPaymentNode
Private Sub CalculateMonthlyRepayments(ByVal strAppNo As String, ByVal strAppFFNo As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlPaymentNode As IXMLDOMNode)
'------------------------------------------------------------------------
' Procedure:    CalculateMonthlyRepayments
' Author:       Mark Coleman
' Date:         11/04/2001
' Purpose:      Calculate the Monthly Repayment for each Loan Component and update the
'               Loan Component Payment accordingly
' Input parameters:
'               nodePayment
' Output parameters:
'               n/a
' Notes:
'               Late binding to ApplicationQuoteBO is used to avoid a circular reference.
'               omAQ has a reference to omPayProc in order to create Application Fee Types.
'               As the current method is run in Batch, speed is less of an issue.
'------------------------------------------------------------------------

On Error GoTo CalculateMonthlyRepayments_Exit

Const strFunctionName As String = "CalculateMonthlyRepayments"

Dim objMortSQBO As MortgageSubQuoteBO
Dim objPayProcDO As IPAymentProcessingDO
Dim objAppQuoteBO As Object 'ApplicationQuoteBO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlElement As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlLoanComponentList As IXMLDOMNodeList
Dim xmlComponent As IXMLDOMNode
Dim xmlOneOffCost As IXMLDOMNode
Dim xmlLCPNode As IXMLDOMNode
Dim xmlLCPList As IXMLDOMNodeList
Dim xmlCalcsRequest As IXMLDOMNode
'Dim xmlCalcsOutput As IXMLDOMNode  'BMIDS736
Dim xmlAmountNode As IXMLDOMNode

' PSC 15/08/2002 BMIDS00330 - Start
Dim xmlCompListNode As IXMLDOMNode
Dim xmlIntOnlyNode As IXMLDOMNode
Dim xmlCapAndIntNode As IXMLDOMNode
Dim xmlLCDoc As FreeThreadedDOMDocument40
Dim xmlDisp As IXMLDOMNode
Dim xmlOneOffCostList  As IXMLDOMNodeList
Dim xmlOneOffCostListNode As IXMLDOMNode
' PSC 15/08/2002 BMIDS00330 - End

Dim strResponse As String
Dim lngRet As Long
Dim intMainLoanNo As Integer
Dim intCurrentNo As Integer
Dim strMainLoanNo As String
Dim strMSQNumber As String
Dim strValue As String
Dim strMonthlyPayment As String
'Dim strInitialPayment As String
'Dim blnDoneOneOffCosts As Boolean

' PSC 15/08/2002 BMIDS00330 - Start
Dim strAccruedInterest As String
Dim strPattern As String
Dim strLCSeqNo As String
Dim blnInitialAdv As Boolean
Dim lngFeesAddedToloan As Long
' PSC 15/08/2002 BMIDS00330 - End

' BMIDS774
Dim xmlCompletionDateNode As IXMLDOMNode

    
    Set objMortSQBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")

    'Create Request to retrieve a list of Loan Components for the Active Quotation ------------------------
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    Set xmlElement = xmlDoc.createElement("FINDACTIVELOANCOMPONENTS")
    xmlTempRequest.appendChild xmlElement
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strAppNo
    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strAppFFNo
    
    'Get the list of Loan Components and find the Mortgage Sub Quote Number --------------------------------
    strResponse = objMortSQBO.FindActiveLoanComponentsList(xmlTempRequest.xml)
    lngRet = errCheckXMLResponse(strResponse, True)
    xmlDoc.loadXML strResponse
    strMSQNumber = xmlGetMandatoryNodeText(xmlDoc.documentElement, ".//LOANCOMPONENT/MORTGAGESUBQUOTENUMBER")
   
    'Get Full Loan Composition data for the mortgage sub quote ---------------------------------------------
    Set xmlLCDoc = New FreeThreadedDOMDocument40
    xmlLCDoc.validateOnParse = False
    xmlLCDoc.setProperty "NewParser", True
    xmlLCDoc.async = False
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    Set xmlElement = xmlLCDoc.createElement("LOANCOMPOSITION")
    xmlTempRequest.appendChild xmlElement
    Set xmlNode = xmlLCDoc.createElement("APPLICATIONNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strAppNo
    Set xmlNode = xmlLCDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strAppFFNo
    Set xmlNode = xmlLCDoc.createElement("MORTGAGESUBQUOTENUMBER")
    xmlElement.appendChild xmlNode
    xmlNode.Text = strMSQNumber
    Set objAppQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    strResponse = objAppQuoteBO.GetLoanCompositionDetails(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlLCDoc.loadXML strResponse
    
    Set xmlLoanComponentList = xmlLCDoc.selectNodes(".//LOANCOMPONENT[PORTEDLOAN != '1']")
    
    ' Find main loan component number ----------------------------------------------------------------------
    intMainLoanNo = 9999
    For Each xmlComponent In xmlLoanComponentList
        strValue = xmlGetMandatoryNodeText(xmlComponent, ".//LOANCOMPONENTSEQUENCENUMBER")
        intCurrentNo = CSafeInt(strValue)
        If intCurrentNo < intMainLoanNo Then
            intMainLoanNo = intCurrentNo
        End If
    Next
    strMainLoanNo = CStr(intMainLoanNo)
    
    'Setup Request for Calcs Engine -------------------------------------------------------------------------
    Set xmlCalcsRequest = vxmlRequestNode.cloneNode(False)
     
    ' PSC 15/08/2002 BMIDS00330 - Start
    ' Create a list element to attach loan components to
    Set xmlCompListNode = xmlLCDoc.createElement("LOANCOMPONENTLIST")
    
    'JD BMIDS791 add an attribute to the calcs request to say we are calling from payment processing
    xmlSetAttributeValue xmlCompListNode, "FROMPAYPROC", "TRUE"
    
    ' Get the list of loan component payments for this disbursement
    Set xmlLCPList = vxmlPaymentNode.selectNodes("LOANCOMPONENTPAYMENT")
    
    ' Is this an initial advance
    Set xmlDisp = xmlGetMandatoryNode(vxmlPaymentNode, "DISBURSEMENTPAYMENT")
    strValue = xmlGetMandatoryAttributeText(xmlDisp, "PAYMENTTYPE")
    blnInitialAdv = IsValidationType("PaymentType", strValue, "I")
    
    ' Set fees to zero
    lngFeesAddedToloan = 0
     
    ' Set up each loan component for the relevant loan component payment
    For Each xmlLCPNode In xmlLCPList
        
        ' Find the relevant loan component
        strLCSeqNo = xmlGetMandatoryAttributeText(xmlLCPNode, "LOANCOMPONENTSEQUENCENUMBER")
        strPattern = ".//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER='" & strLCSeqNo & "']"
        Set xmlComponent = xmlGetMandatoryNode(xmlLCDoc, strPattern)
        
        ' Copy the loan component payment values into the loan component
        strValue = xmlGetAttributeText(xmlLCPNode, "AMOUNT")
        Set xmlAmountNode = xmlGetMandatoryNode(xmlComponent, "TOTALLOANCOMPONENTAMOUNT")
        xmlAmountNode.Text = strValue
                
        strValue = xmlGetAttributeText(xmlLCPNode, "INTERESTONLYELEMENT")
        Set xmlIntOnlyNode = xmlGetMandatoryNode(xmlComponent, "INTERESTONLYELEMENT")
        xmlIntOnlyNode.Text = strValue
        
        strValue = xmlGetAttributeText(xmlLCPNode, "CAPITALANDINTERESTELEMENT")
        Set xmlCapAndIntNode = xmlGetMandatoryNode(xmlComponent, "CAPITALANDINTERESTELEMENT")
        xmlCapAndIntNode.Text = strValue
        
        ' If this is the main loan and it is an initial advance then add on the one off costs and the shortfall
        If strLCSeqNo = strMainLoanNo And blnInitialAdv Then
        
        
            'Get OneOffCosts ---------------------------------------------------------------------------------------
            'Create element based request for ONEOFFCOST
            ' <ONEOFFCOST>
            '   <APPLICATIONNUMBER>ADP00108391</APPLICATIONNUMBER>
            '   <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
            '   <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
            ' </ONEOFFCOST>
            strMSQNumber = xmlGetMandatoryNodeText(xmlDoc.documentElement, _
                    ".//LOANCOMPONENT[LOANCOMPONENTSEQUENCENUMBER='" & strMainLoanNo & "']/MORTGAGESUBQUOTENUMBER")
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            Set xmlOneOffCost = xmlDoc.createElement("ONEOFFCOST")
            xmlTempRequest.appendChild xmlOneOffCost
            Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
            xmlOneOffCost.appendChild xmlNode
            xmlNode.Text = strAppNo
            Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
            xmlOneOffCost.appendChild xmlNode
            xmlNode.Text = strAppFFNo
            Set xmlNode = xmlDoc.createElement("MORTGAGESUBQUOTENUMBER")
            xmlOneOffCost.appendChild xmlNode
            xmlNode.Text = strMSQNumber
            strResponse = objMortSQBO.GetOneOffCostsDetails(xmlTempRequest.xml)
            lngRet = errCheckXMLResponse(strResponse, True)
            xmlDoc.loadXML strResponse
            Set xmlOneOffCostListNode = xmlDoc.selectSingleNode(".//MORTGAGEONEOFFCOSTLIST")
            
            ' Total up the costs to be added to the loan
            Set xmlOneOffCostList = xmlOneOffCostListNode.selectNodes("MORTGAGEONEOFFCOST[ADDTOLOAN='1']")
            
            For Each xmlOneOffCost In xmlOneOffCostList
                lngFeesAddedToloan = lngFeesAddedToloan + xmlGetNodeAsLong(xmlOneOffCost, "AMOUNT")
                xmlOneOffCostListNode.removeChild xmlOneOffCost
            Next
                        
            ' Add on any shortfall and the fees
            strValue = xmlGetAttributeText(xmlDisp, "SHORTFALL", "0")
            xmlAmountNode.Text = CLng(xmlAmountNode.Text) + CLng(strValue) + lngFeesAddedToloan
            
            ' If we have part and part then add on to capital an interest element
            If Len(xmlCapAndIntNode.Text) > 0 Then
                xmlCapAndIntNode.Text = CLng(xmlCapAndIntNode.Text) + CLng(strValue) + lngFeesAddedToloan
            End If
                        
        End If
        
        'BMIDS774
        strValue = xmlGetAttributeText(xmlDisp, "COMPLETIONDATE")
        'BMIDS774 If EstimatedCompletionDate is >= Sys Date, use it
        'otherwise don't create it so we get the alpha+ default.
        If DateDiff("d", Now, CSafeDate(strValue)) >= 0 Then
            Set xmlCompletionDateNode = xmlDoc.createElement("COMPLETIONDATE")
            xmlCompletionDateNode.Text = strValue
            xmlComponent.appendChild xmlCompletionDateNode
        End If

        xmlCompListNode.appendChild xmlComponent.cloneNode(True)
    Next
    
    ' Call Composite Rate Calcs
    xmlCalcsRequest.appendChild xmlCompListNode
    
    strResponse = objMortSQBO.RunCompositeRateCalc(xmlCalcsRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse
    
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set xmlLoanComponentList = xmlDoc.selectNodes(".//LOANCOMPONENT")
     
    ' Update each loan component payment with the monthly costs
    For Each xmlComponent In xmlLoanComponentList
        ' Reset temporary response node
        Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
        
        'BMIDS736 GHun
        ' Get the calcs output for this component
        'Set xmlCalcsOutput = xmlGetMandatoryNode(xmlComponent, "MORTCALCOUTPUT")
        'BMIDS736 End
    
        ' Find the loan component payment for this component
        strLCSeqNo = xmlGetMandatoryNodeText(xmlComponent, "LOANCOMPONENTSEQUENCENUMBER")
        strPattern = ".//LOANCOMPONENTPAYMENT[@LOANCOMPONENTSEQUENCENUMBER='" & strLCSeqNo & "']"
        Set xmlLCPNode = xmlGetMandatoryNode(vxmlPaymentNode, strPattern)
        
        ' Get Gross Monthly Cost &b Accrued Interest
        'BMIDS736 GHun
        'strMonthlyPayment = xmlGetNodeText(xmlCalcsOutput, ".//OUT130_TOTALGROSSMORTGAGEPAYMENT1")
        'strAccruedInterest = xmlGetNodeText(xmlCalcsOutput, ".//OUT020_ACCRUEDINTEREST")
        strMonthlyPayment = xmlGetNodeText(xmlComponent, "NETMONTHLYCOST")
        strAccruedInterest = xmlGetNodeText(xmlComponent, "ACCRUEDINTEREST")
        'BMIDS736 End
        
        ' Update Loan component payment
        xmlSetAttributeValue xmlLCPNode, "INITIALREPAYMENT", strMonthlyPayment
        xmlSetAttributeValue xmlLCPNode, "MONTHLYREPAYMENT", strMonthlyPayment
        xmlSetAttributeValue xmlLCPNode, "ACCRUEDINTEREST", strAccruedInterest
        objPayProcDO.UpdateLoanComponentPayment xmlLCPNode, xmlTempResponse
    Next
    
    ' Update disbursement payment
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    Set xmlNode = xmlDisp.cloneNode(False)
    strMonthlyPayment = xmlGetNodeText(xmlDoc, ".//TOTALGROSSMONTHLYCOST")
    strAccruedInterest = xmlGetNodeText(xmlDoc, ".//TOTALACCRUEDINTEREST")
    xmlSetAttributeValue xmlNode, "TOTALMONTHLYCOST", strMonthlyPayment
    xmlSetAttributeValue xmlNode, "TOTALACCRUEDINTEREST", strAccruedInterest
    
    objPayProcDO.UpdateDisbursement xmlNode, xmlTempResponse
    ' PSC 15/08/2002 BMIDS00330 - End
  
CalculateMonthlyRepayments_Exit:
    Set objMortSQBO = Nothing
    
    ' PSC 15/08/2002 BMIDS00330 - Start
    Set objPayProcDO = Nothing
    Set objAppQuoteBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    Set xmlLoanComponentList = Nothing
    Set xmlComponent = Nothing
    Set xmlOneOffCost = Nothing
    Set xmlLCPNode = Nothing
    Set xmlLCPList = Nothing
    Set xmlCalcsRequest = Nothing
    'Set xmlCalcsOutput = Nothing
    Set xmlAmountNode = Nothing
    Set xmlCompListNode = Nothing
    Set xmlIntOnlyNode = Nothing
    Set xmlCapAndIntNode = Nothing
    Set xmlLCDoc = Nothing
    Set xmlDisp = Nothing
    Set xmlOneOffCostList = Nothing
    Set xmlOneOffCostListNode = Nothing
    ' PSC 15/08/2002 BMIDS00330 - End
    ' BMIDS774
    Set xmlCompletionDateNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub UpdatePaymentStatus(ByVal xmlDisb As IXMLDOMNode, ByVal strStatus As String)

On Error GoTo UpdatePaymentStatus_Exit

Const strFunctionName As String = "UpdatePaymentStatus"

Dim objPayProcDO As IPAymentProcessingDO

Dim xmlTempResponse As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strValueId As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Create the Request and Response nodes
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    Set xmlTempRequest = xmlDoc.createElement("DISBURSEMENTPAYMENT")
    xmlCopyMandatoryAttribute xmlDisb, xmlTempRequest, "APPLICATIONNUMBER"
    xmlCopyMandatoryAttribute xmlDisb, xmlTempRequest, "PAYMENTSEQUENCENUMBER"
    strValueId = GetFirstComboValueId("PaymentStatus", strStatus)
    xmlSetAttributeValue xmlTempRequest, "PAYMENTSTATUS", strValueId
    
    'Update the Payment Status
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    objPayProcDO.UpdateDisbursement xmlTempRequest, xmlTempResponse

UpdatePaymentStatus_Exit:
    Set objPayProcDO = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempRequest = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Function IsBatchCancelled(ByVal xmlBatchScheduleNode As IXMLDOMNode) As Boolean

On Error GoTo IsBatchCancelled_Exit

Const strFunctionName As String = "IsBatchCancelled"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode

Dim objBatchSched As BatchScheduleBO

Dim strValue As String
Dim strValueId As String
Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Get the current status of the batch schedule
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    Set xmlTempRequest = xmlDoc.createElement("REQUEST") 'vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetBatchSchedule"
    xmlTempRequest.appendChild xmlBatchScheduleNode
    
    strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    xmlDoc.loadXML strResponse
    Set xmlNode = xmlGetMandatoryNode(xmlDoc.documentElement, "BATCHSCHEDULE")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "STATUS")
    
    'Get the combovalue for Cancelled
    strValueId = GetFirstComboValueId("BatchScheduleStatus", "CA")
    
    If strValue = strValueId Then
        'Batch Cancelled
        IsBatchCancelled = True
    Else
        IsBatchCancelled = False
    End If

IsBatchCancelled_Exit:
    Set xmlDoc = Nothing
    Set xmlNode = Nothing
    Set xmlTempRequest = Nothing
    Set objBatchSched = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function GetApplicationLockStatus(ByVal strAppNo As String, strResponse As String) As LOCKSTATUS

On Error GoTo GetApplicationLockStatus_Exit

Const strFunctionName As String = "GetApplicationLockStatus"

Dim objAppManBO As ApplicationManagerBO

Dim xmlTempRequest As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim lngErr As Long

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
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub LockApplication(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String, ByVal strAppFFNo As String)

On Error GoTo LockApplication_Exit

Const strFunctionName As String = "LockApplication"

Dim objAppManBO As ApplicationManagerBO

'Dim xmlTempRequest As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'Create Request
    Set xmlAppNode = xmlDoc.createElement("APPLICATION")
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = strAppNo
    xmlAppNode.appendChild xmlNode
    xmlTempRequest.appendChild xmlAppNode
    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = strAppFFNo
    xmlAppNode.appendChild xmlNode
    
    'Lock the Application and associated customers
    strResponse = objAppManBO.LockCustomersForApplication(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True

LockApplication_Exit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UnlockApplication(ByVal strAppNo As String)

On Error GoTo UnlockApplication_Exit

Const strFunctionName As String = "UnlockApplication"

Dim objAppManBO As ApplicationManagerBO

Dim xmlTempRequest As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'Create Request
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlAppNode.Text = strAppNo
    xmlTempRequest.appendChild xmlAppNode
    
    'Unlock the Application and Customers
    strResponse = objAppManBO.UnlockApplicationAndCustomers(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True

UnlockApplication_Exit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

'BM0339 MDC 06/03/2003
Private Sub UnlockApplicationOnly(ByVal strAppNo As String, ByVal strUserId As String)

On Error GoTo UnlockApplication_Exit

Const strFunctionName As String = "UnlockApplicationOnly"

Dim objAppManBO As ApplicationManagerBO

Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempParent As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'Create Request
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlSetAttributeValue xmlTempRequest, "USERID", strUserId
    Set xmlTempParent = xmlDoc.createElement("APPLICATIONLOCK")
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlAppNode.Text = strAppNo
    xmlTempParent.appendChild xmlAppNode
    xmlTempRequest.appendChild xmlTempParent
    
    'Unlock the Application and Customers
    strResponse = objAppManBO.DeleteLock(xmlTempRequest.xml)
    
    errCheckXMLResponse strResponse, True

UnlockApplication_Exit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub
'BM0339 MDC 06/03/2003 - End

' PSC 20/09/2002 BMIDS00480 - Add optional parameter
Private Function DoRetry(ByVal vxmlRequestNode As IXMLDOMNode, Optional ByVal vlngMessageNo As Long = -1) As Boolean

On Error GoTo DoRetryExit

Const strFunctionName As String = "DoRetry"

Dim objBatchSched As BatchScheduleBO
Dim objMessageQueue As IomMessageQueue

Dim xmlBatchSched As IXMLDOMNode
Dim xmlBatchAudit As IXMLDOMNode
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
Dim strBatchAuditGuid As String

    'Get Request values
    Set xmlBatchSched = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSched, "BATCHRUNNUMBER")
    Set xmlBatchAudit = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    strBatchAuditGuid = xmlGetMandatoryAttributeText(xmlBatchAudit, "BATCHAUDITGUID")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objBatchSched = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    
    'Find maximum number of retries
    lngValue = GetGlobalParamAmount("PProcApplicationRetries")
    
    'Find current number of retries
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindBatchAuditList"
    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGuid
    xmlTempRequest.appendChild xmlNode
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
        xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strBatchAuditGuid
        xmlSetAttributeValue xmlNode, "NUMBEROFRETRIES", CStr(lngRetries + 1)
        xmlTempRequest.appendChild xmlNode
        strResponse = objBatchSched.omBatchRequest(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
        
        '... and retry by posting message back to the queue
        'Get Queue Name & setup Message Queue Request
        strQueueName = GetGlobalParamString("PProcQueueName")
        If Len(Trim$(strQueueName)) = 0 Then
            errThrowError strFunctionName, oeInvalidParameter, "PProcQueueName not found"
        End If
        vxmlRequestNode.removeChild xmlBatchAudit   'Remove the Batch Audit node from the original request
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "SendToQueue"
        Set xmlNode = xmlDoc.createElement("MESSAGEQUEUE")
        xmlSetAttributeValue xmlNode, "QUEUENAME", strQueueName
        xmlSetAttributeValue xmlNode, "PROGID", gstrPAYMENTPROCESSING & ".PaymentProcessingBO"
        xmlSetAttributeValue xmlNode, "XML", vxmlRequestNode.xml
        xmlTempRequest.appendChild xmlNode
        Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
        Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
        Set xmlTempResponse = objMessageQueue.SendToQueue(xmlTempRequest)
        DoRetry = True
        
    Else
        'Max retries exceeded so fail
        Call WriteRejectReport
        
        ' PSC 20/09/2002 BMIDS00480 - Start
        ' Throw the error no passed in else use default
        If vlngMessageNo = -1 Then
            errThrowError strFunctionName, oePPBatchLockingErr
        Else
            errThrowError strFunctionName, vlngMessageNo
        End If
        ' PSC 20/09/2002 BMIDS00480 - End
    End If

DoRetryExit:
    Set objBatchSched = Nothing
    Set objMessageQueue = Nothing
    Set xmlBatchSched = Nothing
    Set xmlBatchAudit = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function



Private Sub IPaymentProcessingNTTxBO_ProcessPayment(ByVal vxmlRequestNode As IXMLDOMNode, _
        ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo ProcessPaymentVbErr

Const strFunctionName As String = "IPaymentProcessingNTTxBO_ProcessPayment"

Dim objBatchSchedBO As omBatch.BatchScheduleBO
Dim objMessageQueue As omPayProc.IomMessageQueue
Dim objPayProcDO As IPAymentProcessingDO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlMessage As IXMLDOMElement

Dim strValue As String
Dim strValueId As String
Dim strResponse As String
Dim strOperation As String
Dim strQueueName As String
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strPaymentJobType As String
Dim strAppNum As String
Dim strPaySeqNum As String
Dim lngPaymentCount As Long
Dim blnExit As Boolean
Dim strPaymentType As String        ' PSC 20/09/2002 BMIDS00480
Dim strExecuteAfterDate As String   'AM BMIDS722

    'Validation
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlNode, "BATCHRUNNUMBER")
    strPaymentJobType = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTJOBTYPE")
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH/PAYMENTRECORD")
    strAppNum = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    strPaySeqNum = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    ' PSC 20/09/2002 BMIDS00480
    strPaymentType = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTTYPE")
    
    'AM BMIDS722 - Start
    'if the transaction contains an EXECUTEAFTER attribute then we need to get it's
    'value because we will need to add it to the message node before sending to the
    'queue
    If xmlAttributeValueExists(xmlNode, "EXECUTEAFTER") Then
        strExecuteAfterDate = xmlGetAttributeText(xmlNode, "EXECUTEAFTER")
    End If
    'AM BMIDS722 - End
    
    'Instantiate Objects
    Set objBatchSchedBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
    Set objPayProcDO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingDO")
    Set objMessageQueue = gobjContext.CreateInstance(App.Title & ".omMessageQueue")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Create Batch Audit --------------------------------------------------------------------
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "CreateBatchAudit"
    Set xmlNode = xmlDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue xmlNode, "NUMBEROFRETRIES", "1"
    xmlTempRequest.appendChild xmlNode
    strResponse = objBatchSchedBO.omBatchRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    
    'Get BatchAuditGuid of newly created record
    xmlDoc.loadXML strResponse
    Set xmlNode = xmlDoc.selectSingleNode(".//BATCHAUDIT")
    strValue = xmlGetMandatoryAttributeText(xmlNode, "BATCHAUDITGUID")
    
    'Create PayProcBatchAudit record -------------------------------------------------------
    Set xmlTempResponse = vxmlResponseNode.cloneNode(False)
    Set xmlNode = xmlDoc.createElement("PAYPROCBATCHAUDIT")
    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue xmlNode, "BATCHAUDITGUID", strValue
    xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNum
    xmlSetAttributeValue xmlNode, "PAYMENTSEQUENCENUMBER", strPaySeqNum
    
    ' PSC 20/09/2002 BMIDS00480
    xmlSetAttributeValue xmlNode, "PAYMENTTYPE", strPaymentType
    objPayProcDO.CreatePaymentAudit xmlNode, xmlTempResponse
    
    'Send a message to the Payment Processing Queue ----------------------------------------
    
    'Get Operation related to the Payment Job Type
    If IsValidationType("PaymentJobType", CSafeInt(strPaymentJobType), "P") Then
        'Pay Sanctioned Payment
        strOperation = "PaySanctionedPayment"
    ElseIf IsValidationType("PaymentJobType", CSafeInt(strPaymentJobType), "I") Then
        'Interface Paid Payment
        strOperation = "InterfacePaidPayment"
    Else
        'Error
        errThrowError strFunctionName, oePPBatchInvalidJobType, strPaymentJobType
    End If

    'Set up XML to be sent as a message
    Set xmlMessage = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlMessage, "OPERATION", strOperation
    Set xmlNode = xmlDoc.createElement("BATCHSCHEDULE")
    xmlSetAttributeValue xmlNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlMessage.appendChild xmlNode
    Set xmlNode = xmlDoc.createElement("PAYMENTRECORD")
    xmlSetAttributeValue xmlNode, "APPLICATIONNUMBER", strAppNum
    xmlSetAttributeValue xmlNode, "PAYMENTSEQUENCENUMBER", strPaySeqNum
    xmlMessage.appendChild xmlNode

    'Get Queue Name & setup Message Queue Request
    strQueueName = GetGlobalParamString("PProcQueueName")
    If Len(Trim$(strQueueName)) = 0 Then
        errThrowError strFunctionName, oeInvalidParameter, "PProcQueueName not found"
    End If
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "SendToQueue"
    Set xmlNode = xmlDoc.createElement("MESSAGEQUEUE")
    xmlSetAttributeValue xmlNode, "QUEUENAME", strQueueName
    xmlSetAttributeValue xmlNode, "PROGID", gstrPAYMENTPROCESSING & ".PaymentProcessingBO"
    xmlSetAttributeValue xmlNode, "XML", xmlMessage.xml
    
    'AM BMIDS722 - Start
    'if the EXECUTEAFTER attribute was set then we must add it to the message node
    'value and add it to the message node as EXECUTEAFTERDATE
    If Len(strExecuteAfterDate) > 0 Then
        xmlSetAttributeValue xmlNode, "EXECUTEAFTERDATE", strExecuteAfterDate
    End If
    'AM BMIDS722 - End
    
    xmlTempRequest.appendChild xmlNode
    Set xmlTempResponse = objMessageQueue.SendToQueue(xmlTempRequest)

    'SYS3018 MDC 05/12/2001
    gobjContext.SetComplete
    'SYS3018 End

ProcessPaymentExit:
    Set objBatchSchedBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlNode = Nothing
    Set xmlMessage = Nothing

    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError strFunctionName, TypeName(Me)
    End If
    
    Exit Sub

ProcessPaymentVbErr:
    gobjContext.SetAbort
    GoTo ProcessPaymentExit

End Sub

Private Sub IPaymentProcessingNTTxBO_SanctionPrintCheque(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode, _
                                                            ByVal bLocked As Boolean)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AW     15/09/2002  BM029
'
On Error GoTo SanctionPrintChequeVbErr

Const strFunctionName As String = "IPaymentProcessingNTTxBO_SanctionPrintCheque"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMNode
Dim xmlAuditRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlChequeAuditNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim strResponse As String
Dim strUserId As String, strAppNo As String, strChequeNo As String, strSeqNo As String, strAppFFNo As String
Dim objTxBO As IPaymentProcessingTxBO
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
    
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    'Validate APPLICATIONNUMBER attribute is specified
    strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strSeqNo = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    'Validate CHEQUENUMBER attribute is specified
    strChequeNo = xmlGetMandatoryAttributeText(xmlNode, "CHEQUENUMBER")
    
    'Get the APPLICATIONFACTFINDNUMBER
    strAppFFNo = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
        
    If Len(Trim$(strAppFFNo)) = 0 Then
        'Get the ApplicationFactFind number
        strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    End If
    'GD 15/10/2002  BMIDS00614 START
    If xmlGetAttributeText(vxmlRequestNode, "REPRINTMODE") <> "1" Then
        ' LOCK APP
        If Not bLocked Then
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            Call LockApplicationOnly(xmlTempRequest, strAppNo, strAppFFNo)
        End If
    End If
    'GD 15/10/2002  BMIDS00614 END
    ' LOCK APP
    'If Not bLocked Then
        'Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        'Call LockApplication(xmlTempRequest, strAppNo, strAppFFNo)
    'End If
        
    'SANCTION PAYMENT
    If objTxBO Is Nothing Then
        Set objTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    'If Not (CSafeBool(xmlGetAttributeText(vxmlRequestNode.selectSingleNode(".//PAYMENTRECORD"), "REPRINTSTATUS"))) Then
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    objTxBO.UpdateChequeDisbursement vxmlRequestNode, xmlTempResponse
    'Check for errors
    errCheckXMLResponse xmlTempResponse.xml, True
    'End If

    ' CREATE AUDIT RECORD
    Set xmlDoc = Nothing
    Set xmlTempResponse = Nothing
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlAuditRequest = xmlDoc.createElement("REQUEST")
    Set xmlChequeAuditNode = xmlDoc.createElement("PAYMENTCHEQUEAUDIT")
    
    xmlSetAttributeValue xmlChequeAuditNode, "APPLICATIONNUMBER", strAppNo
    xmlSetAttributeValue xmlChequeAuditNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
    xmlSetAttributeValue xmlChequeAuditNode, "CAUSERID", strUserId
    xmlSetAttributeValue xmlChequeAuditNode, "PAYMENTSEQUENCENUMBER", strSeqNo
    xmlSetAttributeValue xmlChequeAuditNode, "CACHEQUENUMBER", strChequeNo
    xmlAuditRequest.appendChild xmlChequeAuditNode
    
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    CreateChequeAuditRecord xmlAuditRequest, xmlTempResponse
    'Check for errors
    errCheckXMLResponse xmlTempResponse.xml, True
        
    'PRINT CHEQUE
    PrintCheque vxmlRequestNode, vxmlResponseNode
    'Check for errors
    errCheckXMLResponse vxmlResponseNode.xml, True
    
    gobjContext.SetComplete
  
SanctionPrintChequeExit:
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlAuditRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlChequeAuditNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

    Exit Sub
    
SanctionPrintChequeVbErr:
    gobjContext.SetAbort
    GoTo SanctionPrintChequeExit
End Sub


Private Sub IPaymentProcessingNTTxBO_SetPaidChequePayment(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMElement)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AW     15/09/2002  BM029
'
On Error GoTo SetPaidChequePaymentVbErr

Const strFunctionName As String = "IPaymentProcessingNTTxBO_SetPaidChequePayment"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMNode
Dim xmlAuditRequest As IXMLDOMNode
Dim xmlTempResponse As IXMLDOMNode
Dim xmlChequeAuditNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim strResponse As String
Dim strUserId As String, strAppNo As String, strChequeNo As String, strSeqNo As String, strAppFFNo As String
Dim objTxBO As IPaymentProcessingTxBO
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
    
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
    'Validate APPLICATIONNUMBER attribute is specified
    strAppNo = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strSeqNo = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    'Validate CHEQUENUMBER attribute is specified
    strChequeNo = xmlGetMandatoryAttributeText(xmlNode, "CHEQUENUMBER")
    
    'Get the APPLICATIONFACTFINDNUMBER
    strAppFFNo = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
        
    If Len(Trim$(strAppFFNo)) = 0 Then
        'Get the ApplicationFactFind number
        strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    End If
        
    'SET PAYMENT TO PAID
    If objTxBO Is Nothing Then
        Set objTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    objTxBO.UpdateChequeDisbursement vxmlRequestNode, xmlTempResponse
    'Check for errors
    errCheckXMLResponse xmlTempResponse.xml, True

'    ' CREATE AUDIT RECORD
'    Set xmlDoc = Nothing
'    Set xmlTempResponse = Nothing
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    Set xmlAuditRequest = xmlDoc.createElement("REQUEST")
'    Set xmlChequeAuditNode = xmlDoc.createElement("PAYMENTCHEQUEAUDIT")
'
'    xmlSetAttributeValue xmlChequeAuditNode, "APPLICATIONNUMBER", strAppNo
'    xmlSetAttributeValue xmlChequeAuditNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
'    xmlSetAttributeValue xmlChequeAuditNode, "CAUSERID", strUserId
'    xmlSetAttributeValue xmlChequeAuditNode, "PAYMENTSEQUENCENUMBER", strSeqNo
'    xmlSetAttributeValue xmlChequeAuditNode, "CACHEQUENUMBER", strChequeNo
'    xmlAuditRequest.appendChild xmlChequeAuditNode
'
'    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
'    CreateChequeAuditRecord xmlAuditRequest, xmlTempResponse
'    'Check for errors
'    errCheckXMLResponse xmlTempResponse.xml, True
    'GD BMIDS0922 START
    'UNLOCK APP
    'UnlockApplication strAppNo
    Dim xmlTempReq As IXMLDOMElement
    Dim xmlAppElem As IXMLDOMElement
    Dim xmlAppLockElem As IXMLDOMElement
    
    Set xmlTempReq = xmlGetRequestNode(vxmlRequestNode.cloneNode(False))
    Set xmlAppLockElem = xmlDoc.createElement("APPLICATIONLOCK")
    Set xmlAppElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlAppElem.Text = strAppNo
    
    xmlAppLockElem.appendChild xmlAppElem
    xmlTempReq.appendChild xmlAppLockElem
    
 
    'Check if lock exists first
    Dim nodeResponseFromGetLockData As IXMLDOMElement
    Dim lngErrNo As Long
    'Dim xmlNode As IXMLDOMNode
    Set nodeResponseFromGetLockData = vxmlRequestNode.ownerDocument.createElement("RESPONSE")
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//PAYMENTRECORD")
    IPaymentProcessingNTTxBO_CheckLocks xmlNode, nodeResponseFromGetLockData
        
    lngErrNo = errCheckXMLResponseNode(nodeResponseFromGetLockData, , False)
               
    If lngErrNo = 0 Then 'Lock found - delete it
        'lngErrNo = errGetOmigaErrorNumber(lngErrNo)
        'If errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
            DeleteApplicationLock xmlTempReq
        'End If
                
    End If
   
    
    'GD BMIDS0922 END
    'Set the response to success.
    vxmlResponseNode.setAttribute "TYPE", "SUCCESS"
    
    gobjContext.SetComplete
  
SetPaidChequePaymentExit:
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlAuditRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlChequeAuditNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

    Exit Sub
    
SetPaidChequePaymentVbErr:
    gobjContext.SetAbort
    GoTo SetPaidChequePaymentExit
End Sub

'BM0339 MDC 07/03/2003
Private Sub IPaymentProcessingNTTxBO_UnlockApplication(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UnlockApplicationVbErr

Dim objCBRequestHeaderNode As IXMLDOMNode

Dim strApplicationNumber As String
Dim strPaymentSequenceNumber As String
Dim blnWaitingResponse As Boolean
Dim strBatchNumber As String
Dim strBatchRunNumber As String
Dim strMessage As String

    Const strFunctionName As String = "IPaymentProcessingNTTxBO_UnlockApplication"

    Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "HEADER")
    strApplicationNumber = xmlGetAttributeText(objCBRequestHeaderNode, "APPLICATIONNUMBER")

    'Check if any other payments are pending
    blnWaitingResponse = IsADisbursementAwaitingResponse(strApplicationNumber)
    
    'If no payments pending, unlock the application
    If blnWaitingResponse = False Then
        UnlockApplicationOnly strApplicationNumber, xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")   'BM0339 MDC 06/03/2003
    End If

    'Return success
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    gobjContext.SetComplete

UnlockApplication_Exit:

    Set objCBRequestHeaderNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub

UnlockApplicationVbErr:
    
    'BM0339 MDC 05/03/2003 - Log errors instead of raising them
    If Not objCBRequestHeaderNode Is Nothing Then
        strBatchNumber = xmlGetAttributeText(objCBRequestHeaderNode, "BATCHNUMBER")
        strBatchRunNumber = xmlGetAttributeText(objCBRequestHeaderNode, "BATCHRUNNUMBER")
        strPaymentSequenceNumber = xmlGetAttributeText(objCBRequestHeaderNode, "PAYMENTSEQUENCENUMBER")
    End If
    strMessage = "Warning: An error occurred in " & strFunctionName & " while unlocking the application."
    strMessage = strMessage & vbCrLf & "Application: " & strApplicationNumber
    strMessage = strMessage & vbCrLf & "PaymentSequenceNumber: " & strPaymentSequenceNumber
    strMessage = strMessage & vbCrLf & "Number: " & Err.Number
    strMessage = strMessage & vbCrLf & "Description: " & Err.Description
    strMessage = strMessage & vbCrLf & "Source: " & Err.Source
        
    Call LogWarningMessage(strBatchNumber, strBatchRunNumber, strMessage)
    
    'Clear the error and return success
    Err.Clear
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    'BM0339 MDC 05/03/2003 - End
    
    gobjContext.SetAbort
    GoTo UnlockApplication_Exit


End Sub
'BM0339 MDC 07/03/2003 - End

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub LogBusinessExceptions(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo LogBusinessExceptionsVbErr

    Const strFunctionName As String = "LogBusinessExceptions"

    
    Dim objResultNode As IXMLDOMNode ' this contains the <> nodes that contain the error information
    Dim objHeaderNode As IXMLDOMNode
    Dim objNewRequest As IXMLDOMNode
    Dim objBatchExceptionNode As IXMLDOMNode
    Dim objBatchScheduleBO As BatchScheduleBO
    
    Dim strNodeValue As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchAuditGuid As String
    Dim strResponse As String
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttribute As IXMLDOMAttribute
    
    ' DM     15/03/02    SYS4274 CompleteInterfacing changes.
    Set objHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST/HEADER")
    Set objResultNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/RESULT")
    
    Set objNewRequest = xmlGetRequestNode(vxmlRequestNode)
    
    xmlSetAttributeValue objNewRequest, "OPERATION", "CREATEBATCHEXCEPTION"
    
    Set objBatchExceptionNode = vxmlRequestNode.ownerDocument.createElement("BATCHEXCEPTION")
    
    objNewRequest.appendChild objBatchExceptionNode
            
    Set objBatchScheduleBO = gobjContext.CreateInstance("omBatch.BatchScheduleBO")
            
    strBatchNumber = xmlGetAttributeText(objHeaderNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetAttributeText(objHeaderNode, "BATCHRUNNUMBER")
    strBatchAuditGuid = xmlGetAttributeText(objHeaderNode, "BATCHAUDITGUID")
            
    ' Squirel away the whole of the ResultNode to the database.
    xmlSetAttributeValue objBatchExceptionNode, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue objBatchExceptionNode, "BATCHRUNNUMBER", strBatchRunNumber
    xmlSetAttributeValue objBatchExceptionNode, "BATCHAUDITGUID", strBatchAuditGuid
    xmlSetAttributeValue objBatchExceptionNode, "ERRORINFORMATION", objResultNode.xml
        
    ' Call the ombatch create batch exception method.
    ' Dont want to do anything with the response. Should not fail because of writing to the log.
    strResponse = objBatchScheduleBO.omBatchRequest(objNewRequest.xml)
    
    'gobjContext.SetComplete
    
LogBusinessExceptionsExit:
    Set objResultNode = Nothing
    Set objHeaderNode = Nothing
    Set objNewRequest = Nothing
    Set objBatchExceptionNode = Nothing
    Set objBatchScheduleBO = Nothing
    
    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError strFunctionName, TypeName(Me)
    End If
    
    Exit Sub
LogBusinessExceptionsVbErr:
    ' Even if this fails, writing to the log should not cause the transaction to abort.
    GoTo LogBusinessExceptionsExit
    
End Sub

Private Sub PrintCompletionLetter(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo PrintCompletionLetterVbErr

    Const strFunctionName As String = "PrintCompletionLetter"

PrintCompletionLetterExit:
    
    ' Clean up section
    If Err.Number <> 0 Then
        'Err.Raise Err.Number, Err.Source, Err.Description
        errCheckError strFunctionName, TypeName(Me)
    End If

Exit Sub

PrintCompletionLetterVbErr:
    
    GoTo PrintCompletionLetterExit

End Sub

Private Sub CreateCompletionStatusTask(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo CreateCompletionStatusTaskVbErr

    Const strFunctionName As String = "CreateCompletionStatusTask"

    Dim xmlSystemStatusElement As IXMLDOMElement
    
    ' PSC 20/09/2002 BMIDS00480
    ' Dim xmlStatusElement As IXMLDOMElement
     Dim xmlHeaderElement As IXMLDOMElement
    
    Dim objCaseStageElement As IXMLDOMElement
    Dim objRequestedCaseStageElement As IXMLDOMElement
    Dim objResponseFromGetCurrentStage As FreeThreadedDOMDocument40
    Dim objResponseFromGetApplicationPriority As FreeThreadedDOMDocument40
    Dim xmlApplicationElement As IXMLDOMElement
    Dim objNewRequest As IXMLDOMNode
    Dim objApplicationBO As omApp.ApplicationBO
    Dim objTaskManagementBO As omTM.OmTmBO
    Dim objAppRequest As IXMLDOMNode
    Dim xmlAppElement As IXMLDOMElement
    Dim xmlApplicationPriority As IXMLDOMElement
    
    
    Dim strTaskID As String
    Dim strStageID As String
    Dim strApplicationNumber As String
    Dim strApplicationPriority As String
    Dim strActivityID As String
    Dim strResponse As String
    Dim strTaskDueDateAndTime As String
    Dim strCSSeqNo As String
    
    ' PSC 20/09/2002 BMIDS00480 - Start
    Set xmlSystemStatusElement = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/RESPONSE")
    ' Set xmlStatusElement = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/STATUS")
     Set xmlHeaderElement = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/HEADER")
    ' PSC 20/09/2002 BMIDS00480 - End
    
    ' Get CASEID for the task
'    If xmlGetAttributeText(xmlSystemStatusElement, "VALUE") = "SUCCESS" And _
'        xmlGetAttributeText(xmlStatusElement, "VALUE") <> "ERROR" Then
    If xmlGetAttributeText(xmlSystemStatusElement, "TYPE") = "SUCCESS" Then
        strTaskID = GetGlobalParamString("TMCompletionsSuccess")
        strTaskDueDateAndTime = ""
    Else
        strTaskID = GetGlobalParamString("TMCompletionsFailure")
        strTaskDueDateAndTime = Format$(Now(), "DD/MM/YYYY HH:MM:SS") ' Check what format we need here
    End If
    ' PSC 20/09/2002 BMIDS00480 - End
    
    ' Get STAGEID
    ' This did not return an application number
    strApplicationNumber = xmlGetAttributeText(xmlHeaderElement, "APPLICATIONNUMBER")
    strActivityID = GetGlobalParamString("TMOmigaActivity")
    
    ' Need to get the Application Priority from the application
    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Set objAppRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue objAppRequest, "OPERATION", "GETAPPLICATIONPRIORITY"
    'Set xmlApplicationElement = vxmlRequestNode.ownerDocument.createElement("APPLICATION")
    'xmlSetAttributeValue xmlApplicationElement, "APPLICATIONNUMBER", strApplicationNumber
    'objAppRequest.appendChild xmlApplicationElement
    
    '-------------------------------------------------------------------------------------------
    Set xmlApplicationPriority = vxmlRequestNode.ownerDocument.createElement("APPLICATIONPRIORITY")
    Set xmlApplicationElement = vxmlRequestNode.ownerDocument.createElement("APPLICATIONNUMBER")
    xmlApplicationElement.Text = strApplicationNumber
        
    xmlApplicationPriority.appendChild xmlApplicationElement
    objAppRequest.appendChild xmlApplicationPriority
    
    '-------------------------------------------------------------------------------------------
    strResponse = objApplicationBO.GetApplicationPriority(objAppRequest.xml)
    
    errCheckXMLResponse strResponse, True
    
    ' load the response
    Set objResponseFromGetApplicationPriority = xmlLoad(strResponse, strFunctionName)
    ' Get the application priority
    strApplicationPriority = xmlGetNodeText(objResponseFromGetApplicationPriority, "/RESPONSE/APPLICATIONPRIORITY/APPLICATIONPRIORITYVALUE")
    
    ' set up the request for taskmanagerBO
    Set objNewRequest = xmlGetRequestNode(vxmlRequestNode)

    xmlSetAttributeValue objNewRequest, "OPERATION", "GETCURRENTSTAGE"

    Set objCaseStageElement = vxmlRequestNode.ownerDocument.createElement("CASESTAGE")
    xmlSetAttributeValue objCaseStageElement, "SOURCEAPPLICATION", "Omiga"
    xmlSetAttributeValue objCaseStageElement, "CASEID", strApplicationNumber
    xmlSetAttributeValue objCaseStageElement, "ACTIVITYID", strActivityID
'
'    ' Append the case stage node to the request.
    objNewRequest.appendChild objCaseStageElement
'
    Set objTaskManagementBO = gobjContext.CreateInstance("omTM.OmTmBO")
'
    strResponse = objTaskManagementBO.OmTmRequest(objNewRequest.xml)
'
    errCheckXMLResponse strResponse, True
        
    ' load the response into an xml document
    Set objResponseFromGetCurrentStage = xmlLoad(strResponse, strFunctionName)
    objResponseFromGetCurrentStage.loadXML strResponse
    
    ' get the stageid from the xml response
    Set objRequestedCaseStageElement = xmlGetMandatoryNode(objResponseFromGetCurrentStage.documentElement, "/RESPONSE/CASESTAGE")
    strStageID = xmlGetAttributeText(objRequestedCaseStageElement, "STAGEID")
    strCSSeqNo = xmlGetAttributeText(objRequestedCaseStageElement, "CASESTAGESEQUENCENO")
    
    ' reuse the request
    xmlSetAttributeValue objNewRequest, "OPERATION", "CREATEADHOCCASETASK"
    xmlChangeNodeName objCaseStageElement, "CASESTAGE", "CASETASK"
    xmlSetAttributeValue objCaseStageElement, "ACTIVITYINSTANCE", "1"
    xmlSetAttributeValue objCaseStageElement, "STAGEID", strStageID
    xmlSetAttributeValue objCaseStageElement, "TASKID", strTaskID
    
    'PSC 19/09/2002 SYS5238
    'xmlSetAttributeValue objCaseStageElement, "CSSEQNO", strCSSeqNo
    xmlSetAttributeValue objCaseStageElement, "TASKDUEDATEANDTIME", strTaskDueDateAndTime
    
    'PSC 19/09/2002 SYS5238 - Start
    xmlSetAttributeValue objNewRequest, "CHANNELID", xmlGetAttributeText(xmlHeaderElement, "CHANNELID")
    xmlSetAttributeValue objCaseStageElement, "CASESTAGESEQUENCENO", strCSSeqNo
    'PSC 19/09/2002 SYS5238 - End
    
    ' Add an Application node to the request xml.
    Set xmlAppElement = vxmlRequestNode.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlAppElement, "APPLICATIONPRIORITY", strApplicationPriority
    
    objNewRequest.appendChild xmlAppElement
    
    strResponse = objTaskManagementBO.OmTmRequest(objNewRequest.xml)

    errCheckXMLResponse strResponse, True
    
CreateCompletionStatusTaskExit:
    
    ' Clean up section
    Set xmlSystemStatusElement = Nothing
    ' PSC 20/09/2002 BMIDS00480
    'Set xmlStatusElement = Nothing
    Set xmlHeaderElement = Nothing
    Set objCaseStageElement = Nothing
    Set objRequestedCaseStageElement = Nothing
    Set objResponseFromGetCurrentStage = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

Exit Sub

CreateCompletionStatusTaskVbErr:
    
    GoTo CreateCompletionStatusTaskExit

End Sub


Private Sub UpdateAdminSubLedgerCodes(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'DM     07/01/02    SYS3185 - Implement CompleteInterfacing
'DM     27/03/02    SYS4313 CompleteInterfacing changes
'------------------------------------------------------------------------------------------
On Error GoTo UpdateAdminSubLedgerCodesVbErr

    Const strFunctionName As String = "UpdateAdminSubLedgerCodes"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim objMortgageNode As IXMLDOMNode
    
    ' RF 11/01/02
    'Dim objNewRequest As DomDocument
    Dim nodeNewRequest As IXMLDOMNode
    
    Dim objMortgageSubQuoteBO As omCM.MortgageSubQuoteBO
    Dim objCBRequestHeader As IXMLDOMElement
        
    Dim strResponse As String
    
    ' Set up the request to be sent to the UpdateLoanComponent method
    
    Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
    
    Set objCBRequestHeader = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST/HEADER")
    
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlApplicationNumber As IXMLDOMElement
    Dim xmlApplicationFactFindNumber As IXMLDOMElement
    Dim xmlMortgageSubQuoteNumber As IXMLDOMElement
    Dim xmlLoanComponentSequenceNumber As IXMLDOMElement
    Dim xmlAdminSubLedgerCode As IXMLDOMElement
    
    'SG 28/05/02 SYS4767 START
    Dim xmlAdminSubLedgerCodeCI As IXMLDOMElement   '++ CR33
    Dim xmlAdminSubLedgerCodeIO As IXMLDOMElement   '++ CR33
    Dim strPartAndPartInd As String                 '++ CR33
    'SG 28/05/02 SYS4767 END
    
    Set xmlLoanComponent = vxmlRequestNode.ownerDocument.createElement("LOANCOMPONENT")
    Set xmlApplicationNumber = vxmlRequestNode.ownerDocument.createElement("APPLICATIONNUMBER")
    Set xmlApplicationFactFindNumber = vxmlRequestNode.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
    Set xmlMortgageSubQuoteNumber = vxmlRequestNode.ownerDocument.createElement("MORTGAGESUBQUOTENUMBER")
    Set xmlLoanComponentSequenceNumber = vxmlRequestNode.ownerDocument.createElement("LOANCOMPONENTSEQUENCENUMBER")
    Set xmlAdminSubLedgerCode = vxmlRequestNode.ownerDocument.createElement("ADMINSUBLEDGERCODE")
    
    'SG 28/05/02 SYS4767 START
    Set xmlAdminSubLedgerCodeCI = vxmlRequestNode.ownerDocument.createElement("ADMINSUBLEDGERCODECI")   '++ CR33
    Set xmlAdminSubLedgerCodeIO = vxmlRequestNode.ownerDocument.createElement("ADMINSUBLEDGERCODEIO")    '++ CR33
    'SG 28/05/02 SYS4767 END
    
    xmlLoanComponent.appendChild xmlApplicationNumber
    xmlLoanComponent.appendChild xmlApplicationFactFindNumber
    xmlLoanComponent.appendChild xmlMortgageSubQuoteNumber
    xmlLoanComponent.appendChild xmlLoanComponentSequenceNumber
    xmlLoanComponent.appendChild xmlAdminSubLedgerCode
    
    ' RF 11/01/02 Made the following 3 items mandatory
    xmlApplicationNumber.Text = _
        xmlGetMandatoryAttributeText(objCBRequestHeader, "APPLICATIONNUMBER")
    xmlApplicationFactFindNumber.Text = _
        xmlGetMandatoryAttributeText(objCBRequestHeader, "APPLICATIONFACTFINDNUMBER")
    xmlMortgageSubQuoteNumber.Text = _
        xmlGetMandatoryAttributeText(objCBRequestHeader, "MORTGAGESUBQUOTENUMBER")
    
    nodeNewRequest.appendChild xmlLoanComponent
    
    ' May need to test for CBREQUEST XML Errors
    
    Set objMortgageNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST/MORTGAGE")
    
    Set xmlNodeList = objMortgageNode.selectNodes("COMPONENT")
    
    Set objMortgageSubQuoteBO = gobjContext.CreateInstance("omCM.MortgageSubQuoteBO")
    ' DM 09/04/02 SYS4377 Testing for nothing from the CBA
    Dim strTestSubLedgerCode As String
    
    For Each xmlNode In xmlNodeList
        ' update the loan component.
       
        ' DM 09/04/02 SYS4377 Testing for nothing from the CBA
        strTestSubLedgerCode = xmlGetAttributeText(xmlNode, "RETURNSUBLEDGERCODE")
        
        ' DM 09/04/02 SYS4377 Testing for nothing from the CBA
        If strTestSubLedgerCode <> "EMPTY" Then
        
            xmlLoanComponentSequenceNumber.Text = xmlGetAttributeText(xmlNode, "LOANCOMPONENTNUMBER")
        
            'SG 28/05/02 SYS4767 START
            'CR33 SA Check for Part & Part Indicator ++ {
            'xmlAdminSubLedgerCode.Text = xmlGetAttributeText(xmlNode, "RETURNSUBLEDGERCODE")   '-- CR33
            strPartAndPartInd = xmlGetAttributeText(xmlNode, "PARTANDPARTIND")
            Select Case strPartAndPartInd
                Case "CI"
                    xmlAdminSubLedgerCodeCI.Text = xmlGetAttributeText(xmlNode, "RETURNSUBLEDGERCODE")
                Case "IO"
                    xmlAdminSubLedgerCodeIO.Text = xmlGetAttributeText(xmlNode, "RETURNSUBLEDGERCODE")
                Case Else
                    xmlAdminSubLedgerCode.Text = xmlGetAttributeText(xmlNode, "RETURNSUBLEDGERCODE")
            End Select
            'CR33 SA ++ end }
            'SG 28/05/02 SYS4767 END
        
            strResponse = objMortgageSubQuoteBO.UpdateLoanComponent(nodeNewRequest.xml)
        
            ' If there is an error in one of the components do we try any subsequent components
            ' or do we quit out of the loop?
            errCheckXMLResponse strResponse, True
        ' DM 09/04/02 SYS4377 Testing for nothing from the CBA
        End If
        
    Next

UpdateAdminSubLedgerCodesExit:
    
    ' Clean up section
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set objMortgageNode = Nothing
    Set nodeNewRequest = Nothing
    Set objMortgageSubQuoteBO = Nothing
    Set objCBRequestHeader = Nothing
        
    Set xmlLoanComponent = Nothing
    Set xmlApplicationNumber = Nothing
    Set xmlApplicationFactFindNumber = Nothing
    Set xmlMortgageSubQuoteNumber = Nothing
    Set xmlLoanComponentSequenceNumber = Nothing
    Set xmlAdminSubLedgerCode = Nothing
    
    ' RF 11/01/02 Ensure Err.Source and warnings are handled correctly
'    If Err.Number <> 0 Then
'        Err.Raise Err.Number, Err.Source, Err.Description
'    End If
    errCheckError strFunctionName, TypeName(Me)

Exit Sub

UpdateAdminSubLedgerCodesVbErr:

    gobjContext.SetAbort
    
    GoTo UpdateAdminSubLedgerCodesExit

End Sub

Private Sub UpdateDisbursementStatus( _
    ByVal vxmlRequestNode As IXMLDOMNode, strValidationType As String)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' history:
'DM     07/01/02    SYS3185 Implement CompleteInterfacing
'RF     27/02/02    SYS4193 Applied corrections previously identified by DM but not checked in.
'                           Does not include fixes made within TemporaryCompleteInterfacing.
'DM     27/03/02    SYS4313 CompleteInterfacing changes
' DM 11/04/2002 SYS4385 Create the request correctly
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDisbursementStatusVbErr

    Const strFunctionName As String = "UpdateDisbursementStatus"

    Dim objXMLRequest As IXMLDOMElement
    Dim objXMLPaymentRecordElement As IXMLDOMElement
    Dim objXMLDisbursementPaymentElement As IXMLDOMElement
    Dim objXMLNode As IXMLDOMNode
    Dim objCBRequestHeaderNode As IXMLDOMNode
    
    ' PSC 14/11/2005 MAR520 - Start
    Dim objDisbursementPaymentIn As IXMLDOMNode
    Dim strFirstPaymentDate As String
    Dim strFirstPayment As String
    Dim strRegularPayment As String
    ' PSC 14/11/2005 MAR520 - End
    
    Dim strValueId As String
    Dim strAppNumber As String
    Dim strPaymentSeqNumber As String
    Dim objPaymentProcessingBO As PaymentProcessingBO
    Dim strResponse As String
    
    strValueId = GetFirstComboValueId("PaymentStatus", strValidationType)
    ' Get the payment sequence number

    ' PSC 20/09/2002 BMIDS00480
    ' Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST/HEADER")
    Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/HEADER")
   
    strAppNumber = xmlGetAttributeText(objCBRequestHeaderNode, "APPLICATIONNUMBER")
     
    
    ' PSC 20/09/2002 BMIDS00480
    strPaymentSeqNumber = xmlGetAttributeText(objCBRequestHeaderNode, "PAYMENTSEQUENCENUMBER")
    ' DM 11/04/2002 SYS4385 Create the request correctly
    Set objXMLRequest = xmlGetRequestNode(vxmlRequestNode)
    ' END DM 11/04/2002 SYS4385 Create the request correctly
    xmlSetAttributeValue objXMLRequest, "OPERATION", "UPDATEDISBURSEMENT"
    
    Set objXMLPaymentRecordElement = vxmlRequestNode.ownerDocument.createElement("PAYMENTRECORD")
    
    xmlSetAttributeValue objXMLPaymentRecordElement, "APPLICATIONNUMBER", strAppNumber
    xmlSetAttributeValue objXMLPaymentRecordElement, "PAYMENTSEQUENCENUMBER", strPaymentSeqNumber
    
    Set objXMLNode = objXMLRequest.appendChild(objXMLPaymentRecordElement)
    
    Set objXMLDisbursementPaymentElement = vxmlRequestNode.ownerDocument.createElement("DISBURSEMENTPAYMENT")
    
    xmlSetAttributeValue objXMLDisbursementPaymentElement, "PAYMENTSTATUS", strValueId
    
    ' PSC 14/11/2005 MAR520 - Start
    ' Successful
    If strValidationType = "I" Then
        Set objDisbursementPaymentIn = xmlGetMandatoryNode(vxmlRequestNode, "RESPONSE/DISBURSEMENTPAYMENT")
        strFirstPaymentDate = xmlGetMandatoryAttributeText(objDisbursementPaymentIn, "FIRSTPAYMENTDATE")
        strFirstPayment = xmlGetMandatoryAttributeText(objDisbursementPaymentIn, "FIRSTPAYMENT")
        strRegularPayment = xmlGetMandatoryAttributeText(objDisbursementPaymentIn, "REGULARPAYMENT")
        xmlSetAttributeValue objXMLDisbursementPaymentElement, "FIRSTPAYMENTDATE", strFirstPaymentDate
        xmlSetAttributeValue objXMLDisbursementPaymentElement, "FIRSTPAYMENT", strFirstPayment
        xmlSetAttributeValue objXMLDisbursementPaymentElement, "REGULARPAYMENT", strRegularPayment
    End If
    ' PSC 14/11/2005 MAR520 - End
    
    objXMLNode.appendChild objXMLDisbursementPaymentElement
    
    Set objPaymentProcessingBO = gobjContext.CreateInstance("omPayProc.PaymentProcessingBO")
    
    If Not objPaymentProcessingBO Is Nothing Then
        
        strResponse = objPaymentProcessingBO.omPayProcRequest(objXMLRequest.xml)
        
        errCheckXMLResponse strResponse, True

        
    End If
    
UpdateDisbursementStatusExit:
    
    ' Clean up section
    Set objXMLRequest = Nothing
    Set objXMLPaymentRecordElement = Nothing
    Set objXMLDisbursementPaymentElement = Nothing
    Set objXMLNode = Nothing
    Set objCBRequestHeaderNode = Nothing
    Set objDisbursementPaymentIn = Nothing ' PSC 14/11/2005 MAR520
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

Exit Sub

UpdateDisbursementStatusVbErr:
    ' May have to set abort in here.
    
    GoTo UpdateDisbursementStatusExit
    
End Sub


' PSC 19/09/2002 SYS4863 - New method
Private Sub UpdateFeeCompletionIndicators( _
    ByVal vxmlRequestNode As IXMLDOMNode, strValidationType As String)
' header ----------------------------------------------------------------------------------
' description: Update the FEEPAYMENT records
' pass: Response from CBA, status to set to
' return: Nothing
' history:
'
' DS 09/07/2002 SYS4863 Created
'------------------------------------------------------------------------------------------
On Error GoTo UpdateFeeCompletionIndicatorsVbErr

    Const strFunctionName As String = "UpdateFeeCompletionIndicators"

    Dim objXMLRequest As IXMLDOMElement
    Dim objXMLResponse As IXMLDOMElement
    Dim objXMLNode As IXMLDOMNode
    Dim strValueId As String
    Dim strAppNumber As String
    Dim strPaymentSeqNumber As String
    Dim objPaymentProcessingDO As IPAymentProcessingDO
    Dim strResponse As String
    Dim objDomDoc As New FreeThreadedDOMDocument40
    objDomDoc.validateOnParse = False
    objDomDoc.setProperty "NewParser", True
    Dim objCBRequestHeaderNode As IXMLDOMNode
    
    strValueId = GetFirstComboValueId("CompletionIndicator", strValidationType)
    ' Get the payment sequence number
    
    ' PSC 20/09/2002 BMIDS00480
    'Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/BODY/CBREQUEST/RESPONSE/CBREQUEST/HEADER")
    Set objCBRequestHeaderNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/HEADER")
    
    strAppNumber = xmlGetAttributeText(objCBRequestHeaderNode, "APPLICATIONNUMBER")
    
    ' PSC 20/09/2002 BMIDS00480
    strPaymentSeqNumber = xmlGetAttributeText(objCBRequestHeaderNode, "PAYMENTSEQUENCENUMBER")
    
    'Get all the fees with PAYMENTSEQUENCENUMBER >= the advance number
    
    Set objPaymentProcessingDO = gobjContext.CreateInstance("omPayProc.PaymentProcessingDO")
    objDomDoc.loadXML ("<REQUEST APPLICATIONNUMBER=""" & strAppNumber & """></REQUEST>")
    Set objXMLRequest = objDomDoc.selectSingleNode("REQUEST")
    objDomDoc.loadXML ("<RESPONSE/>")
    Set objXMLResponse = objDomDoc.selectSingleNode("RESPONSE")
    
    objPaymentProcessingDO.FindFeePaymentList objXMLRequest, objXMLResponse

    errCheckXMLResponse objXMLResponse.xml, True
    
    objDomDoc.loadXML ("<REQUEST/>")
    Set objXMLRequest = objDomDoc.selectSingleNode("REQUEST")
    Dim intPayAeqNum As Integer
    
    'set the completionindicator on all records in the XML and append the nodes to the next request
    Dim xmlTempNode As IXMLDOMElement
    Dim xmlTempNodeList As IXMLDOMNodeList
    Dim xmlTempAttr As IXMLDOMAttribute
    Dim intPaySeqNum As Integer
    
    Set xmlTempNodeList = objXMLResponse.selectNodes("FEEPAYMENT")
    For Each xmlTempNode In xmlTempNodeList
        intPaySeqNum = CInt(xmlTempNode.getAttribute("PAYMENTSEQUENCENUMBER"))
        If intPaySeqNum <= CInt(strPaymentSeqNumber) Then
            xmlTempNode.removeAttribute ("COMPLETIONINDICATOR")
            xmlTempNode.setAttribute "COMPLETIONINDICATOR", strValueId
            objXMLRequest.appendChild xmlTempNode
        End If
    Next
    
    'Call the Update method
    objPaymentProcessingDO.UpdateFeePaymentList objXMLRequest, objXMLResponse
    
    errCheckXMLResponse objXMLResponse.xml, True
    
UpdateFeeCompletionIndicatorsExit:
    
    ' Clean up section
    Set objXMLRequest = Nothing
    Set objXMLResponse = Nothing
    Set objXMLNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempNodeList = Nothing
    Set xmlTempAttr = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

Exit Sub

UpdateFeeCompletionIndicatorsVbErr:
    ' May have to set abort in here.
    
    GoTo UpdateFeeCompletionIndicatorsExit
    
End Sub

Private Function BACSPayment(ByVal vxmlPaymentNode As IXMLDOMNode) As IXMLDOMNode

On Error GoTo BACSPaymentExit:

    Const strFunctionName As String = "BACSPayment"

    Dim objBACSInterfaceBO As Object
    Dim xmlBACSRequest As IXMLDOMNode
    Dim xmlBACSRequestDoc As FreeThreadedDOMDocument40
    Dim xmlBACSResponseDoc As FreeThreadedDOMDocument40
    Dim xmlBACSNode As IXMLDOMNode
    Dim xmlPaymentCopy As IXMLDOMNode   ' PSC 18/09/2002 SYS5077
    Dim strResponse As String
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Dim strAppFFNo As String
    Dim strAppNo As String
    ' PSC 08/10/2002 BMIDS00595 - End
                        
    ' PSC 08/10/2002 BMIDS00595 - Start
    strAppNo = xmlGetMandatoryAttributeText(vxmlPaymentNode, "APPLICATIONNUMBER")
    strAppFFNo = GetFactFindNumberForApplication(strAppNo)
    ' PSC 08/10/2002 BMIDS00595 - End
    
    'Create Request
    Set xmlBACSRequestDoc = New FreeThreadedDOMDocument40
    xmlBACSRequestDoc.validateOnParse = False
    xmlBACSRequestDoc.setProperty "NewParser", True
    Set xmlBACSRequest = xmlBACSRequestDoc.createElement("REQUEST")
    xmlSetAttributeValue xmlBACSRequest, "OPERATION", "BUILDBACSINTERFACE"
    xmlSetAttributeValue xmlBACSRequest, "APPLICATIONNUMBER", strAppNo
    xmlSetAttributeValue xmlBACSRequest, "PAYMENTSEQUENCENUMBER", xmlGetMandatoryAttributeText(vxmlPaymentNode, "PAYMENTSEQUENCENUMBER")
    
    ' PSC 18/09/2002 SYS5077 - Start
    Set xmlPaymentCopy = vxmlPaymentNode.cloneNode(True)
    
    ' PSC 08/10/2002 BMIDS00595
    xmlSetAttributeValue xmlPaymentCopy, "APPLICATIONFACTFINDNUMBER", strAppFFNo
    xmlBACSRequest.appendChild xmlPaymentCopy
    ' PSC 18/09/2002 SYS5077 - End
    
    xmlBACSRequestDoc.appendChild xmlBACSRequest
                
    'Create Response
    Set xmlBACSResponseDoc = New FreeThreadedDOMDocument40
    xmlBACSResponseDoc.validateOnParse = False
    xmlBACSResponseDoc.setProperty "NewParser", True
    
    'Call component
    Set objBACSInterfaceBO = gobjContext.CreateInstance("omBACS.omBACSInterfaceBO")
    strResponse = objBACSInterfaceBO.OmRequest(xmlBACSRequest.xml)
    xmlBACSResponseDoc.loadXML strResponse
    
    Set BACSPayment = xmlBACSResponseDoc
    
    gobjContext.SetComplete
    
BACSPaymentExit:
    
    ' Clean up section
    If Not objBACSInterfaceBO Is Nothing Then Set objBACSInterfaceBO = Nothing
    If Not xmlBACSRequest Is Nothing Then Set xmlBACSRequest = Nothing
    If Not xmlBACSResponseDoc Is Nothing Then Set xmlBACSResponseDoc = Nothing
    If Not xmlBACSRequestDoc Is Nothing Then Set xmlBACSRequestDoc = Nothing
    If Not xmlBACSNode Is Nothing Then Set xmlBACSNode = Nothing
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
    Exit Function

BACSPaymentErr:
    gobjContext.SetAbort
    GoTo BACSPaymentExit

End Function


Private Sub ProcessCHAPSPayment(ByVal xmlPayment As IXMLDOMElement, ByVal sUnitID As String, ByRef xmlResponseNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Writes or appends payments to a CHAPS payment file. The path is obtained from a global
'   parameter.
'
' pass:
'   xmlPayment
'       <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER=""/>
'   sUnitID
'   vxmlResponseElement
'       XML standard RESPONSE node.
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    
    Dim iPos As Integer
    Dim sBody As String
    Dim sTrailer As String
    Dim sFilename As String
    Dim sFolderPath As String
    Dim xmlOut As FreeThreadedDOMDocument40
' BMIDS646   Dim objFile As TextStream
    Dim iRecordCount As Integer
    Dim vdTotalAmount As Variant
' BMIDS646   Dim objFileSystem As FileSystemObject
    Dim objGlobalParam As GlobalParameterDO
    
    Dim sFailureReason As String                        ' PSC 19/09/2002 SYS5551
    Dim xmlErrorNode As IXMLDOMNode
    Dim xmlDescNode As IXMLDOMElement     ' PSC 19/09/2002 SYS5551
    'BMIDS646
    Dim objCRM As omCRMWrapper
    Dim nMaxRetries As Integer
    Dim nSleepMilliSecs As Integer
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlRequest As IXMLDOMNode
    Dim strResponse As String
    Dim strText As String
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    On Error GoTo ProcessCHAPSPaymentExit
    
    Const strFunctionName As String = "ProcessCHAPSPayment"
    
    'Create a response.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    Set xmlResponseNode = xmlOut.createElement("RESPONSE")
    
    'Get the path and filename from the 'CHAPSPathAndFilename' global parameter.
    Set objGlobalParam = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    sFilename = objGlobalParam.GetCurrentParameterByType("CHAPSPathAndFilename", "STRING")
    'BMIDS646 New Globals
    nMaxRetries = GetMandatoryGlobalParamAmount("LockTextFileRetries")
    nSleepMilliSecs = GetMandatoryGlobalParamAmount("LockTextFileSleep")
    
    'BMIDS646 Create path if it doesn't exist.
'    Set objFileSystem = New FileSystemObject
'    iPos = InStrRev(sFilename, "\", , vbTextCompare)
'
'    If iPos > 0 Then
'        sFolderPath = Left$(sFilename, iPos)
'
'        If objFileSystem.FolderExists(sFolderPath) = False Then
'            objFileSystem.CreateFolder sFolderPath
'        End If
'    End If
    
    'Get header AND body from any existing file, otherwise create one.
    sBody = CHAPSOpenOrCreateBody(iRecordCount, vdTotalAmount, sFilename)
    
    'Append any additional payment record.
    '-------------------------------------
    'SG 30/05/02 SYS4767 START
    'SG 20/05/02 MSMS0083 - CHAPSProcessNewPayment now returns boolean so we can validate the body
    'OLD CODE:
    'sBody = sBody & CHAPSProcessNewPayment(sUnitID, iRecordCount, vdTotalAmount, xmlPayment)
    'NEW CODE:
'    If CHAPSProcessNewPayment(sUnitID, iRecordCount, vdTotalAmount, xmlPayment) = False Then
     
     ' PSC 19/09/2002 SYS5551
     If CHAPSProcessNewPayment(sUnitID, iRecordCount, vdTotalAmount, xmlPayment, sFailureReason) = False Then

        'invalid
        xmlResponseNode.setAttribute "TYPE", "ERROR"
        
        'PSC 19/09/2002 SYS5551 - Start
        'Now add reason for failure to be written to BatchAudit table
        Set xmlErrorNode = xmlOut.createElement("ERROR")
        Set xmlDescNode = xmlOut.createElement("DESCRIPTION")
        xmlDescNode.Text = sFailureReason
        xmlErrorNode.appendChild xmlDescNode
        xmlResponseNode.appendChild xmlErrorNode
        ' PSC 19/09/2002 SYS5551 - End
        
        GoTo ProcessCHAPSPaymentExit
    End If
    sBody = sBody & mstrCHAPSPaymentBody
    'End SG 20/05/02 MSMS0083
    'SG 30/05/02 SYS4767 END
    
    'Create a trailer record from the new totals.
    sTrailer = CHAPSCreateTrailer(iRecordCount, vdTotalAmount)
'BMIDS646    'Attempt to open the existing file (otherwise create one).
'    Set objFile = objFileSystem.OpenTextFile(sFilename, ForWriting, True, TristateUseDefault)
'
'    'Output the CHAPS records.
'    objFile.Write sBody
'    objFile.Write sTrailer
'
'    'Close the file.
'    objFile.Close
    
    'BMIDS646
    nMaxRetries = GetMandatoryGlobalParamAmount("LockTextFileRetries")
    nSleepMilliSecs = GetMandatoryGlobalParamAmount("LockTextFileSleep")
    
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    Set xmlElement = xmlDoc.createElement("TEXTOUT")
    xmlElement.setAttribute "PATHANDFILENAME", sFilename
    xmlElement.setAttribute "MAXRETRIES", nMaxRetries
    xmlElement.setAttribute "SLEEPMILLISECS", nSleepMilliSecs
    'WRITEMODE is either "Append" or OverWrite"
    xmlElement.setAttribute "WRITEMODE", "OverWrite"
    xmlRequest.appendChild xmlElement
    
    Set objCRM = gobjContext.CreateInstance(gstrCRM_COMPONENT & ".omCRMWrapper")
    strResponse = objCRM.AddText(xmlRequest.xml, sBody & sTrailer)
                                                
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    'Set the response to success.
    xmlResponseNode.setAttribute "TYPE", "SUCCESS"
    
ProcessCHAPSPaymentExit:
    'BMIDS646
    Set objCRM = Nothing
    Set xmlDoc = Nothing
    Set xmlElement = Nothing
    Set xmlRequest = Nothing
' BMIDS646   Set objFile = Nothing
' BMIDS646   Set objFileSystem = Nothing
    Set objGlobalParam = Nothing
    
    Set xmlErrorNode = Nothing  ' PSC 19/09/2002 SYS5551
    Set xmlDescNode = Nothing   ' PSC 19/09/2002 SYS5551

    errCheckError strFunctionName, TypeName(Me)
End Sub


Private Function CHAPSOpenOrCreateBody(ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, sFilename As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Opens and reads the header from the file specified. If the file is empty a new header
'   is created. If the file is not empty but doesn't start with a header, then an error
'   is raised (i.e. if the file format is unknown).
'
'   Any existing transaction records are also returned.
'
' pass:
'   iRecordCount  - Counter for the number of transaction records.
'   vdTotalAmount - Sum of all payment amounts.
'   sFilename     - Path and filename to use.
' return:
'   String - Header and payment records.
' exceptions:
'   Invalid file.
'------------------------------------------------------------------------------------------

    Dim sBuffer As String
    Dim sHeader As String
    Dim sTrailer As String
    Dim clsFile As TextStream
    Dim clsFileSystem As FileSystemObject

    Const strFunctionName As String = "CHAPSOpenOrCreateBody"

    On Error GoTo CHAPSOpenOrCreateBodyExit
    
    'Attempt to open the existing file (otherwise we'll create one).
    Set clsFileSystem = New FileSystemObject
    Set clsFile = clsFileSystem.OpenTextFile(sFilename, ForReading, True, TristateUseDefault)
    
    If clsFile.AtEndOfStream = False Then
        'Read the contents of the file into a buffer.
        sBuffer = clsFile.ReadAll()
    End If
    
    'Close the file.
    clsFile.Close
    
    'If the buffer is empty, create a header and return it.
    If Len(sBuffer) = 0 Then
        sBuffer = CHAPS_HEADER_PREFIX & CHAPS_SEP & Format$(Now, "dd/mm/yy") & CHAPS_RECORD_SUFFIX & vbCrLf
        iRecordCount = 0
        vdTotalAmount = CDec(0)
    
    'Otherwise, verify the buffer contains a header and a trailer record.
    Else
        sHeader = Left$(sBuffer, CHAPS_HEADER_LENGTH)
        sTrailer = Right$(sBuffer, CHAPS_TRAILER_LENGTH)
        
        'If the header begins with a 'H' and ends with a '#' then we'll assume its valid.
        If Not (Left$(sHeader, 1) = CHAPS_HEADER_PREFIX And Right$(sHeader, 3) = (CHAPS_RECORD_SUFFIX & vbCrLf)) Then
            'Throw an invalid file format error.
            errThrowError strFunctionName, 6004
        End If
        
        'If the trailer begins with a 'T' and ends with a '#' then we'll assume its valid.
        If Not (Left$(sTrailer, 1) = CHAPS_TRAILER_PREFIX And Right$(sTrailer, 3) = (CHAPS_RECORD_SUFFIX & vbCrLf)) Then
            'Throw an invalid file format error.
            errThrowError strFunctionName, 6004
        End If
    
        'Get the current record count and total amount from the trailer.
        iRecordCount = CInt(Mid$(sTrailer, 3, 3))
        vdTotalAmount = CDec(Mid$(sTrailer, 7, 12))
    
        'Remove the trailer record from the buffer.
        sBuffer = Left$(sBuffer, (Len(sBuffer) - CHAPS_TRAILER_LENGTH))
    End If
        
    'Return the file (header and payment records) to the caller.
    CHAPSOpenOrCreateBody = sBuffer
    
    Exit Function
    
CHAPSOpenOrCreateBodyExit:
    errCheckError strFunctionName, TypeName(Me)
End Function

' PSC 19/09/2002 SYS5551 Add extra return parameter to determine if Account number or Sort code is invalid
Private Function CHAPSProcessNewPayment(ByVal sUnitID As String, ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, ByRef xmlPayment As IXMLDOMNode, Optional ByRef rstrReasonForFailure As String) As Boolean
'Private Function CHAPSProcessNewPayment(ByVal sUnitID As String, ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, ByRef xmlPayment As IXMLDOMNode) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Processes the payment request XML and creates a CHAPS record for each. These records
'   are then appended to the return string.
'
' pass:
'   sUnitID
'   iRecordCount  - Current number of payment records (+1) in the body so far.
'   vdTotalAmount - Total of all existing payment records.
'   xmlPayment
'       <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER=""/>
' return:
'   String - CHAPS transaction records.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------

    Dim sBuffer As String
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlPaymentDetails As IXMLDOMElement
    Dim xmlMortgageLender As IXMLDOMElement
    Dim xmlMortgageProduct As IXMLDOMElement
    Dim xmlComponentPayment As IXMLDOMElement
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Dim xmlDisb As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objMortSQ As MortgageSubQuoteBO
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim strIRValueId As String
    Dim strPaymentType As String
    Dim strResponse As String
    Dim strAppFFNo As String
    Dim strAppNo As String
    ' PSC 08/10/2002 BMIDS00595 - End
    
    ' PSC 13/03/2007 EP2_1347 - Start
    Dim blnIsIncentive As Boolean
    Dim blnIsValuationRefund As Boolean
    ' PSC 13/03/2007 EP2_1347 - End

    Const strFunctionName As String = "CHAPSProcessNewPayment"

    On Error GoTo CHAPSProcessNewPaymentExit
    
    'SG 30/05/02 SYS4767
    'SG 20/05/02 MSMS0083
    CHAPSProcessNewPayment = True
    
    'Get the full Payment Details.
    Set xmlPaymentDetails = GetPaymentDetails(xmlPayment)
        
    'Get the Customers for the application.
    Set xmlCustomerList = GetCustomerList(xmlPayment)
           
    'Get the values to obtain the loan component.
    Set xmlComponentPayment = xmlPaymentDetails.selectSingleNode("//LOANCOMPONENTPAYMENT")
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Set xmlDisb = xmlGetMandatoryNode(xmlPayment, "DISBURSEMENTPAYMENT")
    strPaymentType = xmlGetMandatoryAttributeText(xmlDisb, "PAYMENTTYPE")
    ' PSC 08/10/2002 BMIDS00595 - End
    
    ' PSC 13/03/2007 EP2_1347 - Start
    blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strPaymentType), "C")
    blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strPaymentType), "VALREFUND")
    ' PSC 13/03/2007 EP2_1347 - End

    ' Loan component payment will be nothing for Incentive Releases
    ' PSC 13/03/2007 EP2_1347
    If xmlComponentPayment Is Nothing And blnIsIncentive = False And blnIsValuationRefund = False Then
        sBuffer = ""
    Else
        ' PSC 08/10/2002 BMIDS00595 - Start
        ' PSC 13/03/2007 EP2_1347
        If blnIsIncentive Or blnIsValuationRefund Then
            
            ' Get the accepted list of components
            strAppNo = xmlGetMandatoryAttributeText(xmlDisb, "APPLICATIONNUMBER")
            strAppFFNo = GetFactFindNumberForApplication(strAppNo)
            
            Set xmlTempRequest = xmlPayment.ownerDocument.createElement("REQUEST")
            Set xmlRoot = xmlPayment.ownerDocument.createElement("FINDACTIVELOANCOMPONENTS")
            xmlTempRequest.appendChild xmlRoot
            Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONNUMBER")
            xmlRoot.appendChild xmlNode
            xmlNode.Text = strAppNo
            Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
            xmlRoot.appendChild xmlNode
            xmlNode.Text = strAppFFNo

            Set objMortSQ = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
            
            strResponse = objMortSQ.FindActiveLoanComponentsList(xmlTempRequest.xml)
            Set xmlResponse = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlResponse.documentElement, , True
            
            ' Set the loan component to the first one as lender details will be the same
            ' fo all of them
            Set xmlLoanComponent = xmlGetMandatoryNode(xmlResponse, ".//LOANCOMPONENT")
        Else
            'Get the Loan Component details (to get to the lender details).
            Set xmlLoanComponent = GetLoanComponent(xmlComponentPayment)
        End If
        ' PSC 08/10/2002 BMIDS00595 - end
        
        'Get the Mortgage Product details (to get to the lender details).
        Set xmlMortgageProduct = GetMortgageProduct(xmlLoanComponent)
           
        'Get the Lender account details.
        Set xmlMortgageLender = GetMortgageLender(xmlMortgageProduct)
           
        'SG 30/05/02 SYS4767 START
        'SG 20/05/02 MSMS0083 START
        If ValidateSortCode(xmlGetNodeText(xmlMortgageLender, "//BARCLAYSCHAPSSORTCODE")) = False Then
            ' PSC 19/09/2002 SYS5551
            rstrReasonForFailure = "Lender Account Details - missing or invalid sort code."
            CHAPSProcessNewPayment = False
            GoTo CHAPSProcessNewPaymentExit
        End If
        
        If ValidateAccountNo(xmlGetNodeText(xmlMortgageLender, "//BARCLAYSCHAPSACCOUNTNUMBER")) = False Then
            ' PSC 19/09/2002 SYS5551
            rstrReasonForFailure = "Lender Account Details - missing or invalid account number." 'SA 06/09/02
            CHAPSProcessNewPayment = False
            GoTo CHAPSProcessNewPaymentExit
        End If
        'SG 20/05/02 END
        'SG 30/05/02 SYS4767 END
           
        'If the record conut equals or exceeds 999 then we must not add the record.
        If iRecordCount >= 999 Then
            errThrowError strFunctionName, 6005, "CHAPS"
        Else
            'Increment the record count.
            iRecordCount = iRecordCount + 1
        
            'Create a payment/transaction record.
            sBuffer = sBuffer & CHAPSCreatePaymentRecord(sUnitID, vdTotalAmount, xmlPaymentDetails, xmlCustomerList, xmlMortgageLender)
        End If
    End If
    
    'SG 30/05/02 SYS4767
    'SG 20/05/02 MSMS0083
    'CHAPSProcessNewPayment = sBuffer
    mstrCHAPSPaymentBody = sBuffer

CHAPSProcessNewPaymentExit:
    Set xmlPaymentDetails = Nothing
    Set xmlCustomerList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlMortgageProduct = Nothing
    Set xmlComponentPayment = Nothing
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Set xmlDisb = Nothing
    Set xmlTempRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set objMortSQ = Nothing
    Set xmlResponse = Nothing
    ' PSC 08/10/2002 BMIDS00595 - End

    
    errCheckError strFunctionName, TypeName(Me)
End Function



Private Function CHAPSCreatePaymentRecord(ByVal sUnitID As String, ByRef vdTotalAmount As Variant, ByVal xmlPaymentDetails As IXMLDOMElement, ByVal xmlCustomerList As IXMLDOMNodeList, ByVal xmlMortgageLender As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:
'   Creates and returns a single CHAPS payment/transaction record from the information
'   supplied.
'
' pass:
'   sUnitID
'   vdTotalAmount - Current running total.
'   xmlPaymentDetails
'   xmlCustomerList
'   xmlMortgageLender
' return:
'   String - CHAPS transaction record.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------
    
    'As the record is a fixed width, it'll be quicker to insert text into the
    'the string than to append (and re-allocate memory).
    Dim sRecord As String * CHAPS_PAYMENT_LENGTH
    Dim xmlCurrent As IXMLDOMElement
    Dim intCurrentPos As Integer    'BM0155 MDC 04/12/2002
    
    Const strFunctionName As String = "CHAPSCreatePaymentRecord"
    
    On Error GoTo CHAPSCreatePaymentRecordExit
        
    'Initialise the record with the field seperator codes (we're effectively going to 'fill in the blanks').
    sRecord = String(CHAPS_PAYMENT_LENGTH, CHAPS_SEP)
        
    'POS FIELD NAME                         FORMAT+LENGTH       OMIGA 4 - MAPPING
    '================================================================================================
    '001 Field Type                          X                   'P'
    '003 Account Sort Code                   9(6)                {parameter value}
    '010 Account Number                      9(8)                {parameter value}
    '017 Beneficiary Name                    X(35)               Payee Name of the payment
    '053 Beneficiary Address 1               X(35)               Payee address: House Name or No + space + Street
    '089 Beneficiary Address 2               X(35)               Payee address: District + space + Town
    '125 Beneficiary Address 3               X(35)               Payee address: County + Post Code
    '161 Beneficiary SortCode                9(6)                Payee Bank sort code (no dashes)
    
    Dim sAddress1 As String
    Dim sAddress2 As String
    Dim sAddress3 As String
    Dim sPayeeName As String
    Dim sPayeeSortCode As String
    Dim sBarclaysCHAPSSortCode As String
    Dim sBarclaysCHAPSAccountNumber As String

    'Beneficiary Name.
    sPayeeName = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYEEHISTORY/THIRDPARTY"), "COMPANYNAME")
    
    'Beneficiary Address 1.
    Set xmlCurrent = xmlPaymentDetails.selectSingleNode("//PAYEEHISTORY/THIRDPARTY/ADDRESS")
    
    If Not xmlCurrent.getAttributeNode("BUILDINGORHOUSENUMBER") Is Nothing Then
        sAddress1 = xmlGetAttributeText(xmlCurrent, "BUILDINGORHOUSENUMBER")
    Else
        sAddress1 = xmlGetAttributeText(xmlCurrent, "BUILDINGORHOUSENAME")
    End If
    
    sAddress1 = sAddress1 & " " & xmlGetAttributeText(xmlCurrent, "STREET")
    
    'Beneficiary Address 2.
    sAddress2 = Trim$(xmlGetAttributeText(xmlCurrent, "DISTRICT") & " " & xmlGetAttributeText(xmlCurrent, "TOWN"))
    
    'Beneficiary Address 3.
    sAddress3 = Trim$(xmlGetAttributeText(xmlCurrent, "COUNTY") & " " & xmlGetAttributeText(xmlCurrent, "POSTCODE"))
    
    'Payee Sort Code.
    sPayeeSortCode = Replace(xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYEEHISTORY"), "BANKSORTCODE"), "-", "")
        
    'Account Sort Code.
    sBarclaysCHAPSSortCode = Replace(xmlMortgageLender.selectSingleNode("//BARCLAYSCHAPSSORTCODE").Text, "-", "")
    
    'Account Number.
    sBarclaysCHAPSAccountNumber = xmlMortgageLender.selectSingleNode("//BARCLAYSCHAPSACCOUNTNUMBER").Text
        
    'BM0155 MDC 04/12/2002
'    Mid$(sRecord, 1, 1) = CHAPS_PAYMENT_PREFIX
'    Mid$(sRecord, 3, 6) = FormatField(sBarclaysCHAPSSortCode, "000000")
'    Mid$(sRecord, 10, 6) = FormatField(sBarclaysCHAPSAccountNumber, "000000")
'    Mid$(sRecord, 17, 35) = FormatField(sPayeeName, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 53, 35) = FormatField(sAddress1, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 89, 35) = FormatField(sAddress2, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 125, 35) = FormatField(sAddress3, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 161, 6) = FormatField(sPayeeSortCode, "000000")
    intCurrentPos = 1
    Mid$(sRecord, intCurrentPos, 1) = CHAPS_PAYMENT_PREFIX
    intCurrentPos = intCurrentPos + 2
    Mid$(sRecord, intCurrentPos, 6) = FormatField(sBarclaysCHAPSSortCode, "000000")
    intCurrentPos = intCurrentPos + 7
    Mid$(sRecord, intCurrentPos, 8) = FormatField(sBarclaysCHAPSAccountNumber, "00000000")
    intCurrentPos = intCurrentPos + 9
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sPayeeName, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sAddress1, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sAddress2, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sAddress3, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 6) = FormatField(sPayeeSortCode, "000000")
    intCurrentPos = intCurrentPos + 7
    'BM0155 MDC 04/12/2002 - End
    
    'POS FIELD NAME                     FORMAT+LENGTH       OMIGA 4 - MAPPING
    '================================================================================================
    '168 Beneficiary Account No.         X(24)              Payee Bank Account Number
    '193 Beneficiary Type                X(8)               'Customer'
    '202 Beneficiary Advice Ind          X                  'N'
    '204 Beneficiary Bank Name           X(35)              Payee Bank Name
    '240 Beneficiary Bank Address 1      X(35)              Spaces
    '276 Beneficiary Bank Address 2      X(35)              Spaces
    '312 Beneficiary Bank Address 3      X(35)              Spaces
    '348 Beneficiary Bank Info 1         X(35)              Spaces
    
    Dim sPayeeBankName As String
    Dim sPayeeAccountNo As String
    
    'Beneficiary Account No.
    sPayeeAccountNo = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYEEHISTORY"), "ACCOUNTNUMBER")
    
    'Beneficiary Bank Name
    sPayeeBankName = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYEEHISTORY"), "BANKNAME")
    
    'BM0155 MDC 04/12/2002
'    Mid$(sRecord, 168, 24) = FormatField(sPayeeAccountNo, "!" & String(24, "@"))
'    Mid$(sRecord, 193, 8) = "Customer"
'    Mid$(sRecord, 202, 1) = "N"
'    Mid$(sRecord, 204, 35) = FormatField(sPayeeBankName, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 240, 35) = String(35, " ")
'    Mid$(sRecord, 276, 35) = String(35, " ")
'    Mid$(sRecord, 312, 35) = String(35, " ")
'    Mid$(sRecord, 348, 35) = String(35, " ")
    Mid$(sRecord, intCurrentPos, 24) = FormatField(sPayeeAccountNo, "!" & String(24, "@"))
    intCurrentPos = intCurrentPos + 25
    Mid$(sRecord, intCurrentPos, 8) = "Customer"
    intCurrentPos = intCurrentPos + 9
    Mid$(sRecord, intCurrentPos, 1) = "N"
    intCurrentPos = intCurrentPos + 2
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sPayeeBankName, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    'BM0155 MDC 04/12/2002 - End
    
    
    'POS FIELD NAME                     FORMAT+LENGTH       OMIGA 4 - MAPPING
    '================================================================================================
    '384 Beneficiary Bank Info 2         X(35)              Spaces
    '420 Beneficiary Bank Info 3         X(35)              Spaces
    '456 Beneficiary Bank Info 4         X(35)              Spaces
    '492 Beneficiary Bank Info 5         X(35)              Spaces
    '528 Ordering Customer 1             X(35)              Unit Code
    '564 Ordering Customer 2             X(35)              Spaces
    '600 Ordering Customer 3             X(35)              Spaces
    '636 Ordering Customer 4             X(35)              Spaces

    'BM0155 MDC 04/12/2002
'    Mid$(sRecord, 384, 35) = String(35, " ")
'    Mid$(sRecord, 420, 35) = String(35, " ")
'    Mid$(sRecord, 456, 35) = String(35, " ")
'    Mid$(sRecord, 492, 35) = String(35, " ")
'    Mid$(sRecord, 528, 35) = FormatField(sUnitID & " ", CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 564, 35) = String(35, " ")
'    Mid$(sRecord, 600, 35) = String(35, " ")
'    Mid$(sRecord, 636, 35) = String(35, " ")
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sUnitID & " ", CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    'BM0155 MDC 04/12/2002 - End

    'POS FIELD NAME                     FORMAT+LENGTH       OMIGA 4 - MAPPING
    '================================================================================================
    '672 Ordering Institution 1          X(35)              {parameter value}
    '708 Ordering Institution 2          X(35)              Spaces
    '744 Ordering Institution 3          X(35)              Spaces
    '780 Ordering Institution 4          X(35)              Spaces
    '815 Payment Reference               undefined          '=='
    '818 Payment Details 1               X(35)              primary customer Title, forename, surname
    '854 Payment Details 2               X(35)              second customer Title, forename, surname
    '890 Payment Details 3               X(35)              "Please Advise Customer"

    Dim sCustomerTitle1 As String
    Dim sCustomerTitle2 As String
    Dim sBarclaysCHAPSOrderingInst As String
    
    ' PSC 18/11/2002 BMIDS00707 - Start
    Dim strValue As String
    
    'Payment Details 1.
    strValue = xmlGetNodeText(xmlCustomerList.Item(0).selectSingleNode("CUSTOMERVERSION"), "TITLE")
    
    If Not IsValidationType("Title", strValue, "O") Then
        sCustomerTitle1 = GetComboText("Title", strValue)
    Else
        sCustomerTitle1 = xmlGetNodeText(xmlCustomerList.Item(0).selectSingleNode("CUSTOMERVERSION"), "TITLEOTHER")
    End If
    
    If Len(sCustomerTitle1) > 0 Then
        sCustomerTitle1 = sCustomerTitle1 & " "
    End If
    
    sCustomerTitle1 = sCustomerTitle1 & xmlGetNodeText(xmlCustomerList.Item(0).selectSingleNode("CUSTOMERVERSION"), "FIRSTFORENAME")
    
    If Len(sCustomerTitle1) > 0 Then
        sCustomerTitle1 = sCustomerTitle1 & " "
    End If
    
    sCustomerTitle1 = sCustomerTitle1 & xmlGetNodeText(xmlCustomerList.Item(0).selectSingleNode("CUSTOMERVERSION"), "SURNAME")
    
    ' sCustomerTitle1 = xmlGetNodeText(xmlCustomerList.Item(0).selectSingleNode("CUSTOMERVERSION"), "CORRESPONDENCESALUTATION")
    
    'Payment Details 2.
    If xmlCustomerList.length > 1 Then
        
        strValue = xmlGetNodeText(xmlCustomerList.Item(1).selectSingleNode("CUSTOMERVERSION"), "TITLE")
        
        If Not IsValidationType("Title", strValue, "O") Then
            sCustomerTitle2 = GetComboText("Title", strValue)
        Else
            sCustomerTitle2 = xmlGetNodeText(xmlCustomerList.Item(1).selectSingleNode("CUSTOMERVERSION"), "TITLEOTHER")
        End If
        
        If Len(sCustomerTitle2) > 0 Then
            sCustomerTitle2 = sCustomerTitle2 & " "
        End If
        
        sCustomerTitle2 = sCustomerTitle2 & xmlGetNodeText(xmlCustomerList.Item(1).selectSingleNode("CUSTOMERVERSION"), "FIRSTFORENAME")
        
        If Len(sCustomerTitle2) > 0 Then
            sCustomerTitle2 = sCustomerTitle2 & " "
        End If
        
        sCustomerTitle2 = sCustomerTitle2 & xmlGetNodeText(xmlCustomerList.Item(1).selectSingleNode("CUSTOMERVERSION"), "SURNAME")
    
'        sCustomerTitle2 = xmlGetNodeText(xmlCustomerList.Item(1).selectSingleNode("CUSTOMERVERSION"), "CORRESPONDENCESALUTATION")
    End If
    ' PSC 18/11/2002 BMIDS00707 - End
    
    'Ordering Institution 1
    sBarclaysCHAPSOrderingInst = xmlMortgageLender.selectSingleNode("//BARCLAYSCHAPSORDERINGINST").Text
    
    'BM0155 MDC 04/12/2002
'    Mid$(sRecord, 672, 35) = FormatField(sBarclaysCHAPSOrderingInst, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 708, 35) = String(35, " ")
'    Mid$(sRecord, 744, 35) = String(35, " ")
'    Mid$(sRecord, 780, 35) = String(35, " ")
'    Mid$(sRecord, 815, 2) = "=="
'    Mid$(sRecord, 818, 35) = FormatField(sCustomerTitle1, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 854, 35) = FormatField(sCustomerTitle2, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 890, 35) = FormatField("Please Advise Customer", CHAPS_TEXT_FORMAT)
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sBarclaysCHAPSOrderingInst, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = String(35, " ")
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 1) = "="
    intCurrentPos = intCurrentPos + 1
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sCustomerTitle1, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sCustomerTitle2, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    Mid$(sRecord, intCurrentPos, 35) = FormatField("Please Advise Customer", CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    'BM0155 MDC 04/12/2002 - End

    'POS FIELD NAME                     FORMAT+LENGTH       OMIGA 4 - MAPPING
    '================================================================================================
    '926 Payment Details 4               X(35)              Application Number
    '962 Payment Amount                  9(9).99            Net Payment amount (i.e. actual amount of payment}
    '971 Terminator                      X                  '#'
    '972 EOL                             X(2)               vbCrLf

    Dim sPaymentAmount As String
    Dim sApplicationNumber As String
    
    'Payment Details 4.
    sApplicationNumber = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD"), "APPLICATIONNUMBER")

    'Payment Amount.
    sPaymentAmount = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/DISBURSEMENTPAYMENT"), "NETPAYMENTAMOUNT")
    
    'BM0155 MDC 04/12/2002
'    Mid$(sRecord, 926, 35) = FormatField(sApplicationNumber, CHAPS_TEXT_FORMAT)
'    Mid$(sRecord, 962, 9) = FormatField(sPaymentAmount, "000000.00")
'    Mid$(sRecord, 971, 1) = "#"
'    Mid$(sRecord, 972, 2) = vbCrLf
    Mid$(sRecord, intCurrentPos, 35) = FormatField(sApplicationNumber, CHAPS_TEXT_FORMAT)
    intCurrentPos = intCurrentPos + 36
    'BM0173 MDC 10/12/2002
    'Mid$(sRecord, intCurrentPos, 9) = FormatField(sPaymentAmount, "000000.00")
    'intCurrentPos = intCurrentPos + 9
    Mid$(sRecord, intCurrentPos, 12) = FormatField(sPaymentAmount, "000000000.00")
    intCurrentPos = intCurrentPos + 12
    'BM0173 MDC 10/12/2002 - End
    Mid$(sRecord, intCurrentPos, 1) = "#"
    intCurrentPos = intCurrentPos + 1
    Mid$(sRecord, intCurrentPos, 2) = vbCrLf
    'BM0155 MDC 04/12/2002 - End
    
    'BM0220 MDC 02/01/2003
    sRecord = Replace$(sRecord, "&", " ")
    'BM0220 MDC 02/01/2003
    
    'Increment the running Total Amount (prefix with zero incase its empty).
    vdTotalAmount = vdTotalAmount + CDec("0" & sPaymentAmount)

    'Return the payment record to the caller.
    CHAPSCreatePaymentRecord = sRecord
    
CHAPSCreatePaymentRecordExit:
    Set xmlCurrent = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function


Private Function FormatField(ByVal sValue As String, ByVal sFormat As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Formats the field into the specified format. If the value is an empty string then
'   it is altered into either a zero or a space (for numeric and string field formats)
'   before being formatted.
'
' pass:
'   sValue - Original field value.
'   sFormat - Format to apply to value.
' return:
'   String - CHAPS formatted field value.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------
    
    'Ascertain if we're looking at a numeric field format or a string. And ensure we
    'have a value.
    If Left$(sFormat, 1) = "0" Then
        If sValue = "" Then
            sValue = "0"
        End If
    Else
        If sValue = "" Then
            sValue = " "
        End If
    End If
    
    FormatField = Format$(sValue, sFormat)
    
End Function

Private Function CHAPSCreateTrailer(ByVal iRecordCount As Integer, ByVal vdTotalAmount As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
'   Creates and returns a trailer record from the information specified.
'
' pass:
'   iRecordCount  - The number of transaction records in the body.
'   vdTotalAmount - The total payment amount in the file.
' return:
'   String - CHAPS trailer record.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------

    Dim sRecord As String
    
    Const strFunctionName As String = "CHAPSCreateTrailer"

    On Error GoTo CHAPSCreateTrailerExit
    
    'Now create a trailer record and return it.
    sRecord = CHAPS_TRAILER_PREFIX & CHAPS_SEP
    sRecord = sRecord & FormatField(iRecordCount, "000") & CHAPS_SEP
    'BM0173 MDC 10/12/2002
    'sRecord = sRecord & FormatField(vdTotalAmount, "000000000.00") & CHAPS_RECORD_SUFFIX & vbCrLf
    sRecord = sRecord & FormatField(vdTotalAmount, "000000000000.00") & CHAPS_RECORD_SUFFIX & vbCrLf
    'BM0173 MDC 10/12/2002 - End
    
    'Return the trailer record.
    CHAPSCreateTrailer = sRecord
    
CHAPSCreateTrailerExit:
    errCheckError strFunctionName, TypeName(Me)
End Function


Private Sub ProcessHexagonPayment(ByVal xmlPayment As IXMLDOMElement, ByRef vxmlResponseNode As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Writes or appends payments to a HEXAGON payment file. The path is obtained from a
'   global parameter.
'
' pass:
'   xmlPayment
'       <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER=""/>
'   vxmlResponseElement
'       XML standard RESPONSE node.
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    
    Dim iPos As Integer
    Dim sBody As String
    Dim sHeader As String
    Dim sFilename As String
    Dim sFolderPath As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objFile As TextStream
    Dim iRecordCount As Integer
    Dim vdTotalAmount As Variant
    Dim objFileSystem As FileSystemObject
    Dim objGlobalParam As GlobalParameterDO
    
    Dim sFailureReason As String                        ' PSC 19/09/2002 SYS5551
    Dim xmlErrorNode, xmlDescNode As IXMLDOMElement     ' PSC 19/09/2002 SYS5551
    
    
    On Error GoTo ProcessHexagonPaymentExit
    
    Const strFunctionName As String = "ProcessHexagonPayment"
    
    'Create a response.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    Set vxmlResponseNode = xmlOut.createElement("RESPONSE")
    
    'Get the path and filename from the 'HexagonPathAndFilename' global parameter.
    Set objGlobalParam = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    sFilename = objGlobalParam.GetCurrentParameterByType("HexagonPathAndFilename", "STRING")
    
    'Create path if it doesn't exist.
    Set objFileSystem = New FileSystemObject
    iPos = InStrRev(sFilename, "\", , vbTextCompare)
    
    If iPos > 0 Then
        sFolderPath = Left$(sFilename, iPos)
        
        If objFileSystem.FolderExists(sFolderPath) = False Then
            objFileSystem.CreateFolder sFolderPath
        End If
    End If
    
    'Get header and body from any existing file, otherwise create one.
    sBody = HexagonOpenOrCreateBody(iRecordCount, vdTotalAmount, sFilename)
    
    'Append any additionally requested payment records.
    'SG 30/05/02 SYS4767 START
    'SG 20/05/02 MSMS0083 - HexagoProcessNewPayment now returns boolean so we can validate the body
    'OLD CODE:
    'sBody = sBody & HexagonProcessNewPayment(iRecordCount, vdTotalAmount, xmlPayment)
    'NEW CODE:
    'PSC 19/09/2002 SYS5551 Add extra parameter to say why this has failed
    If HexagonProcessNewPayment(iRecordCount, vdTotalAmount, xmlPayment, sFailureReason) = False Then
'    If HexagonProcessNewPayment(iRecordCount, vdTotalAmount, xmlPayment) = False Then

        'invalid
        vxmlResponseNode.setAttribute "TYPE", "ERROR"
        
        ' PSC 19/09/2002 SYS5551 - Start
        ' Now add reason for failure to be written to BatchAudit table {
        Set xmlErrorNode = xmlOut.createElement("ERROR")
        Set xmlDescNode = xmlOut.createElement("DESCRIPTION")
        xmlDescNode.Text = sFailureReason
        xmlErrorNode.appendChild xmlDescNode
        vxmlResponseNode.appendChild xmlErrorNode
        ' PSC 19/09/2002 SYS5551 - End
 
        GoTo ProcessHexagonPaymentExit
    End If
    sBody = sBody & mstrHexagonPaymentBody
    'End SG 20/05/02 MSMS0083
    'SG 30/05/02 SYS4767 END
    
    'Create a header record from the new totals.
    sHeader = HexagonCreateHeader(iRecordCount, vdTotalAmount)
    
    'Attempt to open the existing file (otherwise create one).
    Set objFile = objFileSystem.OpenTextFile(sFilename, ForWriting, True, TristateUseDefault)

    'Output the CHAPS records.
    objFile.Write sHeader
    objFile.Write sBody
    
    'Close the file.
    objFile.Close
    
    'Set the response to success.
    vxmlResponseNode.setAttribute "TYPE", "SUCCESS"
    
ProcessHexagonPaymentExit:
    Set xmlOut = Nothing
    Set objFile = Nothing
    Set objFileSystem = Nothing
    Set objGlobalParam = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub


Private Function HexagonOpenOrCreateBody(ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, sFilename As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Reads the header from the file specified and extracts the count and amount from it. If
'   the file is empty these values are initialised. Any existing transaction records are
'   also returned.
'
' pass:
'   iRecordCount  - Counter for the number of transaction records.
'   vdTotalAmount - Sum of all payment amounts.
'   sFilename     - Path and filename to use.
' return:
'   String - Any existing payment records.
' exceptions:
'   Invalid file.
'------------------------------------------------------------------------------------------

    Dim sBuffer As String
    Dim sHeader As String
    Dim clsFile As TextStream
    Dim clsFileSystem As FileSystemObject

    Const strFunctionName As String = "HexagonOpenOrCreateBody"

    On Error GoTo HexagonOpenOrCreateBodyExit
    
    'Attempt to open the existing file (otherwise create one).
    Set clsFileSystem = New FileSystemObject
    Set clsFile = clsFileSystem.OpenTextFile(sFilename, ForReading, True, TristateUseDefault)
    
    If clsFile.AtEndOfStream = False Then
        'Read the contents of the file into a buffer.
        sBuffer = clsFile.ReadAll()
    End If
    
    'Close the file.
    clsFile.Close
    
    'If the buffer is not empty, verify and parse the file.
    If Len(sBuffer) > 0 Then
        sHeader = Left$(sBuffer, HEXAGON_HEADER_LENGTH)
        
        'If the header begins with a control tag '$FN' and ends in a vbCrLf then
        'we'll assume its valid.
        If Not (Left$(sHeader, 3) = "$FN" And Right$(sHeader, 2) = vbCrLf) Then
            'Throw an invalid file format error.
            errThrowError strFunctionName, 6004
        End If
        
        'Get the current record count and total amount from the header.
        iRecordCount = CInt(Mid$(sHeader, 4, 5))
        vdTotalAmount = CDec(Mid$(sHeader, 12, 18))
    
        'Remove the header record from the buffer.
        sBuffer = Right$(sBuffer, (Len(sBuffer) - HEXAGON_HEADER_LENGTH))
    End If
        
    'Return the payment records to the caller.
    HexagonOpenOrCreateBody = sBuffer
    
HexagonOpenOrCreateBodyExit:
    Set clsFile = Nothing
    Set clsFileSystem = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function


'PSC 19/09/2002 SYS5551 Add extra ReasonForFailure parameter
'Private Function HexagonProcessNewPayment(ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, ByVal xmlPayment As IXMLDOMNode) As String
Private Function HexagonProcessNewPayment(ByRef iRecordCount As Integer, ByRef vdTotalAmount As Variant, ByVal xmlPayment As IXMLDOMNode, ByRef rstrReasonForFailure) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Processes the payment request XML and creates a Hexagon record for each. These records
'   are then appended to the return string.
'
' pass:
'   iRecordCount  - Current number of payment records in the body so far.
'   vdTotalAmount - Total of all existing payment records.
'   xmlPayment
'       <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER=""/>
' return:
'   String - Hexagon transaction records.
'------------------------------------------------------------------------------------------

    Dim sBuffer As String
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlLoanComponent As IXMLDOMElement
    Dim xmlPaymentDetails As IXMLDOMElement
    Dim xmlMortgageLender As IXMLDOMElement
    Dim xmlMortgageProduct As IXMLDOMElement
    Dim xmlComponentPayment As IXMLDOMElement
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Dim xmlDisb As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objMortSQ As MortgageSubQuoteBO
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim strIRValueId As String
    Dim strPaymentType As String
    Dim strResponse As String
    Dim strAppFFNo As String
    Dim strAppNo As String
    ' PSC 08/10/2002 BMIDS00595 - End
    
    ' PSC 13/03/2007 EP2_1347 - Start
    Dim blnIsIncentive As Boolean
    Dim blnIsValuationRefund As Boolean
    ' PSC 13/03/2007 EP2_1347 - End

    Const strFunctionName As String = "HexagonProcessNewPayment"

    On Error GoTo HexagonProcessNewPaymentExit
    
    'SG 30/05/02 SYS4767
    'SG 20/05/02 MSMS0083
    HexagonProcessNewPayment = True
    
    'Get the full Payment Details.
    Set xmlPaymentDetails = GetPaymentDetails(xmlPayment)
        
    'Get the Customers for the application.
    Set xmlCustomerList = GetCustomerList(xmlPayment)
           
    'Get the values to obtain the loan component.
    Set xmlComponentPayment = xmlPaymentDetails.selectSingleNode("//LOANCOMPONENTPAYMENT")
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Set xmlDisb = xmlGetMandatoryNode(xmlPayment, "DISBURSEMENTPAYMENT")
    strPaymentType = xmlGetMandatoryAttributeText(xmlDisb, "PAYMENTTYPE")
    ' PSC 08/10/2002 BMIDS00595 - End
    
    ' PSC 13/03/2007 EP2_1347 - Start
    blnIsIncentive = IsValidationType("PaymentType", CSafeInt(strPaymentType), "C")
    blnIsValuationRefund = IsValidationType("PaymentType", CSafeInt(strPaymentType), "VALREFUND")
    ' PSC 13/03/2007 EP2_1347 - End

    ' PSC 13/03/2007 EP2_1347
    If xmlComponentPayment Is Nothing And blnIsIncentive = False And blnIsValuationRefund = False Then
        sBuffer = ""
    Else
        ' PSC 08/10/2002 BMIDS00595 - Start
        ' PSC 13/03/2007 EP2_1347
        If blnIsIncentive Or blnIsValuationRefund Then
            
            ' Get the accepted list of components
            strAppNo = xmlGetMandatoryAttributeText(xmlDisb, "APPLICATIONNUMBER")
            strAppFFNo = GetFactFindNumberForApplication(strAppNo)
            
            Set xmlTempRequest = xmlPayment.ownerDocument.createElement("REQUEST")
            Set xmlRoot = xmlPayment.ownerDocument.createElement("FINDACTIVELOANCOMPONENTS")
            xmlTempRequest.appendChild xmlRoot
            Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONNUMBER")
            xmlRoot.appendChild xmlNode
            xmlNode.Text = strAppNo
            Set xmlNode = xmlPayment.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")
            xmlRoot.appendChild xmlNode
            xmlNode.Text = strAppFFNo

            Set objMortSQ = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
            
            strResponse = objMortSQ.FindActiveLoanComponentsList(xmlTempRequest.xml)
            Set xmlResponse = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlResponse.documentElement, , True
            
            ' Set the loan component to the first one as lender details will be the same
            ' fo all of them
            Set xmlLoanComponent = xmlGetMandatoryNode(xmlResponse, ".//LOANCOMPONENT")
        Else
            'Get the Loan Component details (to get to the lender details).
            Set xmlLoanComponent = GetLoanComponent(xmlComponentPayment)
        End If
        ' PSC 08/10/2002 BMIDS00595 - End
        
        'Get the Mortgage Product details (to get to the lender details).
        Set xmlMortgageProduct = GetMortgageProduct(xmlLoanComponent)
           
        'Get the Lender account details.
        ' PSC 19/09/2002 SYS5551
         Set xmlMortgageLender = GetMortgageLender(xmlMortgageProduct)

        'Sg 30/05/02 SYS4767 START
        'SG 20/05/02 MSMS0083 START
        If ValidateAccountNo(xmlGetNodeText(xmlMortgageLender, "//HEXAGONCHAPSACCOUNTNUMBER")) = False Then
            rstrReasonForFailure = "Lender Account Details - Invalid or missing account number."     ' PSC 19/09/2002 SYS5551
            HexagonProcessNewPayment = False
            GoTo HexagonProcessNewPaymentExit
        End If
        'SG 20/05/02 END
        'SG 30/05/02 SYS4767 END
        'SA 31/05/02 MSMS0085 We can only output 200 records - change 999 to 200
        'If the record conut equals or exceeds 999 then we must not add the record.
        If iRecordCount >= 200 Then
            errThrowError strFunctionName, 6005, "CHAPS"
        Else
            'Increment the record count.
            iRecordCount = iRecordCount + 1
        
            'Create a payment/transaction record.
            sBuffer = sBuffer & HexagonCreatePaymentRecord(vdTotalAmount, xmlPaymentDetails, xmlCustomerList, xmlMortgageLender)
        End If
    End If
    
    'Sg 30/05/02 SYS4767
    'SG 20/05/02 MSMS0083
    'HexagonProcessNewPayment = sBuffer
    mstrHexagonPaymentBody = sBuffer


HexagonProcessNewPaymentExit:
    Set xmlPaymentDetails = Nothing
    Set xmlCustomerList = Nothing
    Set xmlLoanComponent = Nothing
    Set xmlMortgageProduct = Nothing
    Set xmlComponentPayment = Nothing
    
    ' PSC 08/10/2002 BMIDS00595 - Start
    Set xmlDisb = Nothing
    Set xmlTempRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set objMortSQ = Nothing
    Set xmlResponse = Nothing
    ' PSC 08/10/2002 BMIDS00595 - End
    
    
    errCheckError strFunctionName, TypeName(Me)
End Function


Private Function HexagonCreatePaymentRecord(ByRef vdTotalAmount As Variant, ByVal xmlPaymentDetails As IXMLDOMElement, ByVal xmlCustomerList As IXMLDOMNodeList, ByVal xmlMortgageLender As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:
'   Creates and returns a single Hexagon payment/transaction record from the information
'   supplied.
'
' pass:
'   xmlPaymentDetails
'   xmlCustomerList
'   xmlMortgageLender
' return:
'   String - Hexagon transaction record.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------
    
    'As the record is a fixed width, it'll be quicker to insert text into the
    'the string than to append (and re-allocate memory).
    Dim sRecord As String * HEXAGON_PAYMENT_LENGTH
    Dim xmlCurrent As IXMLDOMElement
    
    Const strFunctionName As String = "HexagonCreatePaymentRecord"
    
    On Error GoTo HexagonCreatePaymentRecordExit
        
    'Blank the record.
    sRecord = String(HEXAGON_PAYMENT_LENGTH, " ")
    
    'TAG POS   TYPE        FORMAT  DESCRIPTION
    '====================================================================================
    '$DS 001   Mandatory   XXX     Control Tag indicating start of set of payment instructions
    'F01 004   Mandatory   9(8)    Account number to be debited
    'F02 012   Mandatory   9(6)    Payment Date
    'F03 018   Optional    X(10)   Textual description of payment
    'F04 028   Mandatory   999     Total number of contra-balancing '$LS' entries
    
    'F05 040   Mandatory   9(12)   Total monetary value of contra-balancing '$LS' entries
    '$LS 043   Mandatory   XXX     Control Tag indicating start of individual payment
    'S01 051   Mandatory   9(8)    Account number to be credited
    'S02 057   Mandatory   9(6)    Bank Sort Code
    'S03 063   Mandatory   X(18)   Name of beneficiary
    
    'S04 081   Mandatory   9(11)   Payment amount
    'S05 092   Optional    X(18)   Textual description of individual payment
    '$LE 095   Mandatory   XXX     Control Tag indicating end of individual payment
    '$DE 098   Mandatory   XXX     Control Tag indicating end of set of payment instructions
    'EOL 101   Mandatory   XX      vbCrLf

    Dim sSortCode As String
    Dim sPayeeName As String
    Dim sPaymentDate As String
    Dim sDebitAccount As String
    Dim sPaymentAmount As String
    Dim sCreditAccount As String
        
    'Account number to be debited.
    sDebitAccount = xmlMortgageLender.selectSingleNode("//HEXAGONCHAPSACCOUNTNUMBER").Text
    
    'Payment Date.
    'sPaymentDate = Format$(xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/DISBURSEMENTPAYMENT"), "PROCESSINGDATE"), "ddmmyy")
    'MSMS0085
    sPaymentDate = CStr(Format$(Now(), "DD/MM/YY"))
    
    'Total monetary value of contra-balancing '$LS' entries.
    sPaymentAmount = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/DISBURSEMENTPAYMENT"), "NETPAYMENTAMOUNT")
    
    'Account number to be credited.
    sCreditAccount = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/PAYEEHISTORY"), "ACCOUNTNUMBER")
    
    'Bank Sort Code.
    sSortCode = Replace(xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/PAYEEHISTORY"), "BANKSORTCODE"), "-", "")
    
    'Name of beneficiary.
    sPayeeName = xmlGetAttributeText(xmlPaymentDetails.selectSingleNode("//PAYMENTRECORD/PAYEEHISTORY/THIRDPARTY"), "COMPANYNAME")
    
    Mid$(sRecord, 1, 3) = "$DS"
    Mid$(sRecord, 4, 8) = FormatField(sDebitAccount, "00000000")
    Mid$(sRecord, 12, 6) = FormatField(sPaymentDate, "000000")
    Mid$(sRecord, 18, 10) = String(10, " ")
    Mid$(sRecord, 28, 3) = "001"
    
    Mid$(sRecord, 31, 12) = FormatField(sPaymentAmount, "000000000000")
    Mid$(sRecord, 43, 3) = "$LS"
    Mid$(sRecord, 46, 8) = FormatField(sCreditAccount, "00000000")
    Mid$(sRecord, 54, 6) = FormatField(sSortCode, "000000")
    'MSMS0085 The ! is not forcing the format to process left to right - need to do this another way.
    'Mid$(sRecord, 60, 18) = FormatField(sPayeeName, "!" & String(18, "@"))
    Mid$(sRecord, 60, 18) = Left(sPayeeName & String(18, " "), 18)
    Mid$(sRecord, 78, 11) = FormatField(sPaymentAmount, "00000000000")
    Mid$(sRecord, 89, 18) = String(18, " ")
    Mid$(sRecord, 105, 3) = "$LE"
    Mid$(sRecord, 108, 3) = "$DE"
    Mid$(sRecord, 111, 2) = vbCrLf
    
    'Increment the running Total Amount (prefix with zero incase its empty).
    vdTotalAmount = vdTotalAmount + CDec("0" & sPaymentAmount)

    'Return the payment record to the caller.
    HexagonCreatePaymentRecord = sRecord
    
HexagonCreatePaymentRecordExit:
    Set xmlCurrent = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function

Private Function HexagonCreateHeader(ByVal iRecordCount As Integer, ByVal vdTotalAmount As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
'   Creates and returns a header record from the information specified.
'
' pass:
'   iRecordCount  - The number of transaction records in the body.
'   vdTotalAmount - The total payment amount in the file.
' return:
'   String - Hexagon header record.
' exceptions:
'   none.
'------------------------------------------------------------------------------------------

    Dim sRecord As String
    
    Const strFunctionName As String = "HexagonCreateHeader"

    On Error GoTo HexagonCreateHeaderExit
    
    'Now create a header record and return it.
    sRecord = "$FN" & FormatField(iRecordCount, "00000")
    sRecord = sRecord & "$FT" & FormatField(vdTotalAmount, String(18, "0"))
    sRecord = sRecord & "$FC" & "GBP" & vbCrLf
    
    'Return the header record.
    HexagonCreateHeader = sRecord
    
HexagonCreateHeaderExit:
    errCheckError strFunctionName, TypeName(Me)
End Function

Private Function GetPaymentDetails(ByVal xmlPayment As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Returns full payment details from a payment record element.
'
' pass:
'   xmlPayment - <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER="" />
' return:
'   String - Full payment details.
'------------------------------------------------------------------------------------------
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRequest As IXMLDOMElement
    Dim objPayProcBO As PaymentProcessingBO

    Const strFunctionName As String = "GetPaymentDetails"
    
    On Error GoTo GetPaymentDetailsExit

    'Create a PayProc BO.
    Set objPayProcBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingBO")
    
    'We'll need a working XML Dom Doc.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    'Create a payment details request.
    Set xmlRequest = xmlOut.createElement("REQUEST")
    xmlRequest.setAttribute "OPERATION", "GetFullPaymentDetails"
    xmlRequest.appendChild xmlPayment.cloneNode(True)
    
    'Alter the PaymentMethod from a ValueID to the ValidationType.
    xmlRequest.selectSingleNode("//@PAYMENTMETHOD").Text = "YC"
        
    'Obtain the payment details for the specified payment.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    xmlOut.loadXML objPayProcBO.omPayProcRequest(xmlRequest.xml)
    
    'Return the full details to the caller.
    Set GetPaymentDetails = xmlOut.documentElement
    
GetPaymentDetailsExit:
    Set objPayProcBO = Nothing
    Set xmlOut = Nothing
    Set xmlRequest = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function

Private Function GetCustomerList(ByVal xmlPayment As IXMLDOMElement) As IXMLDOMNodeList
' header ----------------------------------------------------------------------------------
' description:
'   Returns all the customers for an application from a payment record element.
'
' pass:
'   xmlPayment - <PAYMENTRECORD APPLICATIONNUMBER="" PAYMENTSEQUENCENUMBER="" />
' return:
'   String - Customer List.
'------------------------------------------------------------------------------------------
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCurrent As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim objApplManagerBO As ApplicationManagerBO
    
    Const strFunctionName As String = "GetCustomerList"

    On Error GoTo GetCustomerListExit

    'Create an app mgr BO.
    Set objApplManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'We'll need a working XML Dom Doc.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    'Create a customer enquiry request.
    Set xmlRequest = xmlOut.createElement("REQUEST")
    Set xmlCurrent = xmlOut.createElement("APPLICATIONNUMBER")
    xmlCurrent.Text = xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER")
    xmlRequest.appendChild xmlCurrent
                
    Set xmlCurrent = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlCurrent.Text = GetFactFindNumberForApplication(xmlGetMandatoryAttributeText(xmlPayment, "APPLICATIONNUMBER"))
    xmlRequest.appendChild xmlCurrent
                
    'Obtain all the customers for this application.
    xmlOut.loadXML objApplManagerBO.FindCustomersForApplication(xmlRequest.xml)
        
    'Get a customer list and return it to the caller.
    Set GetCustomerList = xmlOut.selectNodes("RESPONSE/CUSTOMER")
    
GetCustomerListExit:
    Set objApplManagerBO = Nothing
    Set xmlOut = Nothing
    Set xmlCurrent = Nothing
    Set xmlRequest = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function


Private Function GetMortgageProduct(ByVal xmlLoanComponent As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Returns the mortgage product details from the specified loan component.
'
' pass:
'   xmlLoanComponent
' return:
'   String - Mortgage Product
'------------------------------------------------------------------------------------------

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCurrent As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim objMortgageProductBO As Object
    Dim strResponse As String               ' PSC 20/09/2002 BMIDS00480

    Const strFunctionName As String = "GetMortgageProduct"
    
    On Error GoTo GetMortgageProductExit
        
    Set objMortgageProductBO = gobjContext.CreateInstance("omMP.MortgageProductBO")
    
    'Create a request dom doc.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
       
    'Create a request element.
    Set xmlRequest = xmlOut.createElement("REQUEST")
    xmlOut.appendChild xmlRequest
    Set xmlCurrent = xmlOut.createElement("MORTGAGEPRODUCT")
    xmlOut.selectSingleNode("REQUEST").appendChild xmlCurrent
    
    'Append required fields
    Set xmlCurrent = xmlOut.createElement("MORTGAGEPRODUCTCODE")
    xmlCurrent.Text = xmlLoanComponent.selectSingleNode("//MORTGAGEPRODUCTCODE").Text
    xmlOut.selectSingleNode("//MORTGAGEPRODUCT").appendChild xmlCurrent
            
    Set xmlCurrent = xmlOut.createElement("STARTDATE")
    xmlCurrent.Text = xmlLoanComponent.selectSingleNode("//STARTDATE").Text
    xmlOut.selectSingleNode("//MORTGAGEPRODUCT").appendChild xmlCurrent
                    
    'Call Mortgage Product component.
    ' PSC 20/09/2002 BMIDS00480 - Start
    strResponse = objMortgageProductBO.FindProductList(xmlRequest.xml)
    Set xmlOut = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlOut.documentElement, , True
    ' PSC 20/09/2002 BMIDS00480 - End
    
    'Return the product to the caller.
    Set GetMortgageProduct = xmlOut.selectSingleNode(".//MORTGAGEPRODUCT")
        
GetMortgageProductExit:
    Set xmlOut = Nothing
    Set xmlRequest = Nothing
    Set xmlCurrent = Nothing
    Set objMortgageProductBO = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Function


Private Function GetMortgageLender(ByVal xmlMortgageProduct As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Returns the mortgage lender details from the specified product.
'
' pass:
'   xmlMortgageProduct
' return:
'   XML - Mortgage Lender
'------------------------------------------------------------------------------------------

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCurrent As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim objLenderBO As MortgageLenderBO
    
    Const strFunctionName As String = "GetMortgageLender"
    
    On Error GoTo GetMortgageLenderExit
    
    Set objLenderBO = gobjContext.CreateInstance(gstrORGANISATION_COMPONENT & ".MortgageLenderBO")
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
    
    'Add the Request element.
    Set xmlRequest = xmlOut.createElement("REQUEST")
    xmlOut.appendChild xmlRequest
    
    'Add the Lender element.
    Set xmlCurrent = xmlOut.createElement("MORTGAGELENDER")
    xmlOut.selectSingleNode("//REQUEST").appendChild xmlCurrent
    
    'Add the OrganisationID element.
    Set xmlCurrent = xmlOut.createElement("ORGANISATIONID")
    xmlCurrent.Text = xmlMortgageProduct.selectSingleNode("//ORGANISATIONID").Text
    xmlOut.selectSingleNode("//MORTGAGELENDER").appendChild xmlCurrent
    
    'Request the lender's details.
    xmlOut.loadXML objLenderBO.GetLenderDetails(xmlRequest.xml)
    errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True

    'Get the lender details.
    Set GetMortgageLender = xmlOut.selectSingleNode("//MORTGAGELENDER")
    
GetMortgageLenderExit:
    Set xmlOut = Nothing
    Set xmlRequest = Nothing
    Set xmlCurrent = Nothing
    Set objLenderBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function

Private Function GetLoanComponent(ByVal xmlComponentPayment As IXMLDOMElement) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description:
'   Return a loan component from the loancomponentpayment element.
'
' pass:
'   xmlComponentPayment
' return:
'   xmlLoanComponent
'------------------------------------------------------------------------------------------

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCurrent As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim objSubQuoteBO As MortgageSubQuoteBO

    Const strFunctionName As String = "GetLoanComponent"
    
    On Error GoTo GetLoanComponentExit
        
    Set objSubQuoteBO = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    
    'Create a request doc.
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    xmlOut.async = False
           
    'Set up the Request node.
    Set xmlRequest = xmlOut.createElement("REQUEST")
    xmlOut.appendChild xmlRequest
    Set xmlCurrent = xmlOut.createElement("MORTGAGESUBQUOTE")
    xmlOut.selectSingleNode("REQUEST").appendChild xmlCurrent
        
    'Append required fields.
    Set xmlCurrent = xmlOut.createElement("APPLICATIONNUMBER")
    xmlCurrent.Text = xmlGetAttributeText(xmlComponentPayment, "APPLICATIONNUMBER")
    xmlOut.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlCurrent
            
    Set xmlCurrent = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlCurrent.Text = xmlGetAttributeText(xmlComponentPayment, "APPLICATIONFACTFINDNUMBER")
    xmlOut.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlCurrent
          
    Set xmlCurrent = xmlOut.createElement("MORTGAGESUBQUOTENUMBER")
    xmlCurrent.Text = xmlGetAttributeText(xmlComponentPayment, "MORTGAGESUBQUOTENUMBER")
    xmlOut.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlCurrent
      
    Set xmlCurrent = xmlOut.createElement("LOANCOMPONENTSEQUENCENUMBER")
    xmlCurrent.Text = xmlGetAttributeText(xmlComponentPayment, "LOANCOMPONENTSEQUENCENUMBER")
    xmlOut.selectSingleNode("//MORTGAGESUBQUOTE").appendChild xmlCurrent
          
    'Call Cost Modelling component
    xmlOut.loadXML objSubQuoteBO.GetLoanComponentData(xmlRequest.xml)
    
    'Return the loan component to the caller.
    Set GetLoanComponent = xmlOut.selectSingleNode(".//LOANCOMPONENT")
        
GetLoanComponentExit:
    Set xmlOut = Nothing
    Set xmlCurrent = Nothing
    Set xmlRequest = Nothing
    Set objSubQuoteBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Function

Private Function ValidateSortCode(ByVal strSortCode As String) As Boolean
'SG 15/05/02 MSMS0083 Function added
'Purpose:   SortCode can either be "123456" or "12-34-56"
'           Check that sortcode is in correct format and contains no invalid characters

    On Error GoTo ValidateSortCodeExit
    
    Const strFunctionName As String = "ValidateSortCode"

    Dim lngPos1 As Long
    Dim lngPos2 As Long
    
    ValidateSortCode = True
    
    If Len(Trim(strSortCode)) = 0 Then
        'Invalid
        ValidateSortCode = False
        GoTo ValidateSortCodeExit
    End If
    
    If Len(strSortCode) <> 6 And Len(strSortCode) <> 8 Then
        'Invalid
        ValidateSortCode = False
        GoTo ValidateSortCodeExit
    End If
    
    If Len(strSortCode) = 6 Then
        'Has 6 characters, must be of format "123456"
        If IsNumeric(strSortCode) = False Then
            'One of the characters is non-numeric - Invalid
            ValidateSortCode = False
            GoTo ValidateSortCodeExit
        End If
    End If
    
    If Len(strSortCode) = 8 Then
        'Has 8 characters, must be of format "12-34-56"
        If Mid(strSortCode, 3, 1) <> "-" Then
            'First "-" is in the wrong place - Invalid
            ValidateSortCode = False
        Else
            If Mid(strSortCode, 6, 1) <> "-" Then
                'First "-" is in the wrong place - Invalid
                ValidateSortCode = False
            Else
                'Check constituent parts
                If IsNumeric(Mid(strSortCode, 1, 2)) = False Then
                    ValidateSortCode = False
                Else
                    If IsNumeric(Mid(strSortCode, 4, 2)) = False Then
                        ValidateSortCode = False
                    Else
                        If IsNumeric(Mid(strSortCode, 6, 2)) = False Then
                            ValidateSortCode = False
                        End If
                    End If
                End If
            End If
        End If
    End If

ValidateSortCodeExit:
    
    errCheckError strFunctionName, TypeName(Me)


End Function

Private Function ValidateAccountNo(ByVal strAccountNo As String) As Boolean
'SG 20/05/02 MSMS0083 Function added

    On Error GoTo ValidateAccountNoExit:
    
    Const strFunctionName As String = "ValidateAccountNo"

    ValidateAccountNo = True

    If Len(Trim(strAccountNo)) <> 8 Then
        'Account number not 8 characters
        ValidateAccountNo = False
        GoTo ValidateAccountNoExit:
    End If

    If Not IsNumeric(strAccountNo) Then
        'Account number contains invalid characters
        ValidateAccountNo = False
        GoTo ValidateAccountNoExit:
    End If

ValidateAccountNoExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Sub PrintCheque(ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AW     15/09/2002  BM029
'
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

On Error GoTo PrintChequeExit

    Const strFunctionName As String = "PrintCheque"
    
    Dim objPrintBO As Object
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlPrintAttributes As IXMLDOMNode
    Dim xmlControlData As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlPrintRequestNode As IXMLDOMNode
    Dim xmlPrintNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strTemplateId As String, strAppFFNo As String
     
    ' Get document attributes
    strTemplateId = GetMandatoryGlobalParamString("PProcPrintChequeTemplateId")
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    'Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    xmlDoc.appendChild vxmlRequestNode
    
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GetPrintAttributes"
    Set xmlPrintAttributes = vxmlRequestNode.ownerDocument.createElement("FINDATTRIBUTES")
    vxmlRequestNode.appendChild xmlPrintAttributes
    xmlSetAttributeValue xmlPrintAttributes, "HOSTTEMPLATEID", strTemplateId

    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    
    Dim strResponse  As String
    
    strResponse = objPrintBO.OmRequest(vxmlRequestNode.xml)
    'Check for errors
    Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponse xmlTempResponseDoc.xml, True

    Debug.Print xmlTempResponseDoc.xml
    
    Set xmlPrintAttributes = Nothing
    Set xmlPrintAttributes = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
    
    'GD BMIDS00569 START
    Dim strDestType As String
    strDestType = xmlGetAttributeText(xmlPrintAttributes, "PRINTERDESTINATIONTYPE")
    strDestType = GetValidationTypeForValueID("PrinterDestination", CSafeInt(strDestType))
    xmlSetAttributeValue xmlPrintAttributes, "PRINTERDESTINATIONTYPE", strDestType
    'GD BMIDS00569 END
    
    'Now print the document
    'Build up request string by copying Printer attributes
    Set xmlControlData = vxmlRequestNode.ownerDocument.createElement("CONTROLDATA")
    
    If Not xmlPrintAttributes Is Nothing Then
    
        xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "HOSTTEMPLATEID", "DOCUMENTID"
        xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DPSTEMPLATEID", "DPSDOCUMENTID"
        xmlCheckMandatoryAttribute xmlPrintAttributes, "DEFAULTCOPIES"
        xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "DEFAULTCOPIES", "COPIES"
        xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "REMOTEPRINTERLOCATION", "PRINTER"
        xmlCopyMandatoryAttributeValue xmlPrintAttributes, xmlControlData, "PRINTERDESTINATIONTYPE", "DESTINATIONTYPE"
        
        'Add PrintData Element and set attributes
        Set xmlPrintNode = vxmlRequestNode.ownerDocument.createElement("PRINTDATA")

        Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
        xmlCopyMandatoryAttributeValue xmlNode, xmlPrintNode, "APPLICATIONNUMBER", "APPLICATIONNUMBER"
        strAppFFNo = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
        
        If Len(Trim$(strAppFFNo)) = 0 Then
            'Get the ApplicationFactFind number
            strAppFFNo = GetFactFindNumberForApplication(xmlGetAttributeText(xmlNode, "APPLICATIONNUMBER"))
        End If
        xmlSetAttributeValue xmlPrintNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
        xmlCopyMandatoryAttributeValue xmlNode, xmlPrintNode, "PAYMENTSEQUENCENUMBER", "PAYMENTSEQUENCENUMBER"
        
        'Method Name (in PDM)
        xmlCheckMandatoryAttribute xmlPrintAttributes, "PDMMETHOD"
        xmlCopyAttributeValue xmlPrintAttributes, xmlPrintNode, "PDMMETHOD", "METHODNAME"
    End If
 
    Set xmlPrintRequestNode = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlPrintRequestNode, "OPERATION", "PrintDocument"
    xmlPrintRequestNode.appendChild xmlControlData
    xmlPrintRequestNode.appendChild xmlPrintNode
    
    ' Send data to print
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    Debug.Print xmlPrintRequestNode.xml
    strResponse = objPrintBO.OmRequest(xmlPrintRequestNode.xml)
    errCheckXMLResponse strResponse, True
    
PrintChequeExit:

    Set objPrintBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPrintAttributes = Nothing
    Set xmlControlData = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlPrintRequestNode = Nothing
    Set xmlPrintNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub CreateChequeAuditRecord( _
    ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateChequeAuditRecordExit

Dim objTxBO As IPaymentProcessingTxBO
Dim xmlNode As IXMLDOMNode
Dim strValue As String

Const strFunctionName As String = "CreateChequeAuditRecord"

    ' Raise error if no PAYMENTCHEQUEAUDIT node exists
    Set xmlNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTCHEQUEAUDIT")
    
    'Validate CACHEQUENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "CACHEQUENUMBER")
    
    'Validate CAUSERID attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "CAUSERID")
    
    'Validate APPLICATIONNUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "APPLICATIONNUMBER")
    
    'Validate PAYMENTSEQUENCENUMBER attribute is specified
    strValue = xmlGetMandatoryAttributeText(xmlNode, "PAYMENTSEQUENCENUMBER")
    
    If objTxBO Is Nothing Then
        Set objTxBO = gobjContext.CreateInstance(App.Title & ".PaymentProcessingTxBO")
    End If
    
    objTxBO.CreateChequeAuditRecord vxmlRequestNode, vxmlResponseNode
    
CreateChequeAuditRecordExit:
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IPaymentProcessingNTTxBO_CheckLocks(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AW     15/09/2002  BM029
'
On Error GoTo CheckLocksVbErr
    Const strFunctionName As String = "IPaymentProcessingNTTxBO_CheckLocks"

    Dim objAppManagerBO As ApplicationManagerBO
    Dim nodeNewRequest As IXMLDOMNode
    Dim elemTemp As IXMLDOMElement
    Dim nodeAppLock As IXMLDOMNode
    Dim xmlResponse As FreeThreadedDOMDocument40
    Dim strApplicationNumber As String
    Dim strResponse As String
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    strApplicationNumber = xmlGetAttributeText(vxmlRequestNode, "APPLICATIONNUMBER")
    'GD BMIDS0922
    'Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set nodeNewRequest = xmlDoc.createElement("REQUEST")
    'Set elemTemp = nodeNewRequest.ownerDocument.createElement("APPLICATIONLOCK")
    Set elemTemp = xmlDoc.createElement("APPLICATIONLOCK")
    Set nodeAppLock = nodeNewRequest.appendChild(elemTemp)
    'Set elemTemp = nodeNewRequest.ownerDocument.createElement("APPLICATIONNUMBER")
    Set elemTemp = xmlDoc.createElement("APPLICATIONNUMBER")
    elemTemp.Text = strApplicationNumber
    nodeAppLock.appendChild elemTemp
         
    ' call omApp
    Set objAppManagerBO = gobjContext.CreateInstance("omApp.ApplicationManagerBO")
    
    strResponse = objAppManagerBO.GetLockData(nodeNewRequest.xml)
    
    Set xmlResponse = xmlLoad(strResponse, strFunctionName)
    
    xmlCopyAttribute xmlResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, "TYPE"
    
    Dim xmlChildNode As IXMLDOMNode
    For Each xmlChildNode In xmlResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next

    gobjContext.SetComplete

CheckLocks_Exit:

    Set objAppManagerBO = Nothing
    Set nodeNewRequest = Nothing
    Set elemTemp = Nothing
    Set nodeAppLock = Nothing
    Set xmlResponse = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub

CheckLocksVbErr:
    
    gobjContext.SetAbort
    GoTo CheckLocks_Exit

End Sub

'PSC 19/09/2002 SYS4184  New procedure
Private Sub UpdateApplicationStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal strAppNo As String, ByVal strAppFFNo As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update Application Stage to "complete" if it is at the stage directly before
' pass:
'   Application Number, Application Fact Find Number
' return:
'   n/a
' history:
' SA    17/07/2002  Created.
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationStage_Exit

    Const strFunctionName As String = "UpdateApplicationStage"

    'BMIDS697 GHun 04/02/2004
    'Dim objTmBO As OmTmBO
    Dim objTmNoTxBO As omTmNoTxBO
    'BMIDS697 End
    Dim objAppBO As ApplicationBO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlGetStageRequest As IXMLDOMNode
    Dim xmlAppStageNode As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim strResponse As String
    Dim strCompletionStageId As String
    Dim intCompletionStage As Integer
    Dim strStageSeqNo As String
    Dim strDetails As String
    
    intCompletionStage = 0
        
    'BMIDS697 GHun 04/02/2004
    'Set objTmBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmBO")
    Set objTmNoTxBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTmNoTxBO")
    'BMIDS697 End
    
    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
        
    'Get the global parameter TMCompletionsStageId and if it's not empty then
    'proceed with moving to completions stage
    strCompletionStageId = GetGlobalParamString("TMCompletionsStageId")
    If Len(strCompletionStageId) > 0 Then
    
        ' Get the current stage of this application
        Set xmlGetStageRequest = vxmlRequestNode.cloneNode(False)
        Set xmlAppStageNode = xmlDoc.createElement("APPLICATIONSTAGE")
        Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlNode.Text = strAppNo
        xmlAppStageNode.appendChild xmlNode
        Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNode.Text = strAppFFNo
        xmlAppStageNode.appendChild xmlNode
        xmlGetStageRequest.appendChild xmlAppStageNode
        strResponse = objAppBO.GetCurrentApplicationStage(xmlGetStageRequest.xml)
        errCheckXMLResponse strResponse, True
        xmlDoc.loadXML strResponse
        Set xmlAppNode = xmlGetMandatoryNode(xmlDoc.documentElement, ".//APPLICATIONSTAGE")
        Set xmlNode = xmlGetMandatoryNode(xmlAppNode, "STAGESEQUENCENO")
        strStageSeqNo = xmlNode.Text 'Get Current Stage Sequence Number

        ' Compare current stage with completions stage
        intCompletionStage = GetGlobalParamAmount("TMCompletionsStage")
        If intCompletionStage = 0 Then
            'If there is no completion stage then Interface Fails
            Call WriteRejectReport
            errThrowError strFunctionName, oePPBatchGPCompStageNotFound
        End If
                    
        If CSafeInt(strStageSeqNo) = intCompletionStage - 1 Then
            ' Move to the "Completion Stage"
            ' If there are any outstanding mandatory tasks on the current stage
            ' this method will fail
            
            'Build MoveToStage Request
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStage"
            'CaseStage element of request
            Set xmlCaseStageNode = xmlDoc.createElement("CASESTAGE")
            xmlSetAttributeValue xmlCaseStageNode, "SOURCEAPPLICATION", "Omiga"
            xmlSetAttributeValue xmlCaseStageNode, "CASEID", strAppNo
            xmlSetAttributeValue xmlCaseStageNode, "ACTIVITYID", GetGlobalParamString("TMOmigaActivity")
            xmlSetAttributeValue xmlCaseStageNode, "ACTIVITYINSTANCE", "1"
            xmlSetAttributeValue xmlCaseStageNode, "STAGEID", strCompletionStageId
            xmlTempRequest.appendChild xmlCaseStageNode
            'Application element of request
            Set xmlAppNode = xmlDoc.createElement("APPLICATION")
            xmlSetAttributeValue xmlAppNode, "APPLICATIONNUMBER", strAppNo
            xmlSetAttributeValue xmlAppNode, "APPLICATIONFACTFINDNUMBER", strAppFFNo
            xmlTempRequest.appendChild xmlAppNode

            'Move to next stage
            'BMIDS697 GHun 04/02/2004
            'strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
            strResponse = objTmNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            'BMIDS697 GHun End
            
            If errCheckXMLResponse(strResponse, False) <> 0 Then
                Call WriteRejectReport
                
                ' PSC 08/01/2003 BM0138 - Start
                Set xmlDoc = xmlLoad(strResponse, strFunctionName)
                strDetails = xmlGetMandatoryNodeText(xmlDoc, "RESPONSE/ERROR/DESCRIPTION")
                errThrowError strFunctionName, oePPBatchMoveToCompStageFailed, strDetails
                ' PSC 08/01/2003 BM0138 - End

            End If
                
        ElseIf CSafeInt(strStageSeqNo) < intCompletionStage - 1 Then
            ' Interface failed - cannot miss out a stage in the process
            Call WriteRejectReport
            errThrowError strFunctionName, oePPBatchStageMissed
        End If
    End If
    
UpdateApplicationStage_Exit:
    Set objTmNoTxBO = Nothing
    Set objAppBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlGetStageRequest = Nothing
    Set xmlAppStageNode = Nothing
    Set xmlAppNode = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

' PSC 20/09/2002 BMIDS00480 - New function
Private Function IsADisbursementAwaitingResponse(ByVal vstrAppNo As String) As Boolean
'------------------------------------------------------------------------------------------
' description:
'   Determines if there are any disbursements with a payment status of Awaiting
'   Interface Response
' pass:
'   vstrAppNo       Application Number
' return:
'       True    If disbursements exist
'       False   If no disbursements exist
' exceptions:
'
'------------------------------------------------------------------------------------------

On Error GoTo IsADisbursementAwaitingResponseExit

    Const strFunctionName As String = "IsADisbursementAwaitingResponse"

    Dim xmlDisbursement As IXMLDOMNode
    Dim xmlResponse As IXMLDOMNode
    Dim xmlDisbursementList As IXMLDOMNodeList
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objPaymentProcessingDO As IPAymentProcessingDO
 
    Dim strValueId As String
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    ' Get value id for Awaiting Interface Response
    strValueId = GetFirstComboValueId("PaymentStatus", "R")
     
    Set xmlDisbursement = xmlDoc.createElement("DISBURSEMENTPAYMENT")
    xmlSetAttributeValue xmlDisbursement, "APPLICATIONNUMBER", vstrAppNo
    xmlSetAttributeValue xmlDisbursement, "PAYMENTSTATUS", strValueId
    
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    
    Set objPaymentProcessingDO = gobjContext.CreateInstance("omPayProc.PaymentProcessingDO")
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    On Error Resume Next
    objPaymentProcessingDO.FindDisbursementPaymentList xmlDisbursement, xmlResponse
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
    
    On Error GoTo IsADisbursementAwaitingResponseExit
    
    If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    
    Set xmlDisbursementList = xmlResponse.selectNodes("DISBURSEMENTPAYMENT")
    
    If xmlDisbursementList.length = 0 Then
        IsADisbursementAwaitingResponse = False
    Else
        IsADisbursementAwaitingResponse = True
    End If
       
IsADisbursementAwaitingResponseExit:
    
    Set xmlDisbursement = Nothing
    Set xmlResponse = Nothing
    Set xmlDisbursementList = Nothing
    Set objPaymentProcessingDO = Nothing
   
    ' Clean up section
    errCheckError strFunctionName, TypeName(Me)

End Function

'GD Added BMIDS0922
Private Sub DeleteApplicationLock(ByVal xmlTempRequest As IXMLDOMNode)


On Error GoTo DeleteApplicationLock_Exit

Const strFunctionName As String = "DeleteApplicationLock"

Dim objAppManBO As ApplicationManagerBO

'Dim xmlTempRequest As IXMLDOMNode
'Dim xmlAppNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40
'Dim xmlAppLockElem As IXMLDOMElement
Dim xmlAppLockElem As IXMLDOMElement
Dim xmlNode As IXMLDOMElement

Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
'''''''''''''''    'Create Request
'''''''''''''''    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
'''''''''''''''    Set xmlAppLockElem = xmlDoc.createElement("APPLICATIONLOCK")
'''''''''''''''    Set xmlAppNode = xmlDoc.createElement("APPLICATIONNUMBER")
'''''''''''''''    xmlAppNode.Text = strAppNo
'''''''''''''''    xmlAppLockElem.appendChild xmlAppNode
'''''''''''''''    xmlTempRequest.appendChild xmlAppLockElem
'''''''''''''''    'Set xmlAppNode = xmlDoc.createElement("APPLICATION")
    Set xmlAppLockElem = xmlTempRequest.selectSingleNode(".//APPLICATIONLOCK")
'''''''''''''''    'GD BMIDS0922 END
''''''''    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
''''''''    xmlNode.Text = strAppNo
''''''''    xmlAppNode.appendChild xmlNode
''''''''    xmlTempRequest.appendChild xmlAppNode
''''''''
''''''''    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
''''''''    xmlNode.Text = strAppFFNo
''''''''    xmlAppNode.appendChild xmlNode
''''''''
''''''''    'GD BMIDS0922 START
    Set xmlNode = xmlDoc.createElement("UNITID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "UNITID")
    xmlAppLockElem.appendChild xmlNode
''''''''
    Set xmlNode = xmlDoc.createElement("USERID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "USERID")
    xmlAppLockElem.appendChild xmlNode
''''''''
    Set xmlNode = xmlDoc.createElement("MACHINEID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "MACHINEID")
    xmlAppLockElem.appendChild xmlNode

    Set xmlNode = xmlDoc.createElement("LOCKDATE")
    xmlNode.Text = Now()
    xmlAppLockElem.appendChild xmlNode

    Set xmlNode = xmlDoc.createElement("TYPEOFLOCK")
    xmlNode.Text = "On"
    xmlAppLockElem.appendChild xmlNode


    Set xmlNode = xmlDoc.createElement("CHANNELID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "CHANNELID")
    xmlAppLockElem.appendChild xmlNode
    'Unlock the Application
    'Need to trap a no rows affected error in case same lock is tried to be unlocked more than once.
    Dim lngErrorNum As Long
    strResponse = objAppManBO.DeleteLock(xmlTempRequest.xml)
    lngErrorNum = errCheckXMLResponse(strResponse, False)
    
'    If lngErrorNum <> 0 Then
'        lngErrorNum = errGetOmigaErrorNumber(lngErrorNum)
'        If lngErrorNum <> oeNoRowsAffected Then
'            errThrowError strFunctionName, lngErrorNum
'        End If
'    End If
    
    
DeleteApplicationLock_Exit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppLockElem = Nothing
    Set xmlDoc = Nothing
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub
'GD BMIDS0922
Private Sub LockApplicationOnly(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String, ByVal strAppFFNo As String)


On Error GoTo LockApplication_Exit

Const strFunctionName As String = "LockApplicationOnly"

Dim objAppManBO As ApplicationManagerBO

'Dim xmlTempRequest As IXMLDOMNode
Dim xmlAppNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlDoc As FreeThreadedDOMDocument40

Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'Create Request
    'GD BMIDS0922 START
    'Set xmlAppNode = xmlDoc.createElement("APPLICATION")
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONLOCK")
    'GD BMIDS0922 END
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = strAppNo
    xmlAppNode.appendChild xmlNode
    xmlTempRequest.appendChild xmlAppNode
    
    Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = strAppFFNo
    xmlAppNode.appendChild xmlNode
    
    'GD BMIDS0922 START
    Set xmlNode = xmlDoc.createElement("UNITID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "UNITID")
    xmlAppNode.appendChild xmlNode
    
    Set xmlNode = xmlDoc.createElement("USERID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "USERID")
    xmlAppNode.appendChild xmlNode
    
    Set xmlNode = xmlDoc.createElement("MACHINEID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "MACHINEID")
    xmlAppNode.appendChild xmlNode
    
    Set xmlNode = xmlDoc.createElement("LOCKDATE")
    xmlNode.Text = Now()
    xmlAppNode.appendChild xmlNode
    
    Set xmlNode = xmlDoc.createElement("TYPEOFLOCK")
    xmlNode.Text = "On"
    xmlAppNode.appendChild xmlNode
    
    
    Set xmlNode = xmlDoc.createElement("CHANNELID")
    xmlNode.Text = xmlGetAttributeText(xmlTempRequest, "CHANNELID")
    xmlAppNode.appendChild xmlNode
    
    

    
    'Lock the Application and associated customers

    'strResponse = objAppManBO.LockCustomersForApplication(xmlTempRequest.xml)
    strResponse = objAppManBO.CreateLock(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True

LockApplication_Exit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

' PSC 14/11/2005 MAR520 - Start
Private Sub UpdateAccountNo(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo UpdateAccountNoVbErr
    Const strFunctionName = "UpdateAccountNo"
 
    Dim nodeNewRequest As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlApplicationNo As IXMLDOMNode
    Dim xmlOtherSysNo As IXMLDOMNode
    Dim xmlHeader As IXMLDOMNode
    Dim objAppBO As ApplicationBO

    Dim strOtherSystemNo As String
    Dim strApplicationNo As String
    Dim strResponse As String
    
    Set xmlHeader = xmlGetMandatoryNode(vxmlRequestNode, "HEADER")
    strApplicationNo = xmlGetMandatoryAttributeText(xmlHeader, "APPLICATIONNUMBER")
    strOtherSystemNo = xmlGetMandatoryAttributeText(xmlHeader, "ACCOUNTNUMBER")

    ' Set up request to update Application
    Set nodeNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApplication = vxmlRequestNode.ownerDocument.createElement("APPLICATION")
    Set xmlApplicationNo = vxmlRequestNode.ownerDocument.createElement("APPLICATIONNUMBER")
    Set xmlOtherSysNo = vxmlRequestNode.ownerDocument.createElement("OTHERSYSTEMACCOUNTNUMBER")
    
    nodeNewRequest.appendChild xmlApplication
    xmlApplication.appendChild xmlApplicationNo
    xmlApplication.appendChild xmlOtherSysNo
    xmlApplicationNo.Text = strApplicationNo
    
    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    ' Update application with UFSS account number if it has been returned
    ' PSC 08/04/2004 BBG164
    If Len(Trim$(strOtherSystemNo)) <> 0 Then
        xmlOtherSysNo.Text = strOtherSystemNo
        strResponse = objAppBO.UpdateApplication(nodeNewRequest.xml)
        errCheckXMLResponse strResponse, True
    End If
    ' PSC 02/03/2004 BBG80 - End
    
    
UpdateAccountNoExit:
    
    Set nodeNewRequest = Nothing
    Set xmlApplication = Nothing
    Set xmlApplicationNo = Nothing
    Set xmlOtherSysNo = Nothing
    Set xmlHeader = Nothing
    Set objAppBO = Nothing

    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
UpdateAccountNoVbErr:
    gobjContext.SetAbort
    GoTo UpdateAccountNoExit
End Sub
' PSC 14/11/2005 MAR520 - End


