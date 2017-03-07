VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BatchScheduleBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BatchScheduleBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Batch Schedule Business Object which 'supports transactions' only
'               Methods which do not require transaction support reside in this
'               class. Any methods that require transactions will be delegated to
'               BatchScheduleTxBO
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'Dependencies:  BatchScheduleTxBO, BatchScheduleDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     02/04/01    SYS2201: Created
'MV     01/05/01    Added New Methods Called IncrementBatchAuditRetries and FindBatchAuditList
'MV     11/05/01    Changed RunBatch and LaunchBatch methods to Amend the USERID ,UNITID,MACHINEID,OPERATION
'                   in the Request
'MC     05/12/01    SYS3018 - Fixes to Batch Process for SQL Server.
'DM     05/12/2001  SYS3185 Add CreateBatchException
'MC     20/12/01    SYS3541 - Write error info to event log
'PSC    12/02/02    SYS4039 - Amend OnMessage to go through  omBatchRequest. Amend omBatchRequest
'                   to check for MTS
'RF     20/02/02    SYS4079 Correct the validation in FindBatchAuditList
'SA     26/02/02    SYS4176 Add new method PrintDocuments
'SA     20/03/02    SYS4287 Print Rate Change Letters
'SA     05/04/02    SYS4340 Call PrintLetter for rate change letter rather than PrintDocuments
'SA     08/04/02    SYS4340 Check for letter string changed
'SA     10/04/02    SYS4379 Destination Type needs to be sent through as a validation type not numeric value
'SA     18/04/02    SYS4412 As above - but for reports!!
'CL     13/05/02    SYS4565 Changes to 'DoRequest' and addition of function 'DoClientRequest'
'PSC    14/11/05    MAR520  Amend PrintDocuments to only try printing report if a template is specified
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements MessageQueueComponentVC2

Private gobjContext As ObjectContext

Private gobjTxBO As IBatchScheduleTXBO
Private gobjDO As IBatchScheduleDO


'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MDC    16/09/2002  BMIDS00048 BM026 RCWP1 Rate Change Letter Printing
'GHun   02/04/2003  BM0438 Added GetBatchScheduleProgress & amended DoRequest to call it
'INR    13/10/2003  BMIDS628 Call BatchScheduleNTTxBO.cls so it uses a new transaction for
'                   IncrementBatchScheduleSuccess and IncrementBatchScheduleFailure
'------------------------------------------------------------------------------------------
'BMIDS628
Private gobjNTTxBO As IBatchScheduleNTTXBO

Private Sub GetBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo GetBatchExit
    
    'Constant
    Const strFunctionName As String = "GetBatch"
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    gobjDO.GetBatch vxmlRequestNode, vxmlResponseNode
    
GetBatchExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ProduceBatchScheduleReport(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo ProduceBatchReportExit
    
    Const strFunctionName As String = "ProduceBatchScheduleReport"
    
    Dim xmlBatchNode As IXMLDOMNode
    
    Set xmlBatchNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchNode, "BATCHRUNNUMBER"
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    gobjDO.ProduceBatchScheduleReport vxmlRequestNode, vxmlResponseNode
    
    
ProduceBatchReportExit:
    
    Set xmlBatchNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub UpdateBatchSchedule(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateBatchScheduleExit
    
    Const strFunctionName As String = "UpdateBatchSchedule"
    
    Dim xmlBatchScheduleRequestNode  As IXMLDOMNode
    
    Set xmlBatchScheduleRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHRUNNUMBER"

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.UpdateBatchSchedule vxmlRequestNode, vxmlResponseNode
    
UpdateBatchScheduleExit:
    
    Set xmlBatchScheduleRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

 Private Sub UpdateBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateBatchExit
    
    Const strFunctionName As String = "UpdateBatch"
    
    Dim xmlBatchRequestNode  As IXMLDOMNode
    
    Set xmlBatchRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    xmlCheckMandatoryAttribute xmlBatchRequestNode, "BATCHNUMBER"
    

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.UpdateBatch vxmlRequestNode, vxmlResponseNode
    
UpdateBatchExit:
    
    Set xmlBatchRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub SetBatchStatus(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo SetBatchStatusExit
    
    Const strFunctionName As String = "SetBatchStatus"
    
    Dim xmlBatchScheduleRequestNode As IXMLDOMNode
    
    Set xmlBatchScheduleRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHRUNNUMBER"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.SetBatchStatus vxmlRequestNode, vxmlResponseNode
    
SetBatchStatusExit:
    
    Set xmlBatchScheduleRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IncrementBatchScheduleSuccess(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo IncrementBatchScheduleSuccessExit
    
    Const strFunctionName As String = "IncrementBatchScheduleSuccess"
    
    Dim xmlBatchScheduleNode As IXMLDOMNode
    
    'Store into Local Variables
    Set xmlBatchScheduleNode = vxmlRequestNode.selectSingleNode(".//BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchScheduleNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleNode, "BATCHRUNNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleNode, "BATCHAUDITGUID"
    
    'INR BMIDS628 Call BatchScheduleNTTxBO.cls so it uses a new transaction
    If gobjNTTxBO Is Nothing Then
        Set gobjNTTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleNTTxBO")
    End If
    
    gobjNTTxBO.IncrementBatchScheduleSuccess vxmlRequestNode, vxmlResponseNode
    
IncrementBatchScheduleSuccessExit:
    
    errCheckError strFunctionName, TypeName(Me)



End Sub

Private Sub IncrementBatchScheduleFailure(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo IncrementBatchScheduleFailureExit
    
    Const strFunctionName As String = "IncrementBatchScheduleFailure"
    
    Dim xmlBatchScheduleRequestNode As IXMLDOMNode
    
    Set xmlBatchScheduleRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHRUNNUMBER"
    
    'INR BMIDS628 Call BatchScheduleNTTxBO.cls so it uses a new transaction
    If gobjNTTxBO Is Nothing Then
        Set gobjNTTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleNTTxBO")
    End If
    
    gobjNTTxBO.IncrementBatchScheduleFailure vxmlRequestNode, vxmlResponseNode
    
IncrementBatchScheduleFailureExit:
    
    Set xmlBatchScheduleRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub GetBatchSchedule(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetBatchScheduleExit
    
    Const strFunctionName As String = "GetBatchSchedule"
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    gobjDO.GetBatchSchedule vxmlRequestNode, vxmlResponseNode
    
GetBatchScheduleExit:
    
    'Set xmlBatchScheduleNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Function CreateBatchAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As IXMLDOMNode
    
    'Error Handler
    On Error GoTo CreateBatchAuditExit
    
    'Constant
    Const strFunctionName As String = "CreateBatchAudit"
    
    Dim xmlBatchAuditRequestNode As IXMLDOMNode
    Dim xmlTempResponseNode  As IXMLDOMNode
    
    'Check for the Mandatory Nodes and Attributes
    Set xmlBatchAuditRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    xmlCheckMandatoryAttribute xmlBatchAuditRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchAuditRequestNode, "BATCHRUNNUMBER"
    
    'Create an Instance
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    Set xmlTempResponseNode = gobjTxBO.CreateBatchAudit(vxmlRequestNode)
    vxmlResponseNode.appendChild xmlTempResponseNode
    Set CreateBatchAudit = vxmlResponseNode
    
CreateBatchAuditExit:
    
    Set xmlBatchAuditRequestNode = Nothing
    Set xmlTempResponseNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Function

Private Sub UpdateBatchAudit(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateBatchAuditExit
    
    Const strFunctionName As String = "UpdateBatchAudit"
    
    Dim xmlBatchAuditRequestNode As IXMLDOMNode
    
    Set xmlBatchAuditRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    xmlCheckMandatoryAttribute xmlBatchAuditRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchAuditRequestNode, "BATCHRUNNUMBER"
    xmlCheckMandatoryAttribute xmlBatchAuditRequestNode, "BATCHAUDITGUID"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.UpdateBatchAudit vxmlRequestNode, vxmlResponseNode
    
UpdateBatchAuditExit:
    Set xmlBatchAuditRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub LaunchBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo LaunchBatchExit
    
    Const strFunctionName As String = "LaunchBatch"
    
    Dim xmlBatchRequestNode As IXMLDOMNode
    
    Set xmlBatchRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    xmlCheckMandatoryAttribute xmlBatchRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchRequestNode, "EXECUTIONDATETIME"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.LaunchBatch vxmlRequestNode, vxmlResponseNode
    
LaunchBatchExit:
    
    Set xmlBatchRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub RunBatch(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo RunBatchExit
    
    'Constant
    Const strFunctionName As String = "RunBatch"
    
    Dim xmlBatchRequestNode As IXMLDOMNode
    'SYS3541 MDC 20/12/2001
    Dim strErrMsg As String
    Dim strBatchNumber As String
    'SYS3541 End
    
    Set xmlBatchRequestNode = xmlGetMandatoryNode(vxmlRequestNode, ".//BATCH")
    
    'SYS3541 MDC 20/12/2001
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchRequestNode, "BATCHNUMBER")
    'xmlCheckMandatoryAttribute xmlBatchRequestNode, "BATCHNUMBER"
    'SYS3541 End

    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.RunBatch vxmlRequestNode, vxmlResponseNode
    
RunBatchExit:
    
    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    'SYS3541 End
    
    Set xmlBatchRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Function MessageQueueComponentVC2_OnMessage(ByVal in_xmlConfig As String, ByVal in_xmlData As String) As Long
        
    Const strFunctionName As String = "MessageQueueComponentVC2_OnMessage"
    
    ' PSC 12/02/02 SYS4039 - Start
    Dim strResponse As String
    Dim lngErrorNo As Long
    
    strResponse = omBatchRequest(in_xmlData)
    lngErrorNo = errCheckXMLResponse(strResponse)
    
IMessageQueueComponentVC2_OnMessageExit:

    If lngErrorNo = 0 Then
        MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
    Else
        App.LogEvent "Message could not be processed " & vbCrLf & _
                     "Request: " & in_xmlData & vbCrLf & _
                     "Respose: " & strResponse, vbLogEventTypeError
                     
        MessageQueueComponentVC2_OnMessage = MESSQ_RESP_RETRY_MOVE_MESSAGE
    End If
    
' PSC 12/02/02 SYS4039 - End
    
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

Public Function omBatchRequest(ByVal vstrXmlIn As String) As String

    On Error GoTo omBatchRequestVbErr

    Const strFunctionName As String = "omBatchRequest"

    Dim xmlInDoc As New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    Dim xmlOutDoc As New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute

  ' PSC 12/02/02 SYS4039 - Start
    If gobjContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If
    ' PSC 12/02/02 SYS4039 - End

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

    omBatchRequest = xmlOutDoc.xml


omBatchRequestExit:

    Set gobjTxBO = Nothing
    Set gobjDO = Nothing

    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing

    Exit Function

omBatchRequestVbErr:

    omBatchRequest = errCreateErrorResponse()

    GoTo omBatchRequestExit

End Function

Private Sub DoRequest(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
    
        Case "LAUNCHBATCH"
            LaunchBatch vxmlRequestNode, vxmlResponseNode
        
        Case "RUNBATCH"
            RunBatch vxmlRequestNode, vxmlResponseNode
            
        Case "GETBATCH"
            GetBatch vxmlRequestNode, vxmlResponseNode
        
        Case "UPDATEBATCH"
            UpdateBatch vxmlRequestNode, vxmlResponseNode
              
        Case "SETBATCHSTATUS"
            SetBatchStatus vxmlRequestNode, vxmlResponseNode
            
        Case "CREATEBATCHSCHEDULE"
            CreateBatchSchedule vxmlRequestNode, vxmlResponseNode
    
        Case "UPDATEBATCHSCHEDULE"
            UpdateBatchSchedule vxmlRequestNode, vxmlResponseNode
    
        Case "GETBATCHSCHEDULE"
            GetBatchSchedule vxmlRequestNode, vxmlResponseNode
            
        Case "CREATEBATCHAUDIT"
            CreateBatchAudit vxmlRequestNode, vxmlResponseNode
    
        Case "UPDATEBATCHAUDIT"
            UpdateBatchAudit vxmlRequestNode, vxmlResponseNode
            
        Case "PRODUCEBATCHSCHEDULEREPORT"
            ProduceBatchScheduleReport vxmlRequestNode, vxmlResponseNode
    
        Case "INCREMENTBATCHSCHEDULESUCCESS"
            IncrementBatchScheduleSuccess vxmlRequestNode, vxmlResponseNode
    
        Case "INCREMENTBATCHSCHEDULEFAILURE"
            IncrementBatchScheduleFailure vxmlRequestNode, vxmlResponseNode
            
        Case "INCREMENTBATCHAUDITRETRIES"
            IncrementBatchAuditRetries vxmlRequestNode, vxmlResponseNode
            
        Case "FINDBATCHAUDITLIST"
            FindBatchAuditList vxmlRequestNode, vxmlResponseNode
                        
        'SYS3018 MDC 04/12/2001
        Case "COMPLETERUNBATCH"
            CompleteRunBatch vxmlRequestNode, vxmlResponseNode
        'SYS3018 End
        
        'DM     05/12/2001 SYS3185 Add CreateBatchException
        Case "CREATEBATCHEXCEPTION"
            CreateBatchException vxmlRequestNode, vxmlResponseNode
        
        'SA SYS4176 Add PrintDocuments
        Case "PRINTDOCUMENTS"
            PrintDocuments vxmlRequestNode, vxmlResponseNode
            
        'BM0438 GHun 02/04/2003
        Case "GETBATCHSCHEDULEPROGRESS"
            GetBatchScheduleProgress vxmlRequestNode, vxmlResponseNode
        'BM0438 End
        
        Case Else
            DoClientRequest strOperation, vxmlRequestNode, vxmlResponseNode
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CompleteRunBatch(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CompleteRunBatchExit

Const strFunctionName As String = "CompleteRunBatch"
    
'SYS3541 MDC 20/12/2001
Dim strBatchNumber As String
Dim strErrMsg As String
Dim xmlBatchNode As IXMLDOMNode
'SYS3541 End

    xmlCheckMandatoryNode vxmlRequestNode, ".//BATCH"
        
    'SYS3541 MDC 20/12/2001
    Set xmlBatchNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCH")
    strBatchNumber = xmlGetAttributeText(xmlBatchNode, "BATCHNUMBER")
    'SYS3541 End
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    gobjTxBO.CompleteRunBatch vxmlRequestNode, vxmlResponseNode
    
CompleteRunBatchExit:
    
    'SYS3541 MDC 20/12/2001
    If Err.Number <> 0 Then
        strErrMsg = vbCrLf & vbCrLf & "Error logged in " & TypeName(Me) & "_" & strFunctionName
        strErrMsg = strErrMsg & vbCrLf & "Batch Number: " & strBatchNumber
        strErrMsg = strErrMsg & vbCrLf & "Error Number: " & Err.Number
        strErrMsg = strErrMsg & vbCrLf & "Error Source: " & Err.Source
        strErrMsg = strErrMsg & vbCrLf & "Error Desc: " & Err.Description
        App.LogEvent strErrMsg, vbLogEventTypeError
    End If
    
    Set xmlBatchNode = Nothing
    'SYS3541 End
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub FindBatchAuditList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo FindBatchAuditListExit
    
    'Constant
    Const strFunctionName As String = "FindBatchAuditList"
    
    Dim xmlThisRequestNode As IXMLDOMNode
    
    'Check for the Mandatory Node
    Set xmlThisRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    
    ' RF 20/02/02 SYS4079 These attributes are not mandatory
    'xmlCheckMandatoryAttribute xmlThisRequestNode, "BATCHNUMBER"
    'xmlCheckMandatoryAttribute xmlThisRequestNode, "BATCHRUNNUMBER"
    'xmlCheckMandatoryAttribute xmlThisRequestNode, "BATCHAUDITGUID"
    
    'Create an Instance
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    gobjDO.FindBatchAuditList vxmlRequestNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

FindBatchAuditListExit:

    Set xmlThisRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateBatchSchedule(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo CreateBatchScheduleExit
    
    'Constant
    Const strFunctionName As String = "CreateBatchSchedule"
    
    Dim xmlBatchScheduleRequestNode As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    
    'Check for the Mandatory Nodes and Attributes
    Set xmlBatchScheduleRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULE")
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHNUMBER"
    
    'If the CreationDateTime is NULL or Empty then pick up as current Date and time
    If IsNull(xmlGetAttributeText(xmlBatchScheduleRequestNode, "CREATIONDATETIME")) Or (xmlGetAttributeText(xmlBatchScheduleRequestNode, "CREATIONDATETIME") = "") Then
        xmlBatchScheduleRequestNode.setAttribute "CREATIONDATETIME", Now()
    End If
    
    'Create an Instance
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.CreateBatchSchedule vxmlRequestNode, vxmlResponseNode
    
    Set xmlTempNode = vxmlRequestNode.selectSingleNode(".//BATCHSCHEDULE")
    vxmlResponseNode.appendChild xmlTempNode
    
CreateBatchScheduleExit:
    
    Set xmlBatchScheduleRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IncrementBatchAuditRetries(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo IncrementBatchAuditRetriesExit
    
    Const strFunctionName As String = "IncrementBatchAuditRetries"
    
    Dim xmlBatchScheduleRequestNode As IXMLDOMNode
    
    Set xmlBatchScheduleRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHRUNNUMBER"
    xmlCheckMandatoryAttribute xmlBatchScheduleRequestNode, "BATCHAUDITGUID"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.IncrementBatchAuditRetries vxmlRequestNode, vxmlResponseNode
    
IncrementBatchAuditRetriesExit:
    
    Set xmlBatchScheduleRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateBatchException(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    ' Call the DO and possibly the TXBO to forward this call.
    Const strFunctionName As String = "CreateBatchException"
    On Error GoTo CreateBatchExceptionExit
    
    ' Check for any mandatory attributes. There might be more than this.
    Dim xmlBatchExceptionRequestNode As IXMLDOMNode
    
    Set xmlBatchExceptionRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHEXCEPTION")
    xmlCheckMandatoryAttribute xmlBatchExceptionRequestNode, "BATCHNUMBER"
    xmlCheckMandatoryAttribute xmlBatchExceptionRequestNode, "BATCHRUNNUMBER"
    xmlCheckMandatoryAttribute xmlBatchExceptionRequestNode, "BATCHAUDITGUID"
    
    
    
    ' Use the commented out code if we do not need to operate within a transaction.
'    If gobjDO Is Nothing Then
'        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
'    End If
'
'    gobjDO.CreateBatchException vxmlRequestNode, vxmlResponseNode
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance(App.Title & ".BatchScheduleTxBO")
    End If
    
    gobjTxBO.CreateBatchException vxmlRequestNode, vxmlResponseNode
    
CreateBatchExceptionExit:
    
    Set xmlBatchExceptionRequestNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub PrintDocuments(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo PrintDocuments
    
    'Constant
    Const strFunctionName As String = "PrintDocuments"
    
    Dim strProgramType As String
    Dim strReportTemplateId As String
    Dim strLetterTemplateId As String
    Dim strResponse As String
    Dim strValidationType As String
    
    Dim objPrintBO As omPrintBO
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlPrintAttribs As IXMLDOMNode
    Dim xmlBatchAttribs As IXMLDOMNode
    Dim xmlControlNode As IXMLDOMNode
    Dim xmlPrintNode As IXMLDOMNode
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlPrintRequest As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strPrinterDestination As String 'SYS4379
    'BMIDS00048 MDC 16/09/2002
    Dim blnPrintLetters As Boolean
    'BMIDS00048 MDC 16/09/2002 - End
            
    'Get the batch record so we can determine type of batch running
    GetBatch vxmlRequestNode, vxmlResponseNode

    strProgramType = xmlGetMandatoryAttributeText(vxmlResponseNode.selectSingleNode(".//BATCH"), "PROGRAMTYPE")
    strValidationType = GetValidationTypeForValueID("BatchProgramType", CInt(strProgramType))
    
    'BMIDS00048 MDC 16/09/2002
    blnPrintLetters = True
    'BMIDS00048 MDC 16/09/2002 - End
    
    'Depending on the type of batch, get the relevant template id.
    Select Case strValidationType
        Case "R"   'Rate Change
            'For rate change we have both a letter and report.
            strReportTemplateId = GetGlobalParamString("RCHANGEREPORTTEMPLATEID")
            strLetterTemplateId = GetGlobalParamString("RCHANGELETTERTEMPLATEID")
            'BMIDS00048 MDC 16/09/2002
            blnPrintLetters = GetGlobalParamBoolean("RATECHANGEPRINTINDICATOR")
            'BMIDS00048 MDC 16/09/2002 - End

        Case "P"   'Payment Processing
            strReportTemplateId = GetGlobalParamString("PPROCTEMPLATEID")
        Case "V"   'Valuer Payments
           strReportTemplateId = GetGlobalParamString("VALPAYMENTTEMPLATEID")
    End Select
    
    ' PSC 14/11/2005 MAR520 - Start
    If Len(strReportTemplateId) > 0 Then

        'First - deal with the reports (letters will come later)
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        
        'Create request to get the Printer Attributes
        Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
        xmlChangeNodeName xmlTempRequest, "BATCH", "BATCHSCHEDULE"
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
        Set xmlPrintAttribs = xmlDoc.createElement("FINDATTRIBUTES")
        xmlSetAttributeValue xmlPrintAttribs, "HOSTTEMPLATEID", strReportTemplateId
        xmlTempRequest.appendChild xmlPrintAttribs
    
        'Get the Printer Attributes
        Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        strResponse = objPrintBO.omRequest(xmlTempRequest.xml)
    
        'Check for errors
        Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponse xmlTempResponseDoc.xml, True
    
        Set xmlPrintAttribs = Nothing
        
        'Select the Printer Attributes element to copy values into request for printing document
        Set xmlPrintAttribs = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
        
        'Select the Batchnumber and batchrunnumber (now Batch as we've changed the node name for call to GetPrintAttributes)
        Set xmlBatchAttribs = xmlGetMandatoryNode(vxmlRequestNode, ".//BATCH")
            
        'Now print the document
        'Build up request string by copying Printer attributes
        Set xmlControlNode = xmlDoc.createElement("CONTROLDATA")
        xmlCopyMandatoryAttributeValue xmlPrintAttribs, xmlControlNode, "HOSTTEMPLATEID", "DOCUMENTID"
        xmlCopyMandatoryAttributeValue xmlPrintAttribs, xmlControlNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
        
        If Not xmlPrintAttribs Is Nothing Then
            'Number of copies
            xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "DEFAULTCOPIES", "COPIES"
        
            'Printer Location
            xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "REMOTEPRINTERLOCATION", "PRINTER"
          
            'SYS4412 Send through Validation Type not numeric value
            'Destination Type
            'xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "PRINTERDESTINATIONTYPE", "DESTINATIONTYPE"
            strPrinterDestination = xmlGetAttributeText(xmlPrintAttribs, "PRINTERDESTINATIONTYPE")
            strPrinterDestination = GetValidationTypeForValueID("PrinterDestination", CInt(strPrinterDestination))
            xmlSetAttributeValue xmlControlNode, "DESTINATIONTYPE", strPrinterDestination
                 
        
            'Add PrintData Element and set attributes
            Set xmlPrintNode = xmlDoc.createElement("PRINTDATA")
            
            'Method Name (in PDM)
            xmlCopyAttributeValue xmlPrintAttribs, xmlPrintNode, "PDMMETHOD", "METHODNAME"
            
        End If
           
        If Not xmlBatchAttribs Is Nothing Then
            'Batch Number
            xmlCopyAttributeValue xmlBatchAttribs, xmlPrintNode, "BATCHNUMBER", "BATCHNUMBER"
            'Batch Run Number
            xmlCopyAttributeValue xmlBatchAttribs, xmlPrintNode, "BATCHRUNNUMBER", "BATCHRUNNUMBER"
        End If
            
        Set xmlPrintRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlPrintRequest, "OPERATION", "PrintDocument"
        xmlPrintRequest.appendChild xmlControlNode
        xmlPrintRequest.appendChild xmlPrintNode
        
        'Send to print
        'Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        strResponse = objPrintBO.omRequest(xmlPrintRequest.xml)
        errCheckXMLResponse strResponse, True
    End If
    ' PSC 14/11/2005 MAR520 - End
    
    'We may or may not have a letter template id (currently only used for ratechange)
    ' If we do, we need to do the same again.
    'BMIDS00048 MDC 16/09/2002 - Check if we need to print letters
    'If strLetterTemplateId <> "" Then
    If strLetterTemplateId <> "" And blnPrintLetters Then
    'BMIDS00048 MDC 16/09/2002 - End
        
        'Create request to get the Printer Attributes
        Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
        xmlChangeNodeName xmlTempRequest, "BATCH", "BATCHSCHEDULE"
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetPrintAttributes"
        Set xmlPrintAttribs = xmlDoc.createElement("FINDATTRIBUTES")
        xmlSetAttributeValue xmlPrintAttribs, "HOSTTEMPLATEID", strLetterTemplateId
        xmlTempRequest.appendChild xmlPrintAttribs
    
        'Get the Printer Attributes
        'Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        strResponse = objPrintBO.omRequest(xmlTempRequest.xml)
    
        'Check for errors
        Set xmlTempResponseDoc = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponse xmlTempResponseDoc.xml, True
    
        Set xmlPrintAttribs = Nothing
        
        'Select the Printer Attributes element to copy values into request for printing document
        'Set xmlPrintAttribs = xmlGetMandatoryNode(xmlTempResponseDoc, ".//ATTRIBUTES")
        Set xmlPrintAttribs = xmlTempResponseDoc.selectSingleNode(".//ATTRIBUTES")
                
        'Select the Batchnumber and batchrunnumber (now Batch as we've changed the node name for call to GetPrintAttributes)
        Set xmlBatchAttribs = xmlGetMandatoryNode(vxmlRequestNode, ".//BATCH")
            
        'Now print the document
        'Build up request string by copying Printer attributes
        Set xmlControlNode = xmlDoc.createElement("CONTROLDATA")
        xmlCopyMandatoryAttributeValue xmlPrintAttribs, xmlControlNode, "HOSTTEMPLATEID", "DOCUMENTID"
        xmlCopyMandatoryAttributeValue xmlPrintAttribs, xmlControlNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
        
        If Not xmlPrintAttribs Is Nothing Then
            'Number of copies
            xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "DEFAULTCOPIES", "COPIES"
        
            'Printer Location
            xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "REMOTEPRINTERLOCATION", "PRINTER"
        
            'Destination Type
            'SYS4379 Send through Validation Type not numeric value
            'xmlCopyAttributeValue xmlPrintAttribs, xmlControlNode, "PRINTERDESTINATIONTYPE", "DESTINATIONTYPE"
                
            strPrinterDestination = xmlGetAttributeText(xmlPrintAttribs, "PRINTERDESTINATIONTYPE")
            strPrinterDestination = GetValidationTypeForValueID("PrinterDestination", CInt(strPrinterDestination))
            xmlSetAttributeValue xmlControlNode, "DESTINATIONTYPE", strPrinterDestination
            
            'Add PrintData Element and set attributes
            Set xmlPrintNode = xmlDoc.createElement("PRINTDATA")
            
            'Method Name (in PDM)
            xmlCopyAttributeValue xmlPrintAttribs, xmlPrintNode, "PDMMETHOD", "METHODNAME"
            
        End If
           
        If Not xmlBatchAttribs Is Nothing Then
            'Batch Number
            xmlCopyAttributeValue xmlBatchAttribs, xmlPrintNode, "BATCHNUMBER", "BATCHNUMBER"
            'Batch Run Number
            xmlCopyAttributeValue xmlBatchAttribs, xmlPrintNode, "BATCHRUNNUMBER", "BATCHRUNNUMBER"
        End If
            
        Set xmlPrintRequest = vxmlRequestNode.cloneNode(False)
        'SYS4340 Change to call PrintLetters
        xmlSetAttributeValue xmlPrintRequest, "OPERATION", "PrintLetters"
        xmlPrintRequest.appendChild xmlControlNode
        xmlPrintRequest.appendChild xmlPrintNode
        
        'Send to print
        'If objPrintBO Is Nothing Then
        '    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        'End If
        strResponse = objPrintBO.omRequest(xmlPrintRequest.xml)
        errCheckXMLResponse strResponse, True

    End If
    
PrintDocuments:
    
    Set objPrintBO = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlPrintAttribs = Nothing
    Set xmlBatchAttribs = Nothing
    Set xmlControlNode = Nothing
    Set xmlPrintNode = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlPrintRequest = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Function DoClientRequest(ByVal strOperation As String, ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo DoClientRequestExit
    
    Const strFunctionName As String = "DoClientRequest"

    Select Case strOperation
    'Enter Customer specific case details here
    
    Case Else
        errThrowError _
        strFunctionName, _
        oeNotImplemented, _
        strOperation
    End Select

DoClientRequestExit:
         
    errCheckError strFunctionName, TypeName(Me)

End Function

'BM0438
Private Sub GetBatchScheduleProgress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const strFunctionName As String = "GetBatchScheduleProgress"
    
On Error GoTo GetBatchScheduleProgressExit

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance(App.Title & ".BatchScheduleDO")
    End If
    
    gobjDO.GetBatchScheduleProgress vxmlRequestNode, vxmlResponseNode
    
GetBatchScheduleProgressExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub
'BM0438 End
