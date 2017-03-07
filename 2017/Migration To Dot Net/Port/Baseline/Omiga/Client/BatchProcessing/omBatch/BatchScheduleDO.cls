VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BatchScheduleDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BatchScheduleDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Dependencies:  BatchScheduleDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MV     02/04/01    SYS2201: Created
'MV     01/05/01    Added New Method Called FindBatchAuditList
'PSC    02/05/01    Amend CallBatchProcessByName to set output correctly
'MV     02/05/01    Amend BatchStatus in returning GetBatch info in RunBatch method
'PSC    02/05/01    Amend BatchExecutionDateTime in returning GetBatch info in RunBatch method
'DM     06/12/01    SYS3185 Implement CreateBatchExecption
'MC     14/12/01    SYS3440. Fix SetBatchStatus and produce audit report
'PSC    20/02/02    SYS4097 Amend SetBatchStatus to set the status sent in and
'                   write report if status is also "IE"
'                   Correct FindBatchAuditList
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IBatchScheduleDO

Private gobjContext As ObjectContext


'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    09/01/2003  BM0230 - Make BatchRunNumber optional in GetBatchSchedule
'GHun   24/03/2003  BM0438 - Added IncrementBatchScheduleSuccess & Failure, CreateBatchScheduleProgress
'                            GetBatchScheduleProgress, UpdateBatchScheduleProgress, IncrementBatchScheduleProgressTotal
'                            & amended SetBatchStatus
'GHun   09/09/2003  BMIDS622 RateChange restructure and clean up
'GHun   05/11/2003  BMIDS622 Amended ProduceBatchScheduleReport to better handle a missing log path
'------------------------------------------------------------------------------------------

Private Function IBatchScheduleDO_CreateBatchAudit(ByVal vxmlRequestParentNode As IXMLDOMNode) As IXMLDOMNode
    'Error Handler
    On Error GoTo CreateBatchAuditExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_CreateBatchAudit"
    
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlBatchAudit As IXMLDOMNode
    
    Set xmlBatchAudit = vxmlRequestParentNode.selectSingleNode(".//BATCHAUDIT")
    adoCreateFromNode xmlBatchAudit, "BATCHAUDIT"
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlTempElement = xmlOut.createElement("BATCHAUDIT")
    xmlTempElement.setAttribute "BATCHAUDITGUID", xmlGetAttributeText(xmlBatchAudit, "BATCHAUDITGUID")
        
    Set IBatchScheduleDO_CreateBatchAudit = xmlTempElement

CreateBatchAuditExit:
    gobjContext.SetComplete
    Set xmlBatchAudit = Nothing
    Set xmlTempElement = Nothing
    Set xmlOut = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Function

Private Sub IBatchScheduleDO_CreateBatchException(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    ' Create and use a stored procedure that updates the table required.

    On Error GoTo CreateBatchExceptionVbErr
    
    Dim adoParameterExceptionNumber As ADODB.Parameter
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim xmlSequenceAttribute As IXMLDOMAttribute
    Dim xmlBatchExceptionNode As IXMLDOMNode
    Const ProcName = "USP_GETNEXTBATCHEXCEPTIONNUM"
    Const OutSize = 12
    
    Const strFunctionName As String = "IBatchScheduleDO_CreateBatchException"
    
    ' Get a sequence number
    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    
    adoConnection.Open
        
    Set adoCommand = New ADODB.Command
    
    Set adoCommand.ActiveConnection = adoConnection
    
    With adoCommand
        .CommandType = adCmdStoredProc
        .CommandText = ProcName
        Set adoParameterExceptionNumber = .CreateParameter("p_NextNumber", adVarChar, adParamOutput, OutSize)
        .Parameters.Append adoParameterExceptionNumber
        .Prepared = True
        .Execute
    End With
    
    Dim strExceptionSequenceNumber As String
    strExceptionSequenceNumber = adoParameterExceptionNumber.Value
    
    ' close the connection
    If Not adoConnection Is Nothing Then
        adoConnection.Close
    End If
    
    Set xmlBatchExceptionNode = vxmlRequest.selectSingleNode("BATCHEXCEPTION")
    Set xmlSequenceAttribute = xmlBatchExceptionNode.ownerDocument.createAttribute("BATCHSEQNUMBER")
    
    If Not strExceptionSequenceNumber = "" Then
        xmlSequenceAttribute.Text = strExceptionSequenceNumber
    Else
        xmlSequenceAttribute.Text = ""
    End If
    
    xmlBatchExceptionNode.Attributes.setNamedItem xmlSequenceAttribute
    
    adoCreateFromNode xmlBatchExceptionNode, "BATCHEXCEPTION"
    
    gobjContext.SetComplete
    
CreateBatchExceptionExit:
    
    Set adoParameterExceptionNumber = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set xmlSequenceAttribute = Nothing
    Set xmlBatchExceptionNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
CreateBatchExceptionVbErr:
    
    gobjContext.SetAbort
    
    GoTo CreateBatchExceptionExit

End Sub

Private Sub IBatchScheduleDO_CreateBatchSchedule(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo CreateBatchScheduleExit
    
    'constant
    Const strFunctionName As String = "IBatchScheduleDO_CreateBatchSchedule"
    
    Dim xmlBatchSchedule As IXMLDOMNode
    
    Set xmlBatchSchedule = vxmlRequestNode.selectSingleNode("BATCHSCHEDULE")
    
    adoCreateFromNode xmlBatchSchedule, "BATCHSCHEDULE"
    
CreateBatchScheduleExit:
    
    Set xmlBatchSchedule = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

'BM0438
Private Sub IBatchScheduleDO_CreateBatchScheduleProgress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Const cstrFunctionName As String = "IBatchScheduleDO_CreateBatchScheduleProgress"
    Dim xmlBatchScheduleProgress As IXMLDOMNode
    
On Error GoTo CreateBatchScheduleProgressExit
    
    Set xmlBatchScheduleProgress = vxmlRequestNode.selectSingleNode("BATCHSCHEDULEPROGRESS")
    adoCreateFromNode xmlBatchScheduleProgress, "BATCHSCHEDULEPROGRESS"
    
CreateBatchScheduleProgressExit:
    gobjContext.SetComplete
    Set xmlBatchScheduleProgress = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BM0438 End

Private Sub IBatchScheduleDO_FindBatchAuditList(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Error Handler
    On Error GoTo FindBatchAuditListExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_FindBatchAuditList"
    
    'PSC 25/02/02 SYS4097 - Start
    Dim xmlBatchAudit As IXMLDOMNode
    
    Set xmlBatchAudit = xmlGetMandatoryNode(vxmlRequestParentNode, "BATCHAUDIT")
    adoGetAsXML xmlBatchAudit, vxmlResponseNode, "BATCHAUDIT"
    'PSC 25/02/02 SYS4097 - End
    
    'Checking the Record Existence
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
FindBatchAuditListExit:
    
    gobjContext.SetComplete
    
    Set xmlBatchAudit = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IBatchScheduleDO_GetBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    'Error Handler
    On Error GoTo GetBatchExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_GetBatch"
    
    'SYS2250 MDC 10/09/2001. Get the BATCH node if required
    Dim xmlBatchNode As IXMLDOMNode
    If vxmlRequestParentNode.nodeName = "BATCH" Then
        Set xmlBatchNode = vxmlRequestParentNode
    Else
        Set xmlBatchNode = xmlGetMandatoryNode(vxmlRequestParentNode, ".//BATCH")
    End If
    adoGetAsXML xmlBatchNode, vxmlResponseNode, "GETBATCH"
    Set xmlBatchNode = Nothing
    'Get Batch Details
'    adoGetAsXML vxmlRequestParentNode, vxmlResponseNode, "GETBATCH"
    'SYS2250 End
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    gobjContext.SetComplete
    
GetBatchExit:
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IBatchScheduleDO_GetBatchSchedule(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Varibles
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    
    'Error Handler
    On Error GoTo GetBatchScheduleExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_GetBatchSchedule"
        
    Dim xmlBatchScheduleNode As IXMLDOMNode
    
    Set xmlBatchScheduleNode = xmlGetMandatoryNode(vxmlRequestParentNode, "BATCHSCHEDULE")
    
    'Storing into the Local Variables
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchScheduleNode, "BATCHNUMBER")
    'BM0230 MDC 08/01/2003 - Enable retrieval of all runs for a given batch number
'    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchScheduleNode, "BATCHRUNNUMBER")
'    If IsNull(strBatchRunNumber) Or (Len(Trim$(strBatchRunNumber)) = 0) Then
    strBatchRunNumber = xmlGetAttributeText(xmlBatchScheduleNode, "BATCHRUNNUMBER")
    If strBatchRunNumber = "0" Then
    'BM0230 MDC 08/01/2003 - End
        
        adoGetAsXML vxmlRequestParentNode, vxmlResponseNode, "GETLATESTBATCHSCHEDULE"
    Else
        adoGetAsXML xmlBatchScheduleNode, vxmlResponseNode, "BATCHSCHEDULE"
    End If
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetBatchScheduleExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

'BM0438
Private Sub IBatchScheduleDO_GetBatchScheduleProgress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName   As String = "IBatchScheduleDO_GetBatchScheduleProgress"
    Dim xmlBatchScheduleNode As IXMLDOMNode
    
On Error GoTo GetBatchScheduleProgressExit
    
    Set xmlBatchScheduleNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHSCHEDULEPROGRESS")
    
    adoGetAsXML xmlBatchScheduleNode, vxmlResponseNode, "BATCHSCHEDULEPROGRESS"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
GetBatchScheduleProgressExit:
    gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BM0438 End

Private Sub IBatchScheduleDO_ProduceBatchScheduleReport( _
                            ByVal vxmlRequestParentNode As IXMLDOMNode, _
                            ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Varaibles
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlBatchNode As IXMLDOMNode
    
    Dim strBatchLogFilePath As String
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchLogFile  As String
    Dim intBatchLogFile  As Integer
    
    Dim intFileNo As Integer
    Dim strValue As String
    Dim intValue As Integer
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    'BMIDS622
    Dim blnWriteLog As Boolean
    'BMIDS622 End
    
    'Error Handler
    On Error GoTo ProduceBatchReportExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_ProduceBatchScheduleReport"
            
    'Store into Local Variables
    Set xmlBatchNode = vxmlRequestParentNode.selectSingleNode(".//BATCHSCHEDULE")
    
    strBatchNumber = xmlGetAttributeText(xmlBatchNode, "BATCHNUMBER")
    strBatchRunNumber = xmlGetAttributeText(xmlBatchNode, "BATCHRUNNUMBER")
    intBatchLogFile = GetGlobalParamAmount("BatchLogFile")
    
'Stage1
    
    'Get Batch Info
    
    'Preparing Request Element
    
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlElement = xmlTempDoc.createElement("BATCH")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    
    xmlTempElement.appendChild xmlElement
    
    'Call GetBatch
    IBatchScheduleDO_GetBatch xmlTempElement, vxmlResponseNode
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
'Stage2
    
    
    'Get BatchAudit Info
    xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
    
    'Call FindBatchAuditList
    
    'SYS3440 MDC 13/12/2001. Only select BatchAudit records for the current batch run
    
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    Set xmlElement = xmlTempDoc.createElement("BATCHAUDIT")
    xmlSetAttributeValue xmlElement, "BATCHNUMBER", strBatchNumber
    xmlSetAttributeValue xmlElement, "BATCHRUNNUMBER", strBatchRunNumber
    xmlTempElement.appendChild xmlElement
    'SYS3440 End.
    
    'PSC 25/02/02 SYS4097 - Stsrt
    On Error Resume Next
        
    IBatchScheduleDO_FindBatchAuditList xmlTempElement, vxmlResponseNode
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description

    On Error GoTo ProduceBatchReportExit
    
    If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    'PSC 25/02/02 SYS4097 - End
    
'Stage3
    
    If intBatchLogFile = 1 Then
        
        'Get the BatchLogFilePath GlobalParameter Value from GlobalParameter DB Table
        strBatchLogFilePath = GetGlobalParamString("BatchLogFilePath")
        
        'BMIDS622
        On Error Resume Next
        blnWriteLog = True
        'BMIDS622 End
        
        'Checking for the Directory Existence to write the outputFile
        'If does not Exist then Create One
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
        
        'BMIDS622 Restore orignal error handler
        On Error GoTo ProduceBatchReportExit
        If blnWriteLog Then
        'BMIDS622 End
        
            'strBatchLogFilePath = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber & ".xml"
            strBatchLogFile = strBatchLogFilePath & "\Batch_" & strBatchNumber & "_" & strBatchRunNumber
            
            'SYS3440 MDC 13/12/2001. Need to clone node otherwise Response from SetbatchStatus is empty
            xmlTempDoc.appendChild vxmlResponseNode.cloneNode(True)
            
            xmlTempDoc.Save (strBatchLogFile & ".xml")
        
            'SYS3440 MDC 13/12/2001. Create a user friendly Audit Report for each batch run.
    
            intFileNo = FreeFile
            Open strBatchLogFile & ".log" For Output As #intFileNo
            Print #intFileNo, "Batch Report"
            Print #intFileNo, "~~~~~~~~~~~~" & vbCrLf
            
            Print #intFileNo, "Batch Number: " & strBatchNumber
            Print #intFileNo, "Batch Run Number: " & strBatchRunNumber
            
            Set xmlNode = xmlGetNode(vxmlResponseNode, "BATCH")
            intValue = xmlGetAttributeAsInteger(xmlNode, "PROGRAMTYPE")
            strValue = GetComboText("BatchProgramType", intValue)
            Print #intFileNo, "Program Type: " & strValue
            
            strValue = xmlGetAttributeText(xmlNode, "DESCRIPTION")
            Print #intFileNo, "Description: " & strValue
            
            strValue = xmlGetAttributeText(xmlNode, "BATCHEXECUTIONDATETIME")
            Print #intFileNo, "Run Date/Time: " & strValue
            
            intValue = xmlGetAttributeAsInteger(xmlNode, "BATCHSTATUS")
            strValue = GetComboText("BatchStatus", intValue)
            Print #intFileNo, "Status: " & strValue & vbCrLf
                   
            strValue = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFRECORDS")
            Print #intFileNo, "Total Records Processed: " & strValue
            
            strValue = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFSUCCESSES")
            Print #intFileNo, "Total Successes: " & strValue
            
            strValue = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFFAILURES")
            Print #intFileNo, "Total Failures: " & strValue
            
            intValue = xmlGetAttributeAsInteger(xmlNode, "BATCHSCHEDULESTATUS")
            strValue = GetComboText("BatchScheduleStatus", intValue)
            Print #intFileNo, "Schedule Status: " & strValue
            
            Close #intFileNo
            'SYS3440 End
        End If 'BMIDS622
        
    End If
    
ProduceBatchReportExit:
    
    Set xmlNode = Nothing
    Set xmlTempDoc = Nothing
    Set xmlElement = Nothing
    Set xmlTempElement = Nothing
    Set xmlBatchNode = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IBatchScheduleDO_SetBatchStatus( _
                ByVal vxmlRequestParentNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Varaibles
    'BM0438 No longer used
    'Dim lngTotalNumberOfRecords As Long
    'Dim lngTotalNumberOfSuccess As Long
    'Dim lngTotalNumberOfFailures As Long
    'Dim intIndex As Integer
    'Dim cintValueIds As Collection
    'Dim xmlNode As IXMLDOMNode
    'BM0438 End
    
    Dim strBatchNumber As String
    Dim strBatchRunNumber As String
    Dim strBatchScheduleStatus As String
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlBatchSchedule As IXMLDOMElement
    Dim xmlTempElement  As IXMLDOMElement
    
    'Error Handler
    On Error GoTo SetBatchStatusExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_SetBatchStatus"
    
'Stage1:

    'Create TempResponseElement
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    'xmlTempDoc.appendChild xmlTempResponseElem
    
    'Store into Local Variables
    ' PSC 20/02/2002 SYS4097 - Start
    Set xmlBatchSchedule = xmlGetMandatoryNode(vxmlRequestParentNode, ".//BATCHSCHEDULE")
    strBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSchedule, "BATCHNUMBER")
    strBatchRunNumber = xmlGetMandatoryAttributeText(xmlBatchSchedule, "BATCHRUNNUMBER")
    strBatchScheduleStatus = xmlGetMandatoryAttributeText(xmlBatchSchedule, "STATUS")
    ' PSC 20/02/2002 SYS4097 - End
    
    'Preparing TempRequest Element
    Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
    Set xmlElement = xmlTempDoc.createElement("BATCHSCHEDULE")
    xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
    xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
        
    xmlTempElement.appendChild xmlElement
    
    'Call Get BatchSchedule for Batch details
    IBatchScheduleDO_GetBatchSchedule xmlTempElement, xmlTempResponseElem
    If xmlTempResponseElem.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
        
    'BM0438 No longer required
    'load the BatchSchedule Node into TempVariable
    'Set xmlNode = xmlTempResponseElem.selectSingleNode(".//BATCHSCHEDULE")
    
    'Store into Local Variables
    'lngTotalNumberOfRecords = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFRECORDS")
    'lngTotalNumberOfSuccess = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFSUCCESSES")
    'lngTotalNumberOfFailures = xmlGetAttributeText(xmlNode, "TOTALNUMBEROFFAILURES")
    
    'Set xmlTempDoc = Nothing
    'Set xmlTempResponseElem = Nothing
    'Set xmlNode = Nothing
    'BM0438 End
    
'Stage2:

    'Cheking If NoOfSuccess +  NoOfFailures = TotalNumberOfRecords then raise an Error else
    'Increment NoOfSuccess by 1
    
    'BM0438 No longer required
    'If lngTotalNumberOfSuccess + lngTotalNumberOfFailures <> lngTotalNumberOfRecords Then
    '    errThrowError strFunctionName, 305
    'End If
    'BM0438 End
       
'Stage3:

    'If totalNumberOfFailures is greater than 0 then
    'Create an element with "COMPLETED WITH ERRORS"  else with the Default status
    ' PSC 20/02/2002 SYS4097 - Start
'    Set cintValueIds = New Collection
'
'    GetValueIdsForValidationType "BatchScheduleStatus", "CWE", cintValueIds
'    For intIndex = 1 To cintValueIds.Count
'        strBatchScheduleStatus = cintValueIds.Item(intIndex)
'    Next
'
'    If lngTotalNumberOfFailures > 0 Then
'        xmlElement.setAttribute "STATUS", strBatchScheduleStatus
'    Else
'        xmlElement.setAttribute "STATUS", "30"
'    End If

    xmlElement.setAttribute "STATUS", strBatchScheduleStatus
    xmlCopyMandatoryAttribute xmlBatchSchedule, xmlElement, "STATUS"
    xmlCopyAttribute xmlBatchSchedule, xmlElement, "ERRORNUMBER"
    xmlCopyAttribute xmlBatchSchedule, xmlElement, "ERRORSOURCE"
    xmlCopyAttribute xmlBatchSchedule, xmlElement, "ERRORDESCRIPTION"
    'BMIDS622 Copy the CompletionDateTime attribute if it exists
    xmlCopyAttribute xmlBatchSchedule, xmlElement, "COMPLETIONDATETIME"
    'BMIDS622 End
    ' PSC 20/02/2002 SYS4097 - End
    
    'Create TempResponseElement
    Set xmlTempResponseElem = xmlTempDoc.createElement("RESPONSE")
    xmlTempResponseElem.setAttribute "TYPE", "SUCCESS"
    'xmlTempDoc.appendChild xmlTempResponseElem
    
    'call Update Batch Schedule with the new Status
    IBatchScheduleDO_UpdateBatchSchedule xmlElement, xmlTempResponseElem
    errCheckXMLResponse xmlTempResponseElem.xml, True
        
'Stage4:

    'If status is "COMPLETEWITHERRORS" or "COMPLETE" or "INITIALISATIONERROR" then
    'Produce batch Report
    ' PSC 20/02/2002 SYS4097
    If IsValidationType("BatchScheduleStatus", CInt(strBatchScheduleStatus), "CWE") Or _
       IsValidationType("BatchScheduleStatus", CInt(strBatchScheduleStatus), "C") Or _
       IsValidationType("BatchScheduleStatus", CInt(strBatchScheduleStatus), "IE") Then
        
        'Preparing TempRequestElement
        Set xmlTempElement = vxmlRequestParentNode.cloneNode(False)
        xmlTempElement.setAttribute "OPERATION", "PRODUCEBATCHSCHEDULEREPORT"
        
        Set xmlElement = xmlTempDoc.createElement("BATCHSCHEDULE")
        xmlElement.setAttribute "BATCHNUMBER", strBatchNumber
        xmlElement.setAttribute "BATCHRUNNUMBER", strBatchRunNumber
        
        xmlTempElement.appendChild xmlElement
        
        'Produce Batch Report with TempRequestElement
        IBatchScheduleDO_ProduceBatchScheduleReport xmlTempElement, vxmlResponseNode
        
    End If
    
SetBatchStatusExit:
    
    Set xmlTempDoc = Nothing
    Set xmlTempResponseElem = Nothing
    Set xmlElement = Nothing
    'Set xmlNode = Nothing
    'Set cintValueIds = Nothing
    Set xmlBatchSchedule = Nothing
    Set xmlTempElement = Nothing
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IBatchScheduleDO_UpdateBatch(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    'Declare Varaibles
    Dim strProgramType As String
    Dim xmlUpdatePayProcBatchParamsNode  As IXMLDOMNode
    
    'Error Handler
    On Error GoTo UpdateBatchExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_UpdateBatch"
    
    'Store ProgramType into a Local Variable
    strProgramType = xmlGetAttributeText(vxmlRequestParentNode, "PROGRAMTYPE")
    
    'Update Batch Table
    adoUpdateFromNode vxmlRequestParentNode, "BATCH"
    
    'If the Program type is PaymentProcessing then
    'Update "PAYPROCBATCHPARAMS" Db table
    If Len(strProgramType) > 0 Then
        If IsValidationType("BatchProgramType", CInt(strProgramType), "P") Then
            Set xmlUpdatePayProcBatchParamsNode = vxmlRequestParentNode.selectSingleNode(".//PAYPROCBATCHPARAMS")
            IBatchScheduleDO_UpdatePayProcBatchParams xmlUpdatePayProcBatchParamsNode, vxmlResponseNode
            errCheckXMLResponse vxmlResponseNode.xml, True
            Set xmlUpdatePayProcBatchParamsNode = Nothing
        End If
    End If
    
UpdateBatchExit:
    Set xmlUpdatePayProcBatchParamsNode = Nothing
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IBatchScheduleDO_UpdateBatchAudit(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateBatchAuditExit
    
    Const strFunctionName As String = "IBatchScheduleDO_UpdateBatchAudit"
    
    adoUpdateFromNode vxmlRequestParentNode, "BATCHAUDIT"
   
    
UpdateBatchAuditExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IBatchScheduleDO_UpdateBatchSchedule(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateBatchScheduleExit
    
    Const strFunctionName As String = "IBatchScheduleDO_UpdateBatchSchedule"
    
    adoUpdateFromNode vxmlRequestNode, "BATCHSCHEDULE"
    
UpdateBatchScheduleExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

'BM0438 GHun 02/04/2003
Private Sub IBatchScheduleDO_UpdateBatchScheduleProgress(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "IBatchScheduleDO_UpdateBatchScheduleProgress"
    
On Error GoTo UpdateBatchScheduleProgressExit
    
    adoUpdateFromNode vxmlRequestNode, "BATCHSCHEDULEPROGRESS"
    
UpdateBatchScheduleProgressExit:
    gobjContext.SetComplete
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BM0438 End

Private Function IBatchScheduleDO_UpdatePayProcBatchParams(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode) As Variant
    
    'Error Handler
    On Error GoTo UpdatePayProcBatchParamsExit
    
    'Constant
    Const strFunctionName As String = "IBatchScheduleDO_UpdatePayProcBatchParams"
    
    'Update PayProcBatchParams
    adoUpdateFromNode vxmlRequestParentNode, "PAYPROCBATCHPARAMS"
    
UpdatePayProcBatchParamsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
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

'BM0438
Private Sub IBatchScheduleDO_IncrementBatchScheduleFailure(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    IncrementBatchScheduleProgressTotal vxmlRequest, vxmlResponseNode, False
End Sub
'BM0438 End

'BM0438
Private Sub IBatchScheduleDO_IncrementBatchScheduleSuccess(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    IncrementBatchScheduleProgressTotal vxmlRequest, vxmlResponseNode, True
End Sub
'BM0438 End

'BM0438 GHun 27/03/2003
Private Sub IncrementBatchScheduleProgressTotal(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode, ByVal blnIsSuccess As Boolean)
    Const cstrFunctionName  As String = "IncrementBatchScheduleProgressTotal"
    Dim sBatchNumber        As String
    Dim lBatchRunNumber     As Long
    Dim sBatchAuditGUID     As String
    Dim lErrorNumber        As Long
    Dim sErrorSource        As String
    Dim sErrorDescription   As String
    Dim adoConnection       As ADODB.Connection
    Dim adoCommand          As ADODB.Command
    Dim adoReturnParam      As ADODB.Parameter
    Dim xmlBatchSchedule    As IXMLDOMElement

On Error GoTo IncrementBatchScheduleProgressTotalExit
    
    Set xmlBatchSchedule = xmlGetMandatoryNode(vxmlRequest, "BATCHSCHEDULE")
    sBatchNumber = xmlGetMandatoryAttributeText(xmlBatchSchedule, "BATCHNUMBER")
    lBatchRunNumber = xmlGetMandatoryAttributeAsLong(xmlBatchSchedule, "BATCHRUNNUMBER")
    sBatchAuditGUID = xmlGetMandatoryAttributeText(xmlBatchSchedule, "BATCHAUDITGUID")
    If Not blnIsSuccess Then
        lErrorNumber = xmlGetAttributeAsLong(xmlBatchSchedule, "ERRORNUMBER")
        sErrorSource = xmlGetAttributeText(xmlBatchSchedule, "ERRORSOURCE")
        sErrorDescription = xmlGetAttributeText(xmlBatchSchedule, "ERRORDESCRIPTION")
    End If

    Set adoConnection = New ADODB.Connection
    adoConnection.ConnectionString = adoGetDbConnectString
    adoConnection.Open
        
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = adoConnection
    
    'Create the parameters and call the appropriate stored proc
    With adoCommand
        .CommandType = adCmdStoredProc
        Set adoReturnParam = .CreateParameter("@ReturnValue", adInteger, adParamReturnValue)
        .Parameters.Append adoReturnParam
        .Parameters.Append .CreateParameter("@p_BatchNumber", adBSTR, adParamInput, 10, sBatchNumber)
        .Parameters.Append .CreateParameter("@p_BatchRunNumber", adInteger, adParamInput, , lBatchRunNumber)
        .Parameters.Append .CreateParameter("@p_BatchAuditGUID", adBinary, adParamInput, 16, adoAssistEx.GuidStringToByteArray(sBatchAuditGUID))
        If blnIsSuccess Then
            .CommandText = "USP_INCREMENTBATCHSCHEDULESUCCESS"
        Else
            .CommandText = "USP_INCREMENTBATCHSCHEDULEFAILURE"
            .Parameters.Append .CreateParameter("@p_ErrorNumber", adInteger, adParamInput, , lErrorNumber)
            .Parameters.Append .CreateParameter("@p_ErrorSource", adBSTR, adParamInput, 500, sErrorSource)
            .Parameters.Append .CreateParameter("@p_ErrorDescription", adBSTR, adParamInput, 255, sErrorDescription)
        End If
        .Execute , , adExecuteNoRecords
    End With
    
    'Close the database connection
    adoConnection.Close
    Set adoCommand.ActiveConnection = Nothing
    
    'The return value holds the LastTransaction indicator
    Set xmlBatchSchedule = vxmlRequest.ownerDocument.createElement("BATCHSCHEDULE")
    If adoReturnParam.Value = 1 Then
        xmlBatchSchedule.setAttribute "LASTTRANSACTION", 1
    Else
        xmlBatchSchedule.setAttribute "LASTTRANSACTION", 0
    End If

    vxmlResponseNode.appendChild xmlBatchSchedule

IncrementBatchScheduleProgressTotalExit:
    gobjContext.SetComplete
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set adoReturnParam = Nothing
    Set xmlBatchSchedule = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BM0438 End

