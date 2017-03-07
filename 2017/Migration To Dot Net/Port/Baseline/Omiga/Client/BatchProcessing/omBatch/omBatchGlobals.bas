Attribute VB_Name = "omBatchGlobals"
'Workfile:      omBatchGlobals.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Batch Processing General module.

'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MC     05/12/2001  SYS3018 - Fixes to Batch Process for SQL Server.
'PSC    25/02/2002  SYS4097 - Process Initialisation Errors
'-------------------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    08/03/2005  MAR1355 - Amend CallBatchProcessAsync to set up correct request
'-------------------------------------------------------------------------------------------------------

Option Explicit

Public Enum TMERROR
    ' (msgTm) Task Management specific errors
    oeTmNoNextStage = 4800
    oeTmStageNotApplicable = 4801
    oeTmMandatoryTasksOutstanding = 4802
    oeTmNotExceptionStage = 4803
    oeTmNoStageDetail = 4804
    oeTmNoTaskDetail = 4805
    oeTmNoStageAuthority = 4806
    oeTmNoTaskAuthority = 4807
End Enum

'SYS3018 MDC 04/12/2001 - API Declarations & Variables
Public Declare Function SetTimer Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
    
Public Declare Function KillTimer Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
    
Public g_strRequest As String
Public g_strProgId As String
Public g_strMethodName As String
Public g_lngTimerId As Long
'SYS3018 End

Public Sub Main()
    ' adoAssist
    adoLoadSchema
    adoBuildDbConnectionString
End Sub

'SYS3018 MDC 04/12/2001 - New function
Public Sub CallBatchProcessAsync(ByVal hwnd As Long, ByVal uMsg As Long, _
                                    ByVal idEvent As Long, ByVal dwTime As Long)

On Error GoTo CallBatchProcessAsyncErr

Dim objBatchProgramObjectCall As Object
Dim objBatchBO As BatchScheduleBO

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequest As IXMLDOMNode
Dim xmlBatchNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlBatchScheduleNode As IXMLDOMNode

Dim strResponse As String
Dim strBatchNumber As String
Dim strStatus As String
Dim blnInitialisation As Boolean

    ' PSC 25/02/02 SYS4097
    blnInitialisation = True

    'Stop the timer so that it only fires once
    KillTimer 0, g_lngTimerId

    'Call the required component
    Set objBatchProgramObjectCall = CreateObject(g_strProgId)
    strResponse = CallByName(objBatchProgramObjectCall, g_strMethodName, VbMethod, g_strRequest)
    errCheckXMLResponse strResponse, True
    
    ' PSC 25/02/02 SYS4097
    blnInitialisation = False

    'Get the batch Number
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.loadXML g_strRequest
    Set xmlBatchNode = xmlGetMandatoryNode(xmlDoc.documentElement, ".//BATCH")
    strBatchNumber = xmlGetAttributeText(xmlBatchNode, "BATCHNUMBER")
    
    'Create a Request to call CompleteRunBatch
    ' PSC 08/03/2005 MAR1355
    Set xmlRequest = xmlGetRequestNode(xmlDoc.documentElement)
    xmlSetAttributeValue xmlRequest, "OPERATION", "CompleteRunBatch"
    Set xmlBatchNode = xmlDoc.createElement("BATCH")
    xmlSetAttributeValue xmlBatchNode, "BATCHNUMBER", strBatchNumber
    xmlRequest.appendChild xmlBatchNode
    
    'Call CompleteRunBatch
    Set objBatchBO = CreateObject(App.Title & ".BatchScheduleBO")
    strResponse = objBatchBO.omBatchRequest(xmlRequest.xml)
    errCheckXMLResponse strResponse, True
    
CallBatchProcessAsyncExit:
    Set objBatchProgramObjectCall = Nothing
    Set objBatchBO = Nothing
    Set xmlRequest = Nothing
    Set xmlDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlBatchNode = Nothing
    Set xmlBatchScheduleNode = Nothing
    Exit Sub
    
CallBatchProcessAsyncErr:
    App.LogEvent vbCr & "Error: " & Err.Number & vbCr & "Source: " & Err.Source _
                    & vbCr & "Description: " & Err.Description, vbLogEventTypeError
    
    ' PSC 25/02/02 SYS4097 - Start
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.loadXML g_strRequest
    Set xmlTempRequest = xmlDoc.documentElement.cloneNode(False)

    Set xmlBatchNode = xmlGetMandatoryNode(xmlDoc.documentElement, ".//BATCH")
    
    Set xmlBatchScheduleNode = xmlDoc.createElement("BATCHSCHEDULE")
    xmlTempRequest.appendChild xmlBatchScheduleNode
    
    xmlCopyAttribute xmlBatchNode, xmlBatchScheduleNode, "BATCHNUMBER"
    xmlCopyAttribute xmlBatchNode, xmlBatchScheduleNode, "BATCHRUNNUMBER"
    
    xmlSetAttributeValue xmlBatchScheduleNode, "ERRORNUMBER", CStr(Err.Number)
    xmlSetAttributeValue xmlBatchScheduleNode, "ERRORSOURCE", CStr(Err.Source)
    xmlSetAttributeValue xmlBatchScheduleNode, "ERRORDESCRIPTION", CStr(Err.Description)
    
    ' Set status to initialisation error if initialisation else just update
    ' the error info
    If blnInitialisation = True Then
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "SetBatchStatus"
        strStatus = GetFirstComboValueId("BatchScheduleStatus", "IE")
        xmlSetAttributeValue xmlBatchScheduleNode, "STATUS", strStatus
    Else
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "UpdateBatchSchedule"
    End If
    
    Set objBatchBO = CreateObject(App.Title & ".BatchScheduleBO")

    strResponse = objBatchBO.omBatchRequest(xmlTempRequest.xml)
    ' PSC 25/02/02 SYS4097 - End

    GoTo CallBatchProcessAsyncExit
    
End Sub
'SYS3018 End

