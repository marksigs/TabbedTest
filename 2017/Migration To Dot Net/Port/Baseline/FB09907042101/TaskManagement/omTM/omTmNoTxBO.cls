VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "omTmNoTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OmTmNoTxBO.cls
'Copyright:     Copyright © 2003 Marlborough Stirling

'Description:   Task Manager Business Object (No Transactions). BMids specific.
'
'-------------------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    22/04/2003  BM0340 Created
'GHun   13/06/2003  BM0340 Amended MoveToNextStage & MoveToStage to use Stage details (instead of CaseStage)
'                   for creating ApplicationStage
'GD     09/07/2003  BM0340 Amended AddDefaultValuesToCaseTask tlook at vxmlTaskNode, not vxmlCaseTaskNode for Task Owner info.
'GHun   27/11/2003  BMIDS658 Added FindCustomersForCreditCheck and GetBureauRecords
'GHun   12/12/2003  BMIDS658 Amended FindCustomersForCreditCheck to accept element based request
'GHun   14/01/2004  BMIDS649 Amended MoveToStage and MoveToNextStage to call ProcessAutomaticTasks
'GHun   10/02/2004  BMIDS681 CC053 Move to next stage without user intervention. Amended
'                   MoveToStage & MoveToNextStage and added HasManualStageTasks & IsMovableStage
'HMA    10/02/2004  BMISD691 Pass EXCEPTIONREASON to OmTMRules in RunUserRules
'DRC    03/03/2004  BMIDS680 When processing Automatic tasks, if one fails then stop processing &
'                   Put automatic tasks in order of their TASKIDS
'GHun   04/03/2004  BMIDS697 Moved DeclineApplication from OmTmBO to here
'GHun   08/03/2004  BMIDS681 Changed HasManualStageTasks to HasIncompleteTasks so failed automatic tasks also return true.
'GHun   11/03/2004  BMIDS681 Amended ProcessAutomaticTasks to use cloned Printer and Application nodes
'DRC    15/03/04    BMIDS681 Must pass the Printer node From MovetoStage to MovetoNextStage for Automatic tasks
'DRC    17/03/04    BMIDS681 Pass the Printer node in Process Auto tasks only when it exists
'GHun   17/03/2004  BMIDS729 Amend DeclineApplication to pass printer node through
'GHun   20/04/2004  BMIDS754 Amended MoveToNextStage to refresh RB XML before calling rules
'JD     04/08/2004  BMIDS832 return CustomerOrder in FindCustomersForCreditCheck
'------------------------------------------------------------------------------------------------------
'MARS history:
'
'GHun   01/08/2005  MAR7 Applied CORE174 changes (moved AddDefaultValuesToCaseTask & SetCaseTaskDueDateAndTime to omTMFunctions.bas)
'PSC    15/09/2005  MAR32 Task Automation Changes
'MV     03/10/2005  MAR35   Amended OmTmNoTxRequest() and Added CreateLegalRepTask()
'MV     07/10/2005  MAR128  Amended CreateLegalRepTask() Core Review Fixes
'GHun   22/10/2005  MAR231  Changed CreateLegalRepTask
'GHun   01/11/2005  MAR144  Changed CreateLegalRepTask to rename global parameter
'GHun   01/11/2005  MAR357  Changed CreateLegalRepTask
'BC     13/12/2005  MAR842  Create call to GetApplicationData in omTMNoTXRequest
'HMA    18/01/2006  MAR1048 Added CreateFundsReleaseFailedTask
'PSC    20/01/2006  MAR1073 Amend MoveToStage to only append printer node if present
'PSC    24/01/2006  MAR1118 Add logging
'HMA    25/01/2006  MAR1110 Amend CreateFundsReleaseFailedTask to create task note
'GHun   21/02/2006  MAR1198 Changed HasIncompleteTasks to ignore CarryForward tasks
'DRC    02/03/2006  MAR1334 Cancel Disbursements when moving to Decline/Cancel Stage
'HMA    22/03/2006  MAR1422 Changed CreateLegalRepTask and CreateFundsReleaseFailedTask to set task due date.
'JD     30/03/2006  MAR1476 ProcessCarryForwardTasks - copy over the casetaskname for customer tasks.
'GHun   18/04/2006  MAR1622 Changed CreateFundsReleaseFailedTask
'HMA    24/04/2006  MAR1408 Changes to MoveToStage to set TaskDueDateAndTime for Post Completion tasks.
'HMA    03/05/2006  MAR1689 Changes for Post Completion task dates.
'PSC    15/05/2006  MAR1784 Amend to call CopyTaskDataFromStage only if there at least one CASESTAGE node returned
'                   from the rules. Also change ProcessAutomaticTasks and MoveToNextStage to use a copy of the
'                   request rather than the original
'GHun   02/06/2006  MAR1848 Added GetAndSynchroniseCustomerDetails And GetCustomerNumbers wrappers
'AW     26/09/06    EP1155  Set up attributes required for retrieving archived task notes
'IK     20/01/2007  EP2_902 AddDefaultValuesToCaseStage moved to omTmFunctions.bas
'                           modify CopyTaskDataFromStage, copy REMOTEOWNERTASKIND
'------------------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext
Private gobjMsgTmBO As Object
'Private Const cstrTaskComplete = 40

Public Function OmTmNoTxRequest(ByVal vstrXmlIn As String) As String

On Error GoTo OmTmNoTxRequestVbErr

Dim xmlInDoc As FreeThreadedDOMDocument40
Dim xmlOutDoc As FreeThreadedDOMDocument40
Dim xmlElem As IXMLDOMElement
Dim xmlNode As IXMLDOMElement
Dim xmlRequestNode As IXMLDOMNode
Dim xmlResponseElem As IXMLDOMElement
Dim xmlResponseNode As IXMLDOMNode

Dim strOperation As String

Const cstrFunctionName As String = "OmTmNoTxRequest"

    Set gobjMsgTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    If gobjMsgTmBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, "MsgTm.MsgTmBO"
    End If
    
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
    xmlCheckMandatoryAttribute xmlRequestNode, "OPERATION"
    
    strOperation = xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    strOperation = UCase$(strOperation)
    
    Select Case strOperation
            
        Case "PROCESSAUTOMATICTASKS"
            ProcessAutomaticTasks xmlRequestNode
        
        Case "MOVETONEXTSTAGE"
            MoveToNextStage xmlRequestNode, xmlResponseNode
        
        Case "MOVETOSTAGE"
            MoveToStage xmlRequestNode, xmlResponseNode
        
        'BMIDS658
        Case "FINDCUSTOMERSFORCREDITCHECK"
            FindCustomersForCreditCheck xmlRequestNode, xmlResponseNode
            
        Case "GETBUREAURECORDS"
            GetBureauRecords xmlRequestNode, xmlResponseNode
        'BMIDS658 End
        
        'BMIDS697 GHun 04/03/2004 Move here from omTmBO
        Case "DECLINEAPPLICATION"
            DeclineApplication xmlRequestNode
        'BMIDS697 End
        ' PSC 15/09/2005 MAR32 - Start
        Case "GETADDITIONALCASETASKDATA"
            GetAdditionalCaseTaskData xmlRequestNode, xmlResponseNode
        Case "GETCURRENTSTAGE"
            GetCurrentStage xmlRequestNode, xmlOutDoc
        ' PSC 15/09/2005 MAR32 - End
        Case "CREATELEGALREPTASK"
            CreateLegalRepTask xmlRequestNode
        ' BC 13/12/2005 MAR842 - Begin
        Case "GETAPPLICATIONDATA"
            Set xmlResponseNode = GetApplicationData(xmlRequestNode, "CDRBTemplate")
            xmlResponseElem.appendChild xmlResponseNode
        ' BC 13/12/2005 MAR842 - End
        'MAR1048
        Case "CREATEFUNDSRELEASEFAILEDTASK"
            CreateFundsReleaseFailedTask xmlRequestNode
        'MAR1848 GHun
        Case "GETANDSYNCHRONISECUSTOMERDETAILS"
            GetAndSynchroniseCustomerDetails xmlRequestNode, xmlResponseNode
        Case "GETCUSTOMERNUMBERS"
            GetCustomerNumbers xmlRequestNode, xmlResponseNode
        'MAR1848 End
        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, "OPERATION = " & strOperation
    End Select
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    OmTmNoTxRequest = xmlOutDoc.xml
    
OmTmNoTxRequestExit:
    Set gobjMsgTmBO = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmTmNoTxRequestVbErr:
    OmTmNoTxRequest = errCreateErrorResponse()
    GoTo OmTmNoTxRequestExit
    
End Function

Private Sub CreateLegalRepTask(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo CreateLegalRepTaskExit
    
    Const cstrFunctionName As String = "CreateLegalRepTask"
    
    Dim strLegalRepInactiveTaskID As String
    Dim xmlFindCaseTaskList As IXMLDOMNode
    Dim xmlThisRequestDoc  As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTMCaseTaskNode  As IXMLDOMNode
    Dim xmlTMNewCaseTaskNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlAttribute As IXMLDOMAttribute
    Dim xmlAppNode As IXMLDOMNode
'Step:

    'Check to see if task needs to be created
        
    Set xmlTMCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "//REQUEST/CASETASK")
    Set xmlAppNode = xmlGetMandatoryNode(vxmlRequestNode, "//REQUEST/APPLICATION")

'Step:
    
    'MAR231 GHun
    'MAR144 GHun rename global parameter and check if it is missing
    strLegalRepInactiveTaskID = GetGlobalParamString("TMLegalRepInactiveTaskID")
    If Len(strLegalRepInactiveTaskID) = 0 Then
        errThrowError cstrFunctionName, oeMissingParameter, "TMLegalRepInactiveTaskID"
    End If
    'MAR144 End
    
    Dim strCondition As String
    strCondition = "CASEID = '" & xmlGetMandatoryAttributeText(xmlTMCaseTaskNode, "CASEID") & "'"
    strCondition = strCondition & " AND ACTIVITYID = " & xmlGetMandatoryAttributeText(xmlTMCaseTaskNode, "ACTIVITYID")
    strCondition = strCondition & " AND ACTIVITYINSTANCE =" & xmlGetMandatoryAttributeText(xmlTMCaseTaskNode, "ACTIVITYINSTANCE")
    strCondition = strCondition & " AND TASKID ='" & strLegalRepInactiveTaskID & "'"
    strCondition = strCondition & " AND TASKSTATUS ='10'"   'MAR231 GHun Check for inactive status instead of pending

    If Not adoCheckRecordExists("TM_CASETASKDETAIL", strCondition) Then
                 
'Step:

        'Get the task details
        Dim xmlGetTaskDetailDoc  As FreeThreadedDOMDocument40
        Dim vxmlResponseDoc As FreeThreadedDOMDocument40
        Dim xmlNode As IXMLDOMNode
        Dim xmlElem As IXMLDOMNode
        
        Set xmlGetTaskDetailDoc = New FreeThreadedDOMDocument40
        Set vxmlResponseDoc = New FreeThreadedDOMDocument40
        
        xmlGetTaskDetailDoc.validateOnParse = False
        xmlGetTaskDetailDoc.setProperty "NewParser", True
        xmlGetTaskDetailDoc.async = False
        
        'Set operation attribute
        Set xmlNode = xmlGetRequestNode(vxmlRequestNode)
        xmlNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskDetail"
    
        Set xmlElem = xmlGetTaskDetailDoc.createElement("TASK")
        xmlSetAttributeValue xmlElem, "TASKID", strLegalRepInactiveTaskID
        
        xmlNode.appendChild xmlElem
        xmlGetTaskDetailDoc.appendChild xmlNode
        
        'Run GetTaskDetail request
        RunTmRequest xmlGetTaskDetailDoc, vxmlResponseDoc
        
        Dim strTaskName As String
        strTaskName = xmlGetAttributeText(vxmlResponseDoc.selectSingleNode("RESPONSE/TASK"), "TASKNAME")
        
'Step:

        'Create Adhoc Case Task
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATECASETASK"
        
        Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
    
        'MAR357 GHun only copy required attributes
        'For Each xmlAttribute In xmlTMCaseTaskNode.Attributes
        '    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, xmlAttribute.nodeName
        'Next
        
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "SOURCEAPPLICATION", xmlGetAttributeText(xmlTMCaseTaskNode, "SOURCEAPPLICATION")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "CASEID", xmlGetAttributeText(xmlTMCaseTaskNode, "CASEID")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "ACTIVITYID", xmlGetAttributeText(xmlTMCaseTaskNode, "ACTIVITYID")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE", xmlGetAttributeText(xmlTMCaseTaskNode, "ACTIVITYINSTANCE")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "CASEACTIVITYGUID", xmlGetAttributeText(xmlTMCaseTaskNode, "CASEACTIVITYGUID")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "STAGEID", xmlGetAttributeText(xmlTMCaseTaskNode, "STAGEID")
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO", xmlGetAttributeText(xmlTMCaseTaskNode, "CASESTAGESEQUENCENO")
        'MAR357 End
        
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strLegalRepInactiveTaskID
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKNAME", strTaskName
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "CASETASKNAME", strTaskName

        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "1"

        'MAR1422 Set up TaskDueDateAndTime
        xmlSetSysDateToNodeAttrib xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME"

        xmlNewRequest.appendChild xmlTMNewCaseTaskNode
        xmlNewRequest.appendChild xmlAppNode
        
        xmlThisRequestDoc.appendChild xmlNewRequest

        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True

        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
         
    End If

CreateLegalRepTaskExit:
    
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlTMCaseTaskNode = Nothing
    Set xmlTMNewCaseTaskNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlAttribute = Nothing
    Set xmlAppNode = Nothing
    
    errCheckError cstrFunctionName
    
End Sub
Private Sub ProcessAutomaticTasks(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo ProcessAutomaticTasksExit


Dim objNTxBO As IOmTmNTxBO

Dim xmlAutoRequestDoc As FreeThreadedDOMDocument40
Dim xmlAutoRequestNode As IXMLDOMNode
Dim xmlAutoCaseTaskNode As IXMLDOMNode
Dim xmlOrderedAutoCaseTaskNode  As IXMLDOMNode
Dim xmlCurrentStageNode As IXMLDOMNode
Dim xmlApplicationNode As IXMLDOMNode
Dim xmlPrinterNode As IXMLDOMNode
Dim xmlCaseActivityNode As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTempNode As IXMLDOMNode
Dim xmlAutoTaskList As IXMLDOMNode
Dim xslDoc As FreeThreadedDOMDocument40
Dim xmlTransformedDoc As FreeThreadedDOMDocument40

'Dim sTempTaskName As String
'Dim iAutoErrorCount As Integer
Dim strResponse As String
Dim strXslPath As String
Dim blnLoadedOk As Boolean

Dim xmlTempRequest As IXMLDOMNode   ' PSC 19/05/2006 MAR1784

Const cstrFunctionName As String = "ProcessAutomaticTasks"
        
    '==================================================================================================
    'Initialise
    '==================================================================================================

    'Validate Request
    '   CURRENTSTAGE for 'MoveToNextStage'
    '   CASESTAGE for 'MoveToStage'
    '   CASEACTIVITY for 'ReinstatePreviousStage'
    
    ' PSC 19/05/2006 MAR1784 - Start
    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)

    Set xmlCaseActivityNode = xmlGetNode(xmlTempRequest, "CASEACTIVITY")
        
    If xmlCaseActivityNode Is Nothing Then
        Set xmlTempNode = xmlGetNode(xmlTempRequest, "CURRENTSTAGE")
        If xmlTempNode Is Nothing Then
            Set xmlTempNode = xmlGetMandatoryNode(xmlTempRequest, "CASESTAGE")
            xmlChangeNodeName xmlTempNode, "CASESTAGE", "CASEACTIVITY"
        Else
            xmlChangeNodeName xmlTempNode, "CURRENTSTAGE", "CASEACTIVITY"
        End If
        Set xmlCaseActivityNode = xmlTempNode
    End If
    ' PSC 19/05/2006 MAR1784 - End
    
    Dim blnDebug As Boolean
    Dim strCaseId As String
    blnDebug = IsDebugEnabled()
    strCaseId = xmlGetAttributeText(xmlCaseActivityNode, "CASEID")
    
    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Start", vbLogEventTypeInformation
    End If
    
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "ACTIVITYINSTANCE"
    
    ' PSC 19/05/2006 MAR1784 - Start
    Set xmlApplicationNode = xmlTempRequest.selectSingleNode("./APPLICATION")
    Set xmlPrinterNode = xmlTempRequest.selectSingleNode("./PRINTER")
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetCurrentStage"
    ' PSC 19/05/2006 MAR1784 - End
    
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    
    '==================================================================================================
    'Get the current stage
    '==================================================================================================
    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Getting Current Stage", vbLogEventTypeInformation
    End If
    
    strResponse = gobjMsgTmBO.TmRequest(xmlTempRequest.xml)     ' PSC 19/05/2006 MAR1784
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    Set xmlCurrentStageNode = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "CASESTAGE")
    '
    'BMIDS 680 DRC Start Make a list of Automatic tasks sorted by TASKID
    '
    Set xslDoc = New FreeThreadedDOMDocument40
    xslDoc.validateOnParse = False
    xslDoc.setProperty "NewParser", True
    xslDoc.async = False

    strXslPath = App.Path & "\ProcessAutomaticTasks.xslt"
    strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
    blnLoadedOk = xslDoc.Load(strXslPath)
    
    If Not blnLoadedOk Then
        errThrowError cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
    End If
    
    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Getting Automatic Tasks via xslt", vbLogEventTypeInformation
    End If
    
    Set xmlTransformedDoc = New FreeThreadedDOMDocument40
    xmlTransformedDoc.validateOnParse = False
    xmlTransformedDoc.setProperty "NewParser", True
    xmlTransformedDoc.async = False
    Set xmlAutoTaskList = xmlResponseDoc.documentElement
    xmlAutoTaskList.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
    Set xmlAutoTaskList = xmlTransformedDoc.selectSingleNode("CASETASKLIST")
    
    '==================================================================================================
    'Process each automatic task (in its own transaction)
    '==================================================================================================
    'Debug.Print xmlTransformedDoc.xml
    ' Get nodes on re-ordered list
    'Set xmlNodeList = xmlCurrentStageNode.selectNodes("CASETASK[@AUTOMATICTASKIND='1']")
    Set xmlNodeList = xmlAutoTaskList.selectNodes("CASETASK")
    'sTempTaskName = vbNullString

    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- " & CStr(xmlNodeList.length) & "  Automatic Tasks found", vbLogEventTypeInformation
    End If

' Raise error when it occurs and exit For loop
' On Error Resume Next
    
    For Each xmlOrderedAutoCaseTaskNode In xmlNodeList

'        ' Get the full task details keyed by the TASKID and TASKINSTANCE
'        Set xmlAutoCaseTaskNode = xmlCurrentStageNode.selectSingleNode("CASETASK[@TASKID='" & _
'          xmlAssistEx.xmlGetAttributeText(xmlOrderedAutoCaseTaskNode, "TASKID") & "' and " & _
'          "@TASKINSTANCE='" & _
'          xmlAssistEx.xmlGetAttributeText(xmlOrderedAutoCaseTaskNode, "TASKINSTANCE") & "']")
  'DRC BMIDS680 End
        Set xmlAutoRequestDoc = New FreeThreadedDOMDocument40
        xmlAutoRequestDoc.validateOnParse = False
        xmlAutoRequestDoc.setProperty "NewParser", True
        xmlAutoRequestDoc.async = False
        Set xmlAutoRequestNode = xmlAutoRequestDoc.appendChild(xmlTempRequest.cloneNode(False))  ' PSC 19/05/2006 MAR1784
        xmlAutoRequestNode.Attributes.getNamedItem("OPERATION").Text = "ProcessTask"
       
        xmlAutoRequestNode.appendChild xmlOrderedAutoCaseTaskNode
        xmlSetAttributeValue xmlOrderedAutoCaseTaskNode, "COMPLETETASKONPRINT", "1"

        'BMIDS681 GHun 11/03/2004 The nodes should be cloned so they will be available for
        'all iterations of the loop, not just the first one
        xmlAutoRequestNode.appendChild xmlApplicationNode.cloneNode(True)
        If Not (xmlPrinterNode Is Nothing) Then
            xmlAutoRequestNode.appendChild xmlPrinterNode.cloneNode(True)
        End If
        'BMIDS681 End
        On Error GoTo AutoTaskAbort
        
        If blnDebug Then
            App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Processing Task: " & xmlGetAttributeText(xmlOrderedAutoCaseTaskNode, "TASKID"), vbLogEventTypeInformation
        End If

        objNTxBO.ProcessTask xmlAutoRequestNode
        On Error GoTo ProcessAutomaticTasksExit
        
        If blnDebug Then
            App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Processed Task: " & xmlGetAttributeText(xmlOrderedAutoCaseTaskNode, "TASKID"), vbLogEventTypeInformation
        End If
        
        'BMIDS 680 - Suppress Error reporting
'        If Err.Number <> 0 Then
'            Err.Clear

'            iAutoErrorCount = iAutoErrorCount + 1
'            sTempTaskName = sTempTaskName & xmlAutoCaseTaskNode.Attributes.getNamedItem("TASKID").xml
'        End If
    Next
    
   
'    If iAutoErrorCount > 0 Then
'        errThrowError cstrFunctionName, oeTmAutomaticTaskDidntComplete, "", iAutoErrorCount, sTempTaskName
'    End If

ProcessAutomaticTasksExit:
    Set objNTxBO = Nothing
    Set xmlAutoRequestDoc = Nothing
    Set xmlAutoRequestNode = Nothing
    Set xmlAutoCaseTaskNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlPrinterNode = Nothing
    Set xmlCaseActivityNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlNodeList = Nothing
    
    Set xmlTempRequest = Nothing        ' PSC 19/05/2006 MAR1784

    
    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- End", vbLogEventTypeInformation
    End If

    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
AutoTaskAbort:
    'BMIDS680 suppress error reporting
    If blnDebug Then
        App.LogEvent "ProcessAutomaticTasks (" & strCaseId & "):- Failed to process task: " & xmlGetAttributeText(xmlOrderedAutoCaseTaskNode, "TASKID") & " Error: " & Err.Description & " Source: " & Err.Source, vbLogEventTypeInformation
    End If
    
    Err.Clear
    Set objNTxBO = Nothing
    Set xmlAutoRequestDoc = Nothing
    Set xmlAutoRequestNode = Nothing
    Set xmlAutoCaseTaskNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlPrinterNode = Nothing
    Set xmlCaseActivityNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlNodeList = Nothing
End Sub

'=======================================================================================================
' BM0340 Functions copied from omTmBO to support MoveToNextStage
'=======================================================================================================

Private Sub MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode, _
                             ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo MoveToStageExit

    Const cstrFunctionName As String = "MoveToStage"

    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextCaseStageDoc As FreeThreadedDOMDocument40
    Dim xmlCreateCaseStageDoc As FreeThreadedDOMDocument40
    Dim xmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40
    'Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
    Dim xmlRbResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlNextCaseStageNode As IXMLDOMNode
    Dim xmlRequestCaseStageNode As IXMLDOMNode
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim xmlComboNode As IXMLDOMNode

    Dim objNTxBO As IOmTmNTxBO
    Dim xmlValidateDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlApplicationRequestDoc As FreeThreadedDOMDocument40

    'BMIDS681 GHun 10/02/2004
    Dim strStageId      As String
    Dim strActivityId   As String
    Dim blnProceed      As Boolean
    Dim xmlRequest      As IXMLDOMNode
    'BMIDS681 End
    
    Dim xmlSetSLADateDoc As FreeThreadedDOMDocument40   ' PSC 22/09/2005 MAR32
    
    'DRC 01/03/2006 MAR1334
    Dim xmlPayProcDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim objPayProc As Object
    Dim lngErrorNumber As Long
    'MAR1334 - End
    
    'MAR1408
    Dim strNewStageId As String
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlROTNode As IXMLDOMNode
    Dim objROTBO As Object
    Dim xmlROTResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempNode As IXMLDOMNode
    Dim strCompletionDate As String
    Dim colValueIDs As Collection
    Dim varValueId As Variant
    Dim strTaskType As String
    Dim bCompletion As Boolean
    Dim strDueDateTime As String
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlROTResponseDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequestCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    
    Dim blnDebug As Boolean
    Dim strCaseId As String
    blnDebug = IsDebugEnabled()
    strCaseId = xmlGetAttributeText(xmlRequestCaseStageNode, "CASEID")

    ' validate request
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Start", vbLogEventTypeInformation
        App.LogEvent "MoveToStage (" & strCaseId & "):- Request: " & vxmlRequestNode.xml, vbLogEventTypeInformation
        App.LogEvent "MoveToStage (" & strCaseId & "):- Validating Request", vbLogEventTypeInformation
    End If
    
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "CASEID"
    'BMIDS681 Get the ActivityId
    strActivityId = xmlGetMandatoryAttributeText(xmlRequestCaseStageNode, "ACTIVITYID")
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"

     If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Moving to stage " & xmlGetAttributeText(xmlRequestCaseStageNode, "STAGEID"), vbLogEventTypeInformation
     End If
     
    ' get requested CASESTAGE detail =====================================================
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Getting requested stage", vbLogEventTypeInformation
    End If
    
    Set xmlNextStageDoc = New FreeThreadedDOMDocument40
    xmlNextStageDoc.validateOnParse = False
    xmlNextStageDoc.setProperty "NewParser", True
    GetRequestedStage vxmlRequestNode, xmlNextStageDoc
    Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")

    If xmlNextStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoNextStage
    End If

    ' remove non DEFAULT tasks
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Removing non default tasks", vbLogEventTypeInformation
    End If
    
    For Each xmlNode In xmlNextStageNode.childNodes
        If xmlGetAttributeAsBoolean(xmlNode, "DEFAULTFLAG") = False Then
            xmlNextStageNode.removeChild xmlNode
        End If
    Next

    ' get current CASESTAGE detail =====================================================
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Getting current stage", vbLogEventTypeInformation
    End If
    
    ' MAR1408 If moving to Completion Stage, add the TaskDueDate for each Post Completion task
    strNewStageId = xmlGetMandatoryAttributeText(xmlNextStageNode, "STAGEID")
    
    If (strNewStageId = GetGlobalParamString("TMCompletionsStageId")) Then
    
        'Get the time from Global Parameter
        strDueDateTime = GetGlobalParamString("TMPostCompTaskDueTime")
    
        'Get the Completion Date from Report On Title
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETREPORTONTITLEDATA"
        
        Set xmlROTNode = xmlApplicationRequestNode.cloneNode(False)
        xmlChangeNodeName xmlROTNode, "APPLICATION", "REPORTONTITLE"

        xmlNewRequest.appendChild xmlROTNode
        
        Set objROTBO = gobjContext.CreateInstance("omROT.omRotBO")
        xmlROTResponseDoc.loadXML objROTBO.OmRotRequest(xmlNewRequest.xml)
                                
        Set xmlTempNode = xmlROTResponseDoc.selectSingleNode("//REPORTONTITLE")
                                
        If Not (xmlTempNode Is Nothing) Then
            strCompletionDate = xmlGetAttributeText(xmlTempNode, "COMPLETIONDATE")
            strCompletionDate = strCompletionDate & " " & strDueDateTime
        End If
    
        If (Len(strCompletionDate) > 0) Then

            Set colValueIDs = New Collection
            GetValueIdsForValidationType "TaskType", "PC", colValueIDs

            For Each xmlNode In xmlNextStageNode.childNodes
        
                'If this is a Post Completion task, set up the Due Date based on Completion Date.
                                                    
                strTaskType = xmlGetAttributeText(xmlNode, "TASKTYPE")
                bCompletion = False
                For Each varValueId In colValueIDs
                    If strTaskType = varValueId Then
                        bCompletion = True
                    End If
                Next
                        
                If (bCompletion = True) Then
                    SetCaseTaskDueDateAndTime xmlNode, xmlNode, strCompletionDate
                End If
            Next
            
        End If
      
    End If
        
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")

    If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Non exception stage. Checking if incomplete mandatory tasks exist", vbLogEventTypeInformation
        End If

        If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
            If blnDebug Then
                App.LogEvent "MoveToStage (" & strCaseId & "):- Non exception stage. Incomplete mandatory tasks exist", vbLogEventTypeInformation
            End If
            errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
        End If
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- No incomplete tasks exist", vbLogEventTypeInformation
        End If

    End If

    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Checking if rules exist for next stage", vbLogEventTypeInformation
    End If
    
    If AnyRulesForStage(xmlNextStageNode) = True Then

        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Rules exist. Getting application data", vbLogEventTypeInformation
        End If
       
        ' get APPLICATION data =============================================================
        Set xmlRbResponseNode = GetApplicationData(xmlApplicationRequestNode, "TMRBTemplateEx")

        'Get required combos
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Rules exist. Combos", vbLogEventTypeInformation
        End If
        
        Set xmlComboDoc = GetComboList()
        For Each xmlComboNode In xmlComboDoc.selectNodes(".//COMBO")
            xmlRbResponseNode.appendChild xmlComboNode
        Next

        ' Run rules
        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
        xmlNextCaseStageDoc.validateOnParse = False
        xmlNextCaseStageDoc.setProperty "NewParser", True
        xmlNextCaseStageDoc.async = False

        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Running rules", vbLogEventTypeInformation
        End If
                
        RunUserRules _
            vxmlRequestNode, _
            xmlNextStageNode, _
            xmlRbResponseNode, _
            xmlNextCaseStageDoc
                 
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")

        'MAR1689  Copy task data (specifically the TaskDueDateAndTime) that has been set up
        ' PSC 15/05/2006 MAR1784 - Start
        If Not xmlNextCaseStageNode Is Nothing Then
            CopyTaskDataFromStage xmlNextStageNode, xmlNextCaseStageNode
        End If
        ' PSC 15/05/2006 MAR1784 - Start
    Else
    
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- No rules exist.", vbLogEventTypeInformation
        End If

        ' create CASESTAGE from STAGE
        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
        xmlNextCaseStageDoc.validateOnParse = False
        xmlNextCaseStageDoc.setProperty "NewParser", True
        xmlNextCaseStageDoc.async = False
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Creating CaseStage from Stage.", vbLogEventTypeInformation
        End If
        
        CreateCaseStageFromStage _
            xmlNextStageNode, _
            xmlNextCaseStageDoc
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")

    End If

    ' stage not applicable
    If xmlNextCaseStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmStageNotApplicable
    End If

    ' prepare CreateCaseStage request ==================================================
    Set xmlCreateCaseStageDoc = New FreeThreadedDOMDocument40
    xmlCreateCaseStageDoc.validateOnParse = False
    xmlCreateCaseStageDoc.setProperty "NewParser", True
    xmlCreateCaseStageDoc.async = False
    
    If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Prepare CreateCaseStageRequest", vbLogEventTypeInformation
    End If
    
    PrepareCreateCaseStageRequest _
        vxmlRequestNode, _
        xmlCurrentStageNode, _
        xmlNextCaseStageNode, _
        xmlCreateCaseStageDoc

    If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Adding default values toi case stage", vbLogEventTypeInformation
    End If
    
    'EP2_902
    AddDefaultValuesToCaseStage _
        gobjContext, _
        vxmlRequestNode, _
        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
        xmlNextStageNode

    ' prepare for update to current stage ==============================================
    If Not xmlCurrentStageNode Is Nothing Then

        ' first check we are auth'd to create the next stage SYS1788 JLD
        Set xmlValidateDoc = New FreeThreadedDOMDocument40
        xmlValidateDoc.validateOnParse = False
        xmlValidateDoc.setProperty "NewParser", True
    
        ' clone initial request node
        Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation name
        xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
        Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
        Set xmlNode = xmlNode.appendChild(xmlElem)
        ' copy STAGEID from next case stage
        xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Validating Stage Authority", vbLogEventTypeInformation
        End If

        RunTmRequest xmlValidateDoc

        Set xmlUpdateCurrentStageDoc = New FreeThreadedDOMDocument40
        xmlUpdateCurrentStageDoc.validateOnParse = False
        xmlUpdateCurrentStageDoc.setProperty "NewParser", True
        xmlUpdateCurrentStageDoc.async = False
        
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Preparing current stage update request", vbLogEventTypeInformation
        End If
        
        PrepareCurrentStageUpdateRequest _
            vxmlRequestNode, _
            xmlCurrentStageNode, _
            xmlUpdateCurrentStageDoc

        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False _
        Then
            If blnDebug Then
                App.LogEvent "MoveToStage (" & strCaseId & "):- Processing carry forward tasks", vbLogEventTypeInformation
            End If

            ProcessCarryForwardTasks _
                xmlCurrentStageNode, _
                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
        Else
            If blnDebug Then
                App.LogEvent "MoveToStage (" & strCaseId & "):- Processing cancel tasks", vbLogEventTypeInformation
            End If
            
            ProcessCancelTasks _
                xmlCurrentStageNode, _
                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE")
        End If

'        ' run both requests as a composite transaction
'        Dim xmlUpdateNodes As Collection
'        Set xmlUpdateNodes = New Collection
'        xmlUpdateNodes.Add xmlUpdateCurrentStageDoc.firstChild
'        xmlUpdateNodes.Add xmlCreateCaseStageDoc.firstChild
'
'        ' PSC 27/11/01 SYS3221 - Start
'        RunCompositeRequest xmlUpdateNodes
'
'    'AQR SYS2266 DRC
'    ' The New Transaction Class is used to ensure that the stage change is committed
'    ' before any automatic tasks are processed
'    '
''        Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
''        objNTxBO.RunCompositeRequest xmlUpdateNodes
'        Set xmlUpdateNodes = Nothing
'    Else
''        Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
''        objNTxBO.RunTmRequest xmlCreateCaseStageDoc
'        RunTmRequest xmlCreateCaseStageDoc
    End If
    ' PSC 27/11/01 SYS3221 - End

'    Set objNTxBO = Nothing
'
'    ' call omApp.ApplicationBO to create application stage on omiga ====================
'    CreateApplicationStage _
'        vxmlRequestNode, _
'        xmlApplicationRequestNode, _
'        xmlNextStageNode, _
'        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'
'    ' get New current (i.e. updated) CASESTAGE detail =====================================================
'    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
'    xmlCurrentStageDoc.async = False
'    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
'
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    vxmlResponseNode.appendChild xmlCurrentStageNode

    Set xmlApplicationRequestDoc = New FreeThreadedDOMDocument40
    xmlApplicationRequestDoc.validateOnParse = False
    xmlApplicationRequestDoc.setProperty "NewParser", True
    
    'BM0340 GHun 13/06/2003 Need to pass in xmlNextStageNode, NOT xmlNextCaseStageNode
    'PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextCaseStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Preparing Create Application Stage Request", vbLogEventTypeInformation
    End If
    
    PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    'BM0340 End
       
    ' PSC 22/09/2005 MAR32 - Start
    Set xmlSetSLADateDoc = New FreeThreadedDOMDocument40
    xmlSetSLADateDoc.validateOnParse = False
    xmlSetSLADateDoc.setProperty "NewParser", True
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Preparing Set SLA date Request", vbLogEventTypeInformation
    End If
    
    PrepareSetSLADateRequest xmlApplicationRequestNode, xmlNextStageNode, xmlSetSLADateDoc
    
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Calling MoveCaseToNewStage", vbLogEventTypeInformation
    End If
    
    If xmlSetSLADateDoc.childNodes.length = 0 Then
        objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc
    Else
        objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc, xmlSetSLADateDoc
    End If
    ' PSC 22/09/2005 MAR32 - End

    'BMIDS649 GHun 14/01/2004 ProcessAutomaticTasks needs to be called here, rather than
    ' from TM030, otherwise it doesn't work with ingestion
    
    'Remove the STAGEID attribute as it will now be out of date
    Set xmlElem = xmlGetNode(vxmlRequestNode, "CURRENTSTAGE")
    'BMIDS697 GHun 05/03/2004
    If xmlElem Is Nothing Then
        Set xmlElem = xmlGetNode(vxmlRequestNode, "CASESTAGE")
    End If
    'BMIDS697 End
    
    If Not xmlElem Is Nothing Then
        xmlElem.removeAttribute "STAGEID"
        'BMIDS681 Also remove CaseStageSequenceNumber/No if it exists, as it is most likely wrong
        xmlElem.removeAttribute "CASESTAGESEQUENCENUMBER"
        xmlElem.removeAttribute "CASESTAGESEQUENCENO"
        'BMIDS681 End
    End If
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Processing automatic tasks", vbLogEventTypeInformation
    End If

    ProcessAutomaticTasks vxmlRequestNode
    'BMIDS649 End

    'BMIDS681 GHun 16/02/2003
    ' Get the current stage
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- Getting current stage", vbLogEventTypeInformation
    End If
    
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    strStageId = xmlGetMandatoryAttributeText(xmlCurrentStageNode, "STAGEID")
    
    'MAR1334 DRC 01/03/2006
    'For Decline or Cancel  stage, need to cancel any disbursement payments
    If (strStageId = GetGlobalParamString("TMCancelStageID")) Or _
       (strStageId = GetGlobalParamString("TMDeclineStageID")) _
    Then
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Cancelling Disbursements", vbLogEventTypeInformation
        End If
        Set objPayProc = gobjContext.CreateInstance("omPayProc.PaymentProcessingBO")
        Set xmlRequest = vxmlRequestNode.cloneNode(False)
        Set xmlPayProcDoc = New FreeThreadedDOMDocument40
        xmlPayProcDoc.validateOnParse = False
        xmlPayProcDoc.setProperty "NewParser", True
        Set xmlElem = xmlPayProcDoc.createElement("HEADER")
        xmlElem.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlCurrentStageNode, "CASEID")
        xmlRequest.appendChild xmlElem
             
        xmlSetAttributeValue xmlRequest, "OPERATION", "CANCELDISBURSEMENTS"
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
              
        xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlRequest.xml))
        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
        
    End If
    'MAR1334 End
    
    blnProceed = True
        
    'Check if the stage contains only automatic tasks
    If HasIncompleteTasks(xmlCurrentStageNode) Then
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Stage has incomplete tasks", vbLogEventTypeInformation
        End If
        
        blnProceed = False
    Else
        'Check if it is possible to move from the current stage
        If Not (IsMovableStage(strActivityId, strStageId)) Then
            If blnDebug Then
                App.LogEvent "MoveToStage (" & strCaseId & "):- Is not movable stage", vbLogEventTypeInformation
            End If
            blnProceed = False
        End If
    End If
    
    If blnProceed Then
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Ok to progress to next stage", vbLogEventTypeInformation
        End If

        'Call MoveToNextStage
        Set xmlRequest = vxmlRequestNode.cloneNode(False)
        Set xmlElem = xmlCurrentStageDoc.createElement("CURRENTSTAGE")
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "STAGEID"
        'CASESTAGESEQUENCENO should not be passed
        xmlRequest.appendChild xmlElem
        xmlRequest.appendChild xmlApplicationRequestNode.cloneNode(True)
        'BMIDS681 DRC 15/03/04 Must pass the Printer node through for Automatic tasks
        Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("./PRINTER")
        
        ' PSC 20/01/2006 MAR1073 - Start
        If Not xmlPrinterNode Is Nothing Then
            xmlRequest.appendChild xmlPrinterNode.cloneNode(True)
        End If
        ' PSC 20/01/2006 MAR1073 - End
        
        'BMIDS681 DRC - End
        If blnDebug Then
            App.LogEvent "MoveToStage (" & strCaseId & "):- Calling move to next stage", vbLogEventTypeInformation
        End If

        MoveToNextStage xmlRequest, vxmlResponseNode
    End If
    'BMIDS681 End

MoveToStageExit:

    Set xmlNode = Nothing
    Set xmlRequestCaseStageNode = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlNextStageNode = Nothing
    Set xmlNextCaseStageNode = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlCreateCaseStageDoc = Nothing
    Set xmlNextStageDoc = Nothing
    Set xmlNextCaseStageDoc = Nothing
    Set xmlUpdateCurrentStageDoc = Nothing
    'Set xmlApplicationDataDoc = Nothing
    Set xmlValidateDoc = Nothing
    Set xmlElem = Nothing
    Set xmlRbResponseNode = Nothing
    Set xmlRequest = Nothing
    Set xmlSetSLADateDoc = Nothing  ' PSC 22/09/2005 MAR32
    
    If blnDebug Then
        App.LogEvent "MoveToStage (" & strCaseId & "):- End", vbLogEventTypeInformation
    End If

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode, _
                                 ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo MoveToNextStageExit
    Const cstrFunctionName As String = "MoveToNextStage"
       
    Dim xmlCurrentStageRequestNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextCaseStageDoc As FreeThreadedDOMDocument40
    Dim xmlCreateCaseStageDoc As FreeThreadedDOMDocument40
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlRbResponseNode As IXMLDOMNode
    Dim xmlNextCaseStageNode As IXMLDOMNode
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim xmlComboNode As IXMLDOMNode
    
    'BS BM0136 07/02/2003
    Dim xmlExceptionStageDoc As FreeThreadedDOMDocument40
    Dim xmlTaskNode As IXMLDOMNode
    'BS BM0136 End 07/02/2003

    Dim objNTxBO As IOmTmNTxBO
    
    Dim xmlValidateDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlTempRequest As IXMLDOMNode   ' PSC 19/05/2006 MAR1784
    
    'BMIDS681 GHun 10/02/2004
    Dim iIteration      As Integer
    Dim strStageId      As String
    Dim strActivityId   As String
    Dim blnProceed      As Boolean
    
    Dim xmlSetSLADateDoc As FreeThreadedDOMDocument40   ' PSC 22/09/2005
    
    iIteration = 0
    'BMIDS681 End
    
    ' validate request
    Set xmlCurrentStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTSTAGE")
    
    Dim blnDebug As Boolean
    Dim strCaseId As String
    blnDebug = IsDebugEnabled()
    strCaseId = xmlGetAttributeText(xmlCurrentStageRequestNode, "CASEID")

    If blnDebug Then
        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Start", vbLogEventTypeInformation
        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Request: " & vxmlRequestNode.xml, vbLogEventTypeInformation
        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Validating Request", vbLogEventTypeInformation
    End If
    
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "CASEID"
    'BMIDS681 Get the ActivityId
    strActivityId = xmlGetMandatoryAttributeText(xmlCurrentStageRequestNode, "ACTIVITYID")
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
        
    'BMIDS681 GHun 10/02/2004 Begin loop
    Do
    'BMIDS681 End
        ' get current CASESTAGE detail =====================================================
        Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
        xmlCurrentStageDoc.validateOnParse = False
        xmlCurrentStageDoc.setProperty "NewParser", True
        xmlCurrentStageDoc.async = False
        
        If blnDebug Then
            App.LogEvent "MoveToNextStage (" & strCaseId & "):- Getting current stage", vbLogEventTypeInformation
        End If
        
        GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
        Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
        
        'BMIDS681 GHun 10/02/2004
        strStageId = xmlGetMandatoryAttributeText(xmlCurrentStageNode, "STAGEID")
        
        blnProceed = True
        
        'Only perform the following checks for the 2nd and subsequent iterations
        If (iIteration > 0) Then
        
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Iteration = " & CStr(iIteration), vbLogEventTypeInformation
            End If

            'Check if the stage contains only automatic tasks
            If HasIncompleteTasks(xmlCurrentStageNode) Then
                If blnDebug Then
                    App.LogEvent "MoveToNextStage (" & strCaseId & "):- Stage has incomplete tasks", vbLogEventTypeInformation
                End If

                blnProceed = False
            Else
                'Check if it is possible to move from the current stage
                If Not (IsMovableStage(strActivityId, strStageId)) Then
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Is not moveable stage", vbLogEventTypeInformation
                    End If
                    blnProceed = False
                End If
            End If
        End If
        
        If blnProceed Then
        'BMIDS681 End
                
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Checking if incomplete mandatory tasks exist", vbLogEventTypeInformation
            End If
                
            If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
                If blnDebug Then
                    App.LogEvent "MoveToStage (" & strCaseId & "):- Non exception stage. Incomplete mandatory tasks exist", vbLogEventTypeInformation
                End If
                errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
            End If
            
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- No incomplete mandatory tasks exist", vbLogEventTypeInformation
            End If
                
            ' get next CASESTAGE detail ========================================================
            Set xmlNextStageDoc = New FreeThreadedDOMDocument40
            xmlNextStageDoc.validateOnParse = False
            xmlNextStageDoc.setProperty "NewParser", True
            xmlNextStageDoc.async = False
            
            Set xmlNextCaseStageNode = Nothing
            Do While xmlNextCaseStageNode Is Nothing
                    
                ' get next STAGE in sequence ===================================================
                If blnDebug Then
                    App.LogEvent "MoveToNextStage (" & strCaseId & "):- Get next stage in sequence", vbLogEventTypeInformation
                End If
                
                GetNextStageInSequence _
                    vxmlRequestNode, _
                    xmlCurrentStageNode, _
                    xmlNextStageNode, _
                    xmlNextStageDoc
                Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
                
                If blnDebug Then
                    App.LogEvent "MoveToNextStage (" & strCaseId & "):- Next stage is - " & xmlGetAttributeText(xmlNextStageNode, "STAGEID"), vbLogEventTypeInformation
                    App.LogEvent "MoveToNextStage (" & strCaseId & "):- Check if rules exist", vbLogEventTypeInformation
                End If
                
                If AnyRulesForStage(xmlNextStageNode) = True Then
            
                    If xmlRbResponseNode Is Nothing Then
                        If blnDebug Then
                            App.LogEvent "MoveToNextStage (" & strCaseId & "):- Rules exist. Getting application data", vbLogEventTypeInformation
                        End If
                        ' get APPLICATION data =============================================================
                        Set xmlRbResponseNode = GetApplicationData(xmlApplicationRequestNode, "TMRBTemplateEx")
                        
                        'Get required combos
                        If blnDebug Then
                            App.LogEvent "MoveToNextStage (" & strCaseId & "):- Rules exist. Getting combo data", vbLogEventTypeInformation
                        End If
                        Set xmlComboDoc = GetComboList()
                        For Each xmlComboNode In xmlComboDoc.selectNodes(".//COMBO")
                            xmlRbResponseNode.appendChild xmlComboNode
                        Next
                        
                    End If
                    
                    ' run rules ========================================================================
                    If xmlNextCaseStageDoc Is Nothing Then
                        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                        xmlNextCaseStageDoc.validateOnParse = False
                        xmlNextCaseStageDoc.setProperty "NewParser", True
                        xmlNextCaseStageDoc.async = False
                    Else
                        xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
                    End If
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Running rules", vbLogEventTypeInformation
                    End If

                    RunUserRules _
                        vxmlRequestNode, _
                        xmlNextStageNode, _
                        xmlRbResponseNode, _
                        xmlNextCaseStageDoc
                    Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
                    
                    'BS BM0136 07/02/2003
                    'If the case has been auto-declined/cancelled then the exception stageid and exception reason
                    'attributes will be returned in the CASESTAGE node of xmlNextCaseStageDoc
                    If Not xmlNextCaseStageNode Is Nothing Then
                        'Next stage is set. If it is an exception stage then change the CURRENTSTAGE node to CASESTAGE
                        'and copy the STAGEID and EXCEPTIONREASON attributes on the requestNode, re-populate
                        'xmlNextStageNode and xmlNextCaseStageNode with the exception stage/task details
        
                        'Is it an exception stage
                        If xmlAttributeValueExists(xmlNextCaseStageNode, "EXCEPTIONREASON") Then
                        
                            If blnDebug Then
                                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Next stage is exception stage", vbLogEventTypeInformation
                            End If
                            
                            ' PSC 19/05/2006 MAR1784 - Start
                            Set xmlTempRequest = vxmlRequestNode.cloneNode(True)

                            'Rename the CURRENTSTAGE node on the RequestNode to CASESTAGE and copy the
                            'STAGEID and EXCEPTIONREASON attributes
                            xmlSetAttributeValue xmlTempRequest.selectSingleNode("CURRENTSTAGE"), "STAGEID", xmlGetAttributeText(xmlNextCaseStageNode, "STAGEID")
                            xmlSetAttributeValue xmlTempRequest.selectSingleNode("CURRENTSTAGE"), "EXCEPTIONREASON", xmlGetAttributeText(xmlNextCaseStageNode, "EXCEPTIONREASON")
                            xmlChangeNodeName xmlTempRequest, "CURRENTSTAGE", "CASESTAGE"
                            ' PSC 19/05/2006 MAR1784 - End
                            
                            'Get exception stage/task details
                            Set xmlExceptionStageDoc = New FreeThreadedDOMDocument40
                            xmlExceptionStageDoc.validateOnParse = False
                            xmlExceptionStageDoc.setProperty "NewParser", True
                            
                            If blnDebug Then
                                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Getting requested stage", vbLogEventTypeInformation
                            End If
                            
                            GetRequestedStage xmlTempRequest, xmlExceptionStageDoc      ' PSC 19/05/2006 MAR1784
                            Set xmlNextStageNode = Nothing
                            Set xmlNextStageNode = xmlExceptionStageDoc.selectSingleNode("RESPONSE/STAGE")
                            
                            If xmlNextStageNode Is Nothing Then
                                 errThrowError cstrFunctionName, oeTmNoNextStage
                            End If
                            
                            'Double-Check it is an exception stage, if not then error
                            If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
                                errThrowError cstrFunctionName, oeTmNotExceptionStage
                            End If
                
                            ' remove non DEFAULT tasks
                            If blnDebug Then
                                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Remove non default tasks", vbLogEventTypeInformation
                            End If

                            For Each xmlTaskNode In xmlNextStageNode.childNodes
                                If xmlGetAttributeAsBoolean(xmlTaskNode, "DEFAULTFLAG") = False Then
                                    xmlNextStageNode.removeChild xmlTaskNode
                                End If
                            Next
                            
                            'Create CaseStage/CaseTask details
                            If xmlNextCaseStageDoc Is Nothing Then
                                Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                                xmlNextCaseStageDoc.validateOnParse = False
                                xmlNextCaseStageDoc.setProperty "NewParser", True
                                xmlNextCaseStageDoc.async = False
                            Else
                                xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
                            End If
                            
                            If blnDebug Then
                                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Creating Case Stage from Stage", vbLogEventTypeInformation
                            End If

                            CreateCaseStageFromStage _
                                xmlNextStageNode, _
                                xmlNextCaseStageDoc
                            Set xmlNextCaseStageNode = Nothing
                            Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
        
                        End If
                    End If
                    'BS BM0136 End 07/02/2003
                    
                Else
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- No rules", vbLogEventTypeInformation
                    End If

                    If xmlNextCaseStageDoc Is Nothing Then
                        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                        xmlNextCaseStageDoc.validateOnParse = False
                        xmlNextCaseStageDoc.setProperty "NewParser", True
                        xmlNextCaseStageDoc.async = False
                    Else
                        xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
                    End If
                    
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Create Case Stage from Stage", vbLogEventTypeInformation
                    End If

                    CreateCaseStageFromStage _
                        xmlNextStageNode, _
                        xmlNextCaseStageDoc
                    Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
                End If
                
            Loop
            
            ' find next CASESTAGE in rules response
            If xmlNextCaseStageNode Is Nothing Then
                errThrowError cstrFunctionName, oeTmNoNextStage
            End If
            
            ' prepare CreateCaseStage request ==================================================
            Set xmlCreateCaseStageDoc = New FreeThreadedDOMDocument40
            xmlCreateCaseStageDoc.validateOnParse = False
            xmlCreateCaseStageDoc.setProperty "NewParser", True
            
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Prepare create case stage request", vbLogEventTypeInformation
            End If

            PrepareCreateCaseStageRequest _
                vxmlRequestNode, _
                xmlCurrentStageNode, _
                xmlNextCaseStageNode, _
                xmlCreateCaseStageDoc
                
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Add default values to case stage", vbLogEventTypeInformation
            End If
            
            'EP2_902
            AddDefaultValuesToCaseStage _
                gobjContext, _
                vxmlRequestNode, _
                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                xmlNextStageNode
            
            ' prepare for update to current stage ==============================================
            ' add incomplete tasks from current stage to next stage ============================
            If Not xmlCurrentStageNode Is Nothing Then
                
                'BS BM0136 07/02/2003
                'Do not check authority level if case auto-declined/cancelled
                If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
                    ' first check we are auth'd to create the next stage   SYS1788 JLD
                    Set xmlValidateDoc = New FreeThreadedDOMDocument40
                    xmlValidateDoc.validateOnParse = False
                    xmlValidateDoc.setProperty "NewParser", True
                    ' clone initial request node
                    Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
                    ' set operation name
                    xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
                    Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
                    Set xmlNode = xmlNode.appendChild(xmlElem)
                    ' copy STAGEID from next case stage
                    xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
                    
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Validate Stage Authority", vbLogEventTypeInformation
                    End If

                    RunTmRequest xmlValidateDoc
                End If
                'BS BM0136 End 07/02/2003
                  
                Set xmlUpdateCurrentStageDoc = New FreeThreadedDOMDocument40
                xmlUpdateCurrentStageDoc.validateOnParse = False
                xmlUpdateCurrentStageDoc.setProperty "NewParser", True
                
                If blnDebug Then
                    App.LogEvent "MoveToNextStage (" & strCaseId & "):- Prepare current stage update request", vbLogEventTypeInformation
                End If

                PrepareCurrentStageUpdateRequest _
                    vxmlRequestNode, _
                    xmlCurrentStageNode, _
                    xmlUpdateCurrentStageDoc
                    
                'BS BM0136 07/02/2003
                'It is now possible for the next stage to be an exception stage (if it was auto-declined/cancelled).
                'If it isn't an exception stage then carry forward tasks as before, else cancel tasks
        '
        '        ProcessCarryForwardTasks _
        '            xmlCurrentStageNode, _
        '            xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
        '            xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
                If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
                    
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Process carry forward tasks", vbLogEventTypeInformation
                    End If

                    ProcessCarryForwardTasks _
                                xmlCurrentStageNode, _
                                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
                Else
                    If blnDebug Then
                        App.LogEvent "MoveToNextStage (" & strCaseId & "):- Process cancel tasks", vbLogEventTypeInformation
                    End If

                    ProcessCancelTasks _
                                xmlCurrentStageNode, _
                                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE")
                End If
                'BS BM0136 End 07/02/2003
                    
            End If
            ' PSC 27/11/01 SYS3221 - End
                
            Set xmlApplicationRequestDoc = New FreeThreadedDOMDocument40
            xmlApplicationRequestDoc.validateOnParse = False
            xmlApplicationRequestDoc.setProperty "NewParser", True
            
            'BM0340 GHun 13/06/2003 Need to pass in xmlNextStageNode, NOT xmlNextCaseStageNode
            'PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextCaseStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Prepare create application stage request", vbLogEventTypeInformation
            End If

            PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
            'BM0340 End
            
            ' PSC 22/09/2005 MAR32 - Start
            Set xmlSetSLADateDoc = New FreeThreadedDOMDocument40
            xmlSetSLADateDoc.validateOnParse = False
            xmlSetSLADateDoc.setProperty "NewParser", True
            
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Prepare set SLA date request", vbLogEventTypeInformation
            End If
            
            PrepareSetSLADateRequest xmlApplicationRequestNode, xmlNextStageNode, xmlSetSLADateDoc
            
            Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
            
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Calling MoveCaseToNewStage", vbLogEventTypeInformation
            End If
            
            If xmlSetSLADateDoc.childNodes.length = 0 Then
                objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc
            Else
                objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc, xmlSetSLADateDoc
            End If
            ' PSC 22/09/2005 MAR32 - End
           
            'BMIDS649 GHun 14/01/2004 ProcessAutomaticTasks needs to be called here, rather than
            ' from TM030, otherwise it doesn't work with ingestion
            
            'Remove the STAGEID attribute as it will now be out of date
            Set xmlElem = xmlGetNode(vxmlRequestNode, "CURRENTSTAGE")
            If Not xmlElem Is Nothing Then
                xmlElem.removeAttribute "STAGEID"
                'BMIDS681 Also remove CaseStageSequenceNumber/No if it exists, as it most likely wrong
                xmlElem.removeAttribute "CASESTAGESEQUENCENUMBER"
                xmlElem.removeAttribute "CASESTAGESEQUENCENO"
                'BMIDS681 End
            End If
            
            If blnDebug Then
                App.LogEvent "MoveToNextStage (" & strCaseId & "):- Processing Automatic tasks", vbLogEventTypeInformation
            End If
            
            ProcessAutomaticTasks vxmlRequestNode
            'BMIDS649 End
        
            'BMIDS681 GHun 10/02/2004
            iIteration = iIteration + 1
            
        End If
        
        'BMIDS754 GHun 20/04/2004 Clear xmlRbResponseNode so that the data will be refreshed
        ' for subsequent iterations of the loop, if necessary
        Set xmlRbResponseNode = Nothing
        'BMIDS754 End
        
    Loop Until Not (blnProceed)
    'BMIDS681 End
    
MoveToNextStageExit:
    Set xmlCurrentStageRequestNode = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlUpdateCurrentStageDoc = Nothing
    Set xmlNextStageNode = Nothing
    Set xmlRbResponseNode = Nothing
    Set xmlNextCaseStageNode = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlUpdateCurrentStageDoc = Nothing
    Set xmlNextStageDoc = Nothing
    Set xmlNextCaseStageDoc = Nothing
    Set xmlCreateCaseStageDoc = Nothing
    Set xmlValidateDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    'BS BM0136 07/02/2003
    Set xmlExceptionStageDoc = Nothing
    Set xmlTaskNode = Nothing
    'BS BM0136 End 07/02/2003
    Set xmlComboDoc = Nothing
    Set xmlComboNode = Nothing
    Set xmlSetSLADateDoc = Nothing  ' PSC 22/09/2005 MAR32
    
    Set xmlTempRequest = Nothing    ' PSC 19/05/2006 MAR1784
    
    If blnDebug Then
        App.LogEvent "MoveToNextStage (" & strCaseId & "):- End", vbLogEventTypeInformation
    End If
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Sub GetCurrentStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo GetCurrentStageExit
    Const cstrFunctionName As String = "GetCurrentStage"

    ' create GetCurrentStage request
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetCurrentStage"
    
    ' request childNode will be:
    '   CURRENTSTAGE for 'MoveToNextStage'
    '   CASESTAGE for 'MoveToStage'
    '   CASEACTIVITY for 'ReinstatePreviousStage'
    If vxmlRequestNode.firstChild.nodeName = "CASEACTIVITY" Then
        
        xmlRequestNode.appendChild vxmlRequestNode.firstChild.cloneNode(False)
    
    Else
        
        Set xmlCurrentStageNode = vxmlRequestNode.firstChild
            
        Set xmlElem = xmlRequestDoc.createElement("CASEACTIVITY")
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
        xmlRequestNode.appendChild xmlElem
        
    End If
        
    RunTmRequest xmlRequestDoc, vxmlCurrentStageDoc
    
    ' find CASESTAGE in response
    If vxmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "No current CASESTAGE detail"
    End If
    
GetCurrentStageExit:
    
    Set xmlRequestNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Function IncompleteMandatoryStageTasks( _
    ByVal vxmlCaseStageNode As IXMLDOMNode) _
    As Boolean
    
    On Error GoTo IncompleteMandatoryStageTasksExit
    Const cstrFunctionName As String = "IncompleteMandatoryStageTasks"
    
    Dim xmlNode As IXMLDOMNode
    
    IncompleteMandatoryStageTasks = False
    
    For Each xmlNode In vxmlCaseStageNode.childNodes
        If xmlGetAttributeAsBoolean(xmlNode, "MANDATORYINDICATOR") = True Then
            If IsIncompleteTask(xmlNode) = True Then
                IncompleteMandatoryStageTasks = True
                Exit For
            End If
        End If
    Next
    
IncompleteMandatoryStageTasksExit:
    
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName
    
End Function

Private Sub GetNextStageInSequence( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlNextStageNode As IXMLDOMNode, _
    ByVal vxmlNextStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo GetNextStageInSequenceExit
    Const cstrFunctionName As String = "GetNextStageInSequence"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' create GetNextStageTaskList request
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    ' copy input request
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    ' set operation name
    xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetNextStageTaskList"
    
    ' create CASESTAGE node for request
    Set xmlElem = xmlRequestDoc.createElement("CASESTAGE")
    
    ' copy current CASESTAGE attributes
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYID"
    
    If vxmlNextStageNode Is Nothing Then
        ' STAGEID from original request
        xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "STAGEID"
    Else
        ' STAGEID from previous 'GetNextStageTaskList' as rejected by rules
        xmlCopyAttribute vxmlNextStageNode, xmlElem, "STAGEID"
    End If
    
    'BG 11/11/01 SYS3458 Add CASEPRIORITY attribute to request for call
    Dim xmlAppPriorityElem As IXMLDOMElement
    
    Set xmlAppPriorityElem = vxmlRequestNode.selectSingleNode(".//APPLICATION[@APPLICATIONPRIORITY]")
    If Not xmlAppPriorityElem Is Nothing Then
        
        xmlElem.setAttribute "CASEPRIORITY", xmlAppPriorityElem.getAttribute("APPLICATIONPRIORITY")
        
    End If
    
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlRequestNode.appendChild xmlElem
    
    ' run base Tm GetNextStageTaskList request
    RunTmRequest xmlRequestDoc, vxmlNextStageDoc
    
    ' validate for next STAGE in response
    xmlCheckMandatoryNode vxmlNextStageDoc, "RESPONSE/STAGE"
    
GetNextStageInSequenceExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Function AnyRulesForStage(ByVal vxmlNextStageNode As IXMLDOMNode) _
    As Boolean
    
    On Error GoTo AnyRulesForStageExit
    Const cstrFunctionName As String = "AnyRulesForStage"
    
    If xmlAttributeValueExists(vxmlNextStageNode, "STAGERULEREFERENCE") = True Then
        AnyRulesForStage = True
        Exit Function
    End If
    
    Dim xmlNode As IXMLDOMNode
    For Each xmlNode In vxmlNextStageNode.childNodes
        If xmlAttributeValueExists(xmlNode, "TASKRULEREFERENCE") = True Then
            AnyRulesForStage = True
            Exit Function
        End If
    Next
    
    AnyRulesForStage = False
    
AnyRulesForStageExit:

    errCheckError cstrFunctionName
    
End Function

Private Sub RunUserRules(ByVal vxmlRequestNode As IXMLDOMNode, _
                            ByVal xmlNextStageNode As IXMLDOMNode, _
                            ByVal xmlRbResponseNode As IXMLDOMNode, _
                            ByVal vxmlNextCaseStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo RunUserRulesExit
    Const cstrFunctionName As String = "RunUserRules"

    ' create TmRules request
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlRbDataNode As IXMLDOMNode
    Dim objTmRules As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    'BMIDS691
    'Get EXCEPTIONREASON from request and put it in the request to send to OmTmRules
    
    Dim objComboDO As Object
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    
    Dim sException As String
    Dim sExceptionText As String
    
    If xmlAttributeValueExists(vxmlRequestNode.firstChild, "EXCEPTIONREASON") Then
        sException = xmlGetAttributeText(vxmlRequestNode.firstChild, "EXCEPTIONREASON")
        sExceptionText = objComboDO.GetComboText("ExceptionReason", sException)
    
        xmlSetAttributeValue xmlRequestNode, "EXCEPTIONREASON", sException
        xmlSetAttributeValue xmlRequestNode, "EXCEPTIONREASON_TEXT", sExceptionText
    End If
      
    ' clone GetNextStageTaskList response
    xmlRequestNode.appendChild xmlNextStageNode.cloneNode(True)
    
    ' add RB response data
    For Each xmlRbDataNode In xmlRbResponseNode.childNodes
        xmlRequestNode.appendChild xmlRbDataNode.cloneNode(True)
    Next
    
    Debug.Print "RunUserRules request"
    Debug.Print xmlRequestDoc.xml
    
    ' create TmRules component
    Set objTmRules = gobjContext.CreateInstance("OmTmRules.OmTmRulesBO")
    
    ' load response
    vxmlNextCaseStageDoc.loadXML objTmRules.RunRules(xmlRequestNode.xml)
    Set objTmRules = Nothing
    
    ' check response
    errCheckXMLResponseNode vxmlNextCaseStageDoc.selectSingleNode("RESPONSE"), , True
    
    Debug.Print "RunUserRules response"
    Debug.Print vxmlNextCaseStageDoc.xml
    
RunUserRulesExit:
    
    Set objTmRules = Nothing
    Set xmlRbDataNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub GetRequestedStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlNextStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo GetRequestedStageExit
    Const cstrFunctionName As String = "GetRequestedStage"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    ' create GetNextStageTaskList request
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    ' copy input request
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlRequestDoc.createElement("STAGE")
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "STAGEID"
    xmlRequestNode.appendChild xmlElem
    
    ' set operation name
    xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageDetail"
    
    RunTmRequest xmlRequestDoc, vxmlNextStageDoc
    
GetRequestedStageExit:
    
    Set xmlElem = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub CreateCaseStageFromStage( _
    ByVal vxmlNextStageNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo CreateCaseStageFromStageExit
    Const cstrFunctionName As String = "CreateCaseStageFromStage"

    Dim xmlNextCaseStageNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    ' clone original request
    Set xmlElem = vxmlNextCaseStageDoc.createElement("CASESTAGE")
    xmlCopyAttribute vxmlNextStageNode, xmlElem, "STAGEID"
    Set xmlNextCaseStageNode = vxmlNextCaseStageDoc.appendChild(xmlElem)
    
    For Each xmlNode In vxmlNextStageNode.childNodes
        Set xmlElem = vxmlNextCaseStageDoc.createElement("CASETASK")
        xmlCopyAttribute xmlNode, xmlElem, "TASKID"
        xmlCopyAttributeValue xmlNode, xmlElem, "TASKNAME", "CASETASKNAME"
        xmlCopyAttributeValue xmlNode, xmlElem, "MANDATORYFLAG", "MANDATORYINDICATOR"
        xmlCopyAttributeValue xmlNode, xmlElem, "TASKDUEDATEANDTIME", "TASKDUEDATEANDTIME"   'MAR1408
        xmlElem.setAttribute "TASKINSTANCE", "1"
        xmlNextCaseStageNode.appendChild xmlElem
    Next
    
CreateCaseStageFromStageExit:
    
    Set xmlNextCaseStageNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError cstrFunctionName
    
End Sub

'MAR1689  Add code to copy data from NextStage to NextCaseStage
Private Sub CopyTaskDataFromStage( _
    ByVal vxmlNextStageNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode)
    
    On Error GoTo CopyTaskDataFromStageExit
    Const cstrFunctionName As String = "CopyTaskDataFromStage"

    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlStageNode As IXMLDOMNode
    Dim sTaskID As String
    Dim strPattern As String
    
    For Each xmlCaseStageNode In vxmlNextCaseStageNode.childNodes
    
        sTaskID = xmlGetAttributeText(xmlCaseStageNode, "TASKID")
        strPattern = "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlCaseStageNode, "TASKID") & "']"
    
        Set xmlStageNode = vxmlNextStageNode.selectSingleNode(strPattern)
        
        If Not (xmlStageNode Is Nothing) Then
            xmlCopyAttributeValue xmlStageNode, xmlCaseStageNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
            xmlCopyAttributeValue xmlStageNode, xmlCaseStageNode, "TASKDUEDATEANDTIME", "TASKDUEDATEANDTIME"
            
            'EP2_902 copy REMOTEOWNERTASKIND
            xmlCopyAttributeValue xmlStageNode, xmlCaseStageNode, "REMOTEOWNERTASKIND", "REMOTEOWNERTASKIND"
            
            xmlSetAttributeValue xmlCaseStageNode, "TASKINSTANCE", "1"
            
        End If
    
    Next
    
CopyTaskDataFromStageExit:
    
    Set xmlCaseStageNode = Nothing
    Set xmlStageNode = Nothing
    
    errCheckError cstrFunctionName
    
End Sub


Private Sub PrepareCreateCaseStageRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
    ByVal vxmlGeneratedRequestDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo PrepareCreateCaseStageRequestExit
    Const cstrFunctionName As String = "PrepareCreateCaseStageRequest"

    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlRequestCaseStageNode As IXMLDOMNode
    
    Dim intCaseStageSeq As Integer
    
    vxmlGeneratedRequestDoc.async = False
    
    ' copy input request (no child nodes)
    Set xmlRequestNode = _
        vxmlGeneratedRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    ' set operation name
    xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseStage"
    
    ' create CASESTAGE request element
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("CASESTAGE")
    Set xmlCaseStageNode = xmlRequestNode.appendChild(xmlElem)
    
    ' copy CASEACTIVITYGUID from current CASESTAGE
    xmlCopyAttribute vxmlCurrentStageNode, xmlCaseStageNode, "CASEACTIVITYGUID"
    
    ' copy STAGEID from rules response
    xmlCopyAttribute vxmlNextCaseStageNode, xmlCaseStageNode, "STAGEID"
    
    ' copy ExceptionReason attribute
    Set xmlRequestCaseStageNode = vxmlRequestNode.selectSingleNode("CASESTAGE")
    If Not xmlRequestCaseStageNode Is Nothing Then
        xmlCopyAttribute xmlRequestCaseStageNode, xmlCaseStageNode, "EXCEPTIONREASON"
    End If
    
    ' generate CASESTAGESEQUENCENO from input request CASESTAGESEQUENCENO + 1
    intCaseStageSeq = CInt(xmlGetAttributeText(vxmlCurrentStageNode, "CASESTAGESEQUENCENO"))
    intCaseStageSeq = intCaseStageSeq + 1
    Set xmlAttrib = vxmlGeneratedRequestDoc.createAttribute("CASESTAGESEQUENCENO")
    xmlAttrib.Value = CStr(intCaseStageSeq)
    xmlCaseStageNode.Attributes.setNamedItem xmlAttrib
    
    ' for each task returned by rules
    For Each xmlNode In vxmlNextCaseStageNode.childNodes
        ' create CASETASK element
        Set xmlElem = vxmlGeneratedRequestDoc.createElement("CASETASK")
        ' copy TASKID from rule response
        xmlCopyAttribute xmlNode, xmlElem, "TASKID"
        ' TASKINSTANCE returned from rules ?
        If Not xmlNode.Attributes.getNamedItem("TASKINSTANCE") Is Nothing Then
            ' Yes, drop it (will be generated)
            xmlNode.Attributes.removeNamedItem "TASKINSTANCE"
        End If
        ' CASETASKNAME returned from rules ?
        If xmlNode.Attributes.getNamedItem("CASETASKNAME") Is Nothing Then
            ' No, use TASKNAME
            xmlCopyAttributeValue xmlNode, xmlElem, "TASKNAME", "CASETASKNAME"
        Else
            ' Yes, use returned value
            xmlCopyAttribute xmlNode, xmlElem, "CASETASKNAME"
        End If
        ' copy additional attributes from rules response
        xmlCopyAttribute xmlNode, xmlElem, "TASKDUEDATEANDTIME"
        xmlCopyAttribute xmlNode, xmlElem, "MANDATORYINDICATOR"
        xmlCopyAttribute xmlNode, xmlElem, "CONTEXT"
        xmlCopyAttribute xmlNode, xmlElem, "OWNINGUNITID"
        xmlCopyAttribute xmlNode, xmlElem, "OWNINGUSERID"
        xmlCopyAttribute xmlNode, xmlElem, "DATEOFOWNERSHIP"
        xmlCopyAttribute xmlNode, xmlElem, "TASKSTATUS"
        xmlCopyAttribute xmlNode, xmlElem, "TASKSTATUSSETDATETIME"
        xmlCopyAttribute xmlNode, xmlElem, "TASKSTATUSSETBYUSERID"
        xmlCopyAttribute xmlNode, xmlElem, "TASKSTATUSSETBYUNITID"
        xmlCopyAttribute xmlNode, xmlElem, "CUSTOMERIDENTIFIER"
        xmlCopyAttribute xmlNode, xmlElem, "ALLOCATEDBYUSERID"
        xmlCopyAttribute xmlNode, xmlElem, "ORIGINATINGSTAGEID"
        
        'EP2_902 copy REMOTEOWNERTASKIND
        xmlCopyAttribute xmlNode, xmlElem, "REMOTEOWNERTASKIND"
        
        ' append CASETASK node to CASESTAGE
        xmlCaseStageNode.appendChild xmlElem
    Next
    
PrepareCreateCaseStageRequestExit:
    
    Set xmlRequestNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlRequestCaseStageNode = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub RunTmRequest( _
    ByVal vxmlRequestDoc As FreeThreadedDOMDocument40, _
    Optional ByVal vxmlResponseDoc As FreeThreadedDOMDocument40 = Nothing)
    
    Const cstrFunctionName As String = "RunTmRequest"
    On Error GoTo RunTmRequestExit
    
    Debug.Print "RunTmRequest request"
    Debug.Print vxmlRequestDoc.xml
    
    If vxmlResponseDoc Is Nothing Then
        Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.async = False

        xmlThisResponseDoc.loadXML gobjMsgTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print xmlThisResponseDoc.xml
        Set xmlThisResponseDoc = Nothing
    Else
        vxmlResponseDoc.async = False
        vxmlResponseDoc.loadXML gobjMsgTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode vxmlResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print vxmlResponseDoc.xml
    End If
    
RunTmRequestExit:
    
    errCheckError cstrFunctionName
End Sub

Private Sub PrepareCurrentStageUpdateRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo PrepareCurrentStageUpdateRequestExit
    Const cstrFunctionName As String = "PrepareCurrentStageUpdateRequest"

    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlSrceNode As IXMLDOMNode
    
    vxmlUpdateCurrentStageDoc.async = False
    
    ' clone initial request node
    Set xmlNode = _
        vxmlUpdateCurrentStageDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation name
    xmlNode.Attributes.getNamedItem("OPERATION").Text = "SetCurrentCaseStageComplete"
    
    ' create CASESTAGE elem
    Set xmlElem = vxmlUpdateCurrentStageDoc.createElement("CASESTAGE")
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "CASEID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlRequestNode.firstChild, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "STAGEID"
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlNode.appendChild xmlElem
    
PrepareCurrentStageUpdateRequestExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlSrceNode = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub ProcessCarryForwardTasks( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageUpdateNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode)
    
    On Error GoTo ProcessCarryForwardTasksExit
    Const cstrFunctionName As String = "ProcessCarryForwardTasks"

    Dim xmlSrceNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim strPattern As String
    Dim enumUpdateStatus As TASKSTATUS
    
    'MO - 11/11/2002 BMIDS00812 - Start
    
    Dim xmlNewStageAllTasksDoc As FreeThreadedDOMDocument40
    Dim xmlNewStageAllTasks As IXMLDOMNode
    Dim xmlStageTask As IXMLDOMNode
    
    'Get all the tasks for the new stage we are moving too.
     ' create GetStageDetail request
    Set xmlNewStageAllTasksDoc = New FreeThreadedDOMDocument40
    xmlNewStageAllTasksDoc.validateOnParse = False
    xmlNewStageAllTasksDoc.setProperty "NewParser", True
    
    'clone the request
    xmlNewStageAllTasksDoc.appendChild vxmlCurrentStageUpdateNode.ownerDocument.documentElement.cloneNode(False)
    xmlSetAttributeValue xmlNewStageAllTasksDoc.documentElement, "OPERATION", "GetStageDetail"
    
    Set xmlElem = xmlNewStageAllTasksDoc.createElement("STAGE")
    'get the activity id and the stage id, from the new stage xml
    xmlCopyAttribute vxmlCurrentStageUpdateNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlNextCaseStageNode, xmlElem, "STAGEID"
    xmlNewStageAllTasksDoc.documentElement.appendChild xmlElem
    
    Set xmlElem = Nothing
    
    ' Get the stage tasks - run request & check response
    RunTmRequest xmlNewStageAllTasksDoc, xmlNewStageAllTasksDoc
    
    Set xmlNewStageAllTasks = xmlGetMandatoryNode(xmlNewStageAllTasksDoc.documentElement, "STAGE")
    
    'MO - 11/11/2002 BMIDS00812 - End
    
    'loop through the tasks
    For Each xmlSrceNode In vxmlCurrentStageNode.childNodes
    
        enumUpdateStatus = omiga4TASKSTATUSUndefined
    
        If IsIncompleteTask(xmlSrceNode) = True Then
            
            If xmlGetAttributeAsBoolean(xmlSrceNode, "CARRYFORWARDINDICATOR") = True Then
                enumUpdateStatus = omiga4TASKSTATUSCarriedForward
            Else
                enumUpdateStatus = omiga4TASKSTATUSNotApplicable
            End If
            
            Set xmlElem = _
                vxmlCurrentStageUpdateNode.ownerDocument.createElement("CASETASK")
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASEID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
            xmlCopyAttribute xmlSrceNode, xmlElem, "STAGEID"
            xmlCopyAttribute xmlSrceNode, xmlElem, "CASESTAGESEQUENCENO"
            xmlCopyAttribute xmlSrceNode, xmlElem, "TASKID"
            xmlCopyAttribute xmlSrceNode, xmlElem, "TASKINSTANCE"
            xmlCopyAttribute xmlSrceNode, xmlElem, "ORIGINATINGSTAGEID"
            xmlElem.setAttribute "TASKSTATUS", CStr(enumUpdateStatus)
            vxmlCurrentStageUpdateNode.appendChild xmlElem
            
            ' is it a carry forward task
            If enumUpdateStatus = omiga4TASKSTATUSCarriedForward Then
                strPattern = _
                    "CASETASK[@TASKID='" & _
                    xmlGetAttributeText(xmlSrceNode, "TASKID") & "'"
                    
                ' IK 30/10/01 AQR SYS2865
                                    
                If xmlAttributeValueExists(xmlSrceNode, "CUSTOMERIDENTIFIER") Then
                    
                    strPattern = strPattern & _
                        " and @CUSTOMERIDENTIFIER='" & _
                        xmlGetAttributeText(xmlSrceNode, "CUSTOMERIDENTIFIER") & "'"
                        
                End If
                                    
                If xmlAttributeValueExists(xmlSrceNode, "CONTEXT") Then
                    
                    strPattern = strPattern & _
                        " and @CONTEXT='" & _
                        xmlGetAttributeText(xmlSrceNode, "CONTEXT") & "'"
                        
                End If
                
                'MAR1476 make sure we copy customer tasks for all customers
                If xmlAttributeValueExists(xmlSrceNode, "CASETASKNAME") Then
                    
                    strPattern = strPattern & _
                        " and @CASETASKNAME='" & _
                        xmlGetAttributeText(xmlSrceNode, "CASETASKNAME") & "'"
                        
                End If
                
                strPattern = strPattern & "]"
                                   
                                
                'MO - 11/11/2002 BMIDS00812 - Start
                'Is this to be carried forward
                'Has the task already been added as a new task?
                If vxmlNextCaseStageNode.selectSingleNode(strPattern) Is Nothing Then
                    
                    'Does the task exist in the task for this new stage?
                    strPattern = "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlSrceNode, "TASKID") & "']"
                    Set xmlStageTask = xmlNewStageAllTasks.selectSingleNode(strPattern)
                    If Not xmlStageTask Is Nothing Then
                        
                        'create a new casetask in the new stage for this carried forward task
                        
                        Set xmlNode = vxmlNextCaseStageNode.ownerDocument.createElement("CASETASK")
                        
                        'copy the attributes from the original task defined in the stage and from the one carried forward
                        xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKID", "TASKID"
                        If xmlGetAttributeText(xmlSrceNode, "CASETASKNAME") = xmlGetAttributeText(xmlStageTask, "TASKNAME") Then
                            xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKNAME", "CASETASKNAME"
                        Else
                            'MAR1476 persist the casetaskname and copy the customeridentifier
                            xmlCopyAttributeValue xmlSrceNode, xmlNode, "CASETASKNAME", "CASETASKNAME"
                            xmlCopyAttributeValue xmlSrceNode, xmlNode, "CUSTOMERIDENTIFIER", "CUSTOMERIDENTIFIER"
                        End If
                        xmlCopyAttributeValue xmlStageTask, xmlNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "TASKDUEDATEANDTIME", "TASKDUEDATEANDTIME"
                        
                        'if the user id is specified in the stage take this one, else take the one carried forward
                        If xmlGetAttributeText(xmlStageTask, "TASKUSERID") > "" Then
                            xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKUSERID", "OWNINGUSERID"
                        Else
                            xmlCopyAttributeValue xmlSrceNode, xmlNode, "OWNINGUSERID", "OWNINGUSERID"
                        End If
                        
                        'and the same for the unit id
                        If xmlGetAttributeText(xmlStageTask, "TASKUNITID") > "" Then
                            xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKUNITID", "OWNINGUNITID"
                        Else
                            xmlCopyAttributeValue xmlSrceNode, xmlNode, "OWNINGUNITID", "OWNINGUNITID"
                        End If
                        
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "ORIGINATINGSTAGEID", "ORIGINATINGSTAGEID"
                        'AW 26/09/06  EP1155  - Start
                        'Set up attributes required for retrieving archived task notes
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "TASKINSTANCE", "PREVIOUSTASKINSTANCE"
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "STAGEID", "PREVIOUSSTAGEID"
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "CASESTAGESEQUENCENO", "PREVIOUSCASESTAGESEQNO"
                        xmlCopyAttributeValue xmlSrceNode, xmlNode, "CARRYFORWARDINDICATOR", "CARRYFORWARDINDICATOR"
                        'AW 26/09/06  EP1155  - End
                        
                        vxmlNextCaseStageNode.appendChild xmlNode
                        
                    End If
                                
                    'MO - 11/11/2002 BMIDS00812 - End
                                
                End If
                
            End If
        
        End If
    
    Next
    
ProcessCarryForwardTasksExit:

    Set xmlSrceNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub ProcessCancelTasks( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageUpdateNode As IXMLDOMNode)
    
    On Error GoTo ProcessCancelTasksExit
    Const cstrFunctionName As String = "ProcessCancelTasks"

    ModifyOutstandingTasks _
        vxmlCurrentStageNode, vxmlCurrentStageUpdateNode, omiga4TASKSTATUSCancelled
        
ProcessCancelTasksExit:
    
    errCheckError cstrFunctionName

End Sub

Private Function IsIncompleteTask(ByVal vxmlCaseTaskNode As IXMLDOMNode) _
    As Boolean
    
    On Error GoTo IsIncompleteTaskExit
    Const cstrFunctionName As String = "IsIncompleteTask"
    
    Dim intStatus As Integer
    
    IsIncompleteTask = False
    intStatus = xmlGetAttributeAsInteger(vxmlCaseTaskNode, "TASKSTATUS")
    If intStatus = omiga4TASKSTATUSUndefined Or _
        intStatus = omiga4TASKSTATUSIncomplete Or _
        intStatus = omiga4TASKSTATUSPending Or _
        intStatus = omiga4TASKSTATUSTASInProgress Or _
        intStatus = omiga4TASKSTATUSTASRetry Or _
        intStatus = omiga4TASKSTATUSTASFailed _
    Then
        IsIncompleteTask = True
    End If
    
IsIncompleteTaskExit:
    
    errCheckError cstrFunctionName
    
End Function

Private Sub GetApplicationOwners(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByRef strUserId As String, ByRef strUnitId As String)

On Error GoTo GetApplicationOwnersExit

Dim xmlFindAppOwnerShipListDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlAppOwnerShipNode As IXMLDOMElement
Dim xmlAppNumberNode As IXMLDOMElement
Dim objAppManBO As Object
Dim xmlTempNodeList As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode

Dim strFindAppOwnerShipListDoc As String
Dim lngErrNo As Long
Dim dteTempDate As Date
Dim dteMaxDate As Date
Dim intListCount  As Integer
Dim intMaxIndex  As Integer
Dim intListIndex As Integer

Const cstrFunctionName As String = "GetApplicationOwners"

    Set xmlFindAppOwnerShipListDoc = New FreeThreadedDOMDocument40
    xmlFindAppOwnerShipListDoc.validateOnParse = False
    xmlFindAppOwnerShipListDoc.setProperty "NewParser", True
    xmlFindAppOwnerShipListDoc.async = False
    
    'Formulate the Request
    Set xmlNode = xmlFindAppOwnerShipListDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlAppOwnerShipNode = xmlFindAppOwnerShipListDoc.createElement("APPLICATIONOWNERSHIP")
    Set xmlAppNumberNode = xmlFindAppOwnerShipListDoc.createElement("APPLICATIONNUMBER")
    
    'Search for CASEID attrib in Request
    xmlAppNumberNode.Text = xmlGetMandatoryNodeText(vxmlRequestNode, ".//*/@CASEID")
    xmlAppOwnerShipNode.appendChild xmlAppNumberNode
    xmlNode.appendChild xmlAppOwnerShipNode
    
    'Call ApplicationManagerBO
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strFindAppOwnerShipListDoc = objAppManBO.FindApplicationOwnershipList(xmlNode.xml)
    xmlFindAppOwnerShipListDoc.loadXML strFindAppOwnerShipListDoc
    Set xmlAppOwnerShipNode = xmlFindAppOwnerShipListDoc.documentElement
    
    'Process response
    lngErrNo = errCheckXMLResponseNode(xmlAppOwnerShipNode, , False)
    If lngErrNo <> 0 Then
        'Check for record not found error
        If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
            'No Application owner
            strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
            strUnitId = xmlGetMandatoryAttributeText(vxmlRequestNode, "UNITID")
        Else
            'raise error and exit
            errCheckXMLResponseNode xmlAppOwnerShipNode, , True
        End If
    Else
        'Find the latest owner details
        Set xmlTempNodeList = xmlGetMandatoryNode(xmlAppOwnerShipNode, ".//USERHISTORYLIST")
        dteMaxDate = "01/01/1800"
        intListCount = xmlTempNodeList.childNodes.length
        
        If intListCount > 0 Then
            For intListIndex = 0 To (intListCount - 1)
                'Set xmlTempNode = xmlTempNodeList.childNodes.Item(intListIndex).selectSingleNode("USERHISTORYDATE")
                'dteTempDate = CDate(xmlTempNode.Text)
                dteTempDate = xmlGetMandatoryNodeAsDate(xmlTempNodeList.childNodes.Item(intListIndex), "USERHISTORYDATE")
                If dteTempDate > dteMaxDate Then
                    dteMaxDate = dteTempDate
                    intMaxIndex = intListIndex ' The item number of the USERHISTORY which we want
                End If
            Next
            'Extract User and Unit ID
            strUserId = xmlGetNodeText(xmlTempNodeList.childNodes.Item(intMaxIndex), ".//USERID")
            strUnitId = xmlGetNodeText(xmlTempNodeList.childNodes.Item(intMaxIndex), ".//UNITID")
        Else
            'No Application owner
            strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
            strUnitId = xmlGetMandatoryAttributeText(vxmlRequestNode, "UNITID")
        End If
        
    End If
        
GetApplicationOwnersExit:
    Set xmlFindAppOwnerShipListDoc = Nothing
    Set xmlNode = Nothing
    Set xmlAppOwnerShipNode = Nothing
    Set xmlAppNumberNode = Nothing
    Set objAppManBO = Nothing
    Set xmlTempNodeList = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ModifyOutstandingTasks( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageUpdateNode As IXMLDOMNode, _
    ByVal venumNewStatus As TASKSTATUS)
    
    On Error GoTo ModifyOutstandingTasksExit
    Const cstrFunctionName As String = "ModifyOutstandingTasks"

    Dim xmlSrceNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
       
    For Each xmlSrceNode In vxmlCurrentStageNode.childNodes
    
        If IsIncompleteTask(xmlSrceNode) = True Then
            
            Set xmlElem = _
                vxmlCurrentStageUpdateNode.ownerDocument.createElement("CASETASK")
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASEID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
            xmlCopyAttribute xmlSrceNode, xmlElem, "STAGEID"
            xmlCopyAttribute xmlSrceNode, xmlElem, "CASESTAGESEQUENCENO"
            xmlCopyAttribute xmlSrceNode, xmlElem, "TASKID"
            xmlCopyAttribute xmlSrceNode, xmlElem, "TASKINSTANCE"
            xmlElem.setAttribute "TASKSTATUS", CStr(venumNewStatus)
            vxmlCurrentStageUpdateNode.appendChild xmlElem
        
        End If
    
    Next
    
ModifyOutstandingTasksExit:

    Set xmlSrceNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Function GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                                            Optional ByVal vstrRBTemplateName As String) As IXMLDOMNode
    
On Error GoTo GetApplicationDataExit

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlRequestElem As IXMLDOMElement
Dim xmlAppElem As IXMLDOMElement

Dim objRB As Object

Dim strRetXML As String

Const cstrFunctionName As String = "GetApplicationData"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlRequestElem = xmlDoc.createElement("REQUEST")
    xmlRequestElem.setAttribute "COMBOLOOKUP", "NO"
    If Len(vstrRBTemplateName) > 0 Then
        xmlRequestElem.setAttribute "RB_TEMPLATE", vstrRBTemplateName
    End If
    
    Set xmlAppElem = xmlDoc.createElement("APPLICATION")
    xmlAppElem.setAttribute "_SCHEMA_", "APPLICATION"
    xmlAppElem.setAttribute "APPLICATIONNUMBER", vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONNUMBER").Text
    
    If xmlAttributeValueExists(vxmlApplicationSrceNode, "APPLICATIONFACTFINDNUMBER") Then
        xmlCopyAttribute vxmlApplicationSrceNode, xmlAppElem, "APPLICATIONFACTFINDNUMBER"
    Else
        xmlAppElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    End If
    
    xmlCopyAttribute vxmlApplicationSrceNode, xmlAppElem, "_EXTRACTTYPE_"
    xmlRequestElem.appendChild xmlAppElem
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    
    strRetXML = objRB.OmDataRequest(xmlRequestElem.xml)
    
    xmlDoc.loadXML strRetXML
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    If xmlGetNode(xmlDoc.documentElement, "APPLICATION") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "No APPLICATION detail"
    End If
    
    Set GetApplicationData = xmlDoc.documentElement
    
GetApplicationDataExit:
    Set objRB = Nothing
    Set xmlDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAppElem = Nothing
    
    errCheckError cstrFunctionName
    
End Function

'=======================================================================================================
' BM0340 Functions copied from omTmBO to support MoveToNextStage - End
'=======================================================================================================

Private Sub PrepareCreateApplicationStageRequest( _
    ByVal vxmlAppNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
    ByVal vxmlGeneratedRequestDoc As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' procedure:    PrepareCreateApplicationStageRequest
' author:       Mark Coleman
' date:         24/04/2003
' description:
' pass:         vxmlAppNode  XML Application data
'               vxmlCurrentStageNode XML Current Stage data
'               vxmlNextCaseStageNode XML Next Stage data
'               vxmlGeneratedRequestDoc XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo PrepareCreateApplicationStageRequest_Exit

Dim xmlRequest As IXMLDOMNode
Dim xmlAppStageNode As IXMLDOMNode
Dim xmlElem As IXMLDOMElement

Const cstrFunctionName As String = "PrepareCreateApplicationStageRequest"
    
    Set xmlRequest = vxmlGeneratedRequestDoc.createElement("REQUEST")
    vxmlGeneratedRequestDoc.appendChild xmlRequest

    Set xmlAppStageNode = vxmlGeneratedRequestDoc.createElement("APPLICATIONSTAGE")
    xmlRequest.appendChild xmlAppStageNode
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlAppNode, "APPLICATIONNUMBER")
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlAppNode, "APPLICATIONFACTFINDNUMBER")
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("STAGENUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlCurrentStageNode, "STAGEID")
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("STAGENAME")
    xmlElem.Text = xmlGetAttributeText(vxmlCurrentStageNode, "STAGENAME")
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("DATETIME")
    xmlElem.Text = Now()
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("STAGESEQUENCENO")
    xmlElem.Text = xmlGetAttributeText(vxmlCurrentStageNode, "STAGESEQUENCENO")
    xmlAppStageNode.appendChild xmlElem
    
    Set xmlElem = vxmlGeneratedRequestDoc.createElement("CASESTAGESEQUENCENO")
    xmlElem.Text = xmlGetAttributeText(vxmlNextCaseStageNode, "CASESTAGESEQUENCENO")
    xmlAppStageNode.appendChild xmlElem

PrepareCreateApplicationStageRequest_Exit:
    Set xmlRequest = Nothing
    Set xmlAppStageNode = Nothing
    Set xmlElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

Private Function GetComboList() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' procedure:    GetComboList
' author:       Mark Coleman
' date:         25/04/2003
' description:
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetComboList_Exit

Dim xmlTemplateDoc As FreeThreadedDOMDocument40
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlNode As IXMLDOMNode
Dim adoConn As ADODB.Connection
Dim adoCmd As ADODB.Command
Dim adoStr As ADODB.Stream
Dim adoParam As ADODB.Parameter
Dim objFSO As FileSystemObject

Dim strFileName As String
Dim strWhere As String

Const cstrFunctionName As String = "GetComboList"
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlNode = xmlDoc.createElement("COMBOLIST")
    xmlDoc.appendChild xmlNode
    
    'Read in combo tags from TMRBTemplateCombos ==============================================================
    strFileName = App.Path & "\TMRBTEMPLATECOMBOS.XML"
    strFileName = Replace(strFileName, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set objFSO = New FileSystemObject
    If objFSO.FileExists(strFileName) Then
    
        Set xmlTemplateDoc = New FreeThreadedDOMDocument40
        xmlTemplateDoc.validateOnParse = False
        xmlTemplateDoc.setProperty "NewParser", True
        xmlTemplateDoc.async = False
        If Not xmlTemplateDoc.Load(strFileName) Then
            errThrowError cstrFunctionName, oeXMLParserError, "Unable to load TMRBTEMPLATECOMBOS.XML"
        End If
        
        'Create a SQL Where clause from them ===============================================================
        Set xmlNodeList = xmlTemplateDoc.selectNodes(".//COMBO")
        If xmlNodeList.length > 0 Then
            For Each xmlNode In xmlNodeList
                If Len(strWhere) > 0 Then
                    strWhere = strWhere & ", "
                End If
                strWhere = strWhere & "'" & xmlGetMandatoryAttributeText(xmlNode, "GROUPNAME") & "'"
            Next
            
            'Execute the procedure and return XML ==============================================================
            Set adoConn = New ADODB.Connection
            adoConn.ConnectionString = adoGetDbConnectString
            
            Set adoCmd = New ADODB.Command
            With adoCmd
                .CommandType = adCmdStoredProc
                .CommandText = "usp_GetComboListXML"
            End With
            
            'Parameters
            Set adoParam = New ADODB.Parameter
            With adoParam
                .Name = "SSQLWHERE"
                .Type = adVarChar
                .Direction = adParamInput
                .Size = 1000
                .Value = strWhere
            End With
            adoCmd.Parameters.Append adoParam
            
            Set adoStr = New ADODB.Stream
            adoConn.Open
            adoStr.Open
            
            adoCmd.ActiveConnection = adoConn
            adoCmd.Properties("Output Stream") = adoStr
            adoCmd.Properties("XML Root") = "COMBOLIST"
            adoCmd.Execute , , adExecuteStream
            
            If Not xmlDoc.loadXML(adoStr.ReadText()) Then
                errThrowError cstrFunctionName, oeXMLParserError, "Unable to load response - " & xmlDoc.parseError.reason
            End If
        
        End If
    
    End If
    
    Set GetComboList = xmlDoc
    
GetComboList_Exit:
    If Not adoStr Is Nothing Then
        If adoStr.State = adStateOpen Then
            adoStr.Close
        End If
    End If
    
    If Not adoConn Is Nothing Then
        If adoConn.State = adStateOpen Then
            adoConn.Close
        End If
    End If
    
    Set xmlTemplateDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set adoConn = Nothing
    Set adoCmd = Nothing
    Set adoStr = Nothing
    Set adoParam = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

'BMIDS658 GHun 26/11/2003
'RunExperianCreditCheck and ImportBureauData both used identical code to find customers
' for the current application, so the duplicate code has been moved to a single function
' which they both can call. This has had to be moved to omTmNoTxBO so that locks created
' by reading the data are released before any updates begin, otherwise deadlocks may
' result under load
Private Sub FindCustomersForCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    FindCustomersForCreditCheck
' author:       GHun
' date:         27/11/2003
' description:  Find customers for a specified application number
' pass:         vxmlRequestNode  XML Input request specifying the application number and fact find number
'               vxmlResponseNode XML Response to add the customerlist data to
' return:       N/A
'------------------------------------------------------------------------------------------
    Const cstrFunctionName  As String = "FindCustomersForCreditCheck"
    Dim xmlResponseDoc      As FreeThreadedDOMDocument40
    Dim xmlCustomerNodeList As IXMLDOMNodeList
    Dim xmlCustomerNode     As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlTempRequest      As IXMLDOMNode
    Dim xmlNewCustomerNode  As IXMLDOMNode
    Dim xmlTempNode         As IXMLDOMNode
    Dim objAppManagerBO     As Object
    Dim strResponse         As String
    Dim strCustomerNumber   As String
    Dim strPattern          As String
    Dim xmlCustRoleNode     As IXMLDOMNode
    
On Error GoTo FindCustomersForCreditCheck_VBErr
    
    'Get a list of customers on the current application ===========================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.appendChild xmlApplicationNode

    Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strResponse = objAppManagerBO.FindCustomersForApplication(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True

    Set xmlCustomerListNode = xmlResponseDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerNodeList = xmlResponseDoc.documentElement.selectNodes("CUSTOMER/CUSTOMERVERSION")
    For Each xmlCustomerNode In xmlCustomerNodeList
        Set xmlNewCustomerNode = xmlResponseDoc.createElement("CUSTOMER")
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "CUSTOMERNUMBER")
        strCustomerNumber = xmlTempNode.Text
        xmlNewCustomerNode.appendChild xmlTempNode
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "CUSTOMERVERSIONNUMBER")
        xmlNewCustomerNode.appendChild xmlTempNode
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "FIRSTFORENAME")
        xmlNewCustomerNode.appendChild xmlTempNode
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "SURNAME")
        xmlNewCustomerNode.appendChild xmlTempNode
        'GD BM0373 05/03/2003 START
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "DATEOFBIRTH")
        xmlNewCustomerNode.appendChild xmlTempNode
        'GD BM0373 05/03/2003 END
        'JD BMIDS832 Add CustomerOrder
        strPattern = "CUSTOMERROLELIST/CUSTOMERROLE[CUSTOMERNUMBER='" & strCustomerNumber & "']"
        Set xmlCustRoleNode = xmlResponseDoc.documentElement.selectSingleNode(strPattern)
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustRoleNode, "CUSTOMERORDER")
        xmlNewCustomerNode.appendChild xmlTempNode
        
        xmlCustomerListNode.appendChild xmlNewCustomerNode
    Next
    
    vxmlResponseNode.appendChild xmlCustomerListNode
    
FindCustomersForCreditCheck_VBErr:
    
    Set xmlResponseDoc = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNewCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set objAppManagerBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'BMIDS658 End

'BMIDS658 GHun 27/11/2003
Private Sub GetBureauRecords(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    GetBureauRecords
' author:       GHun
' date:         27/11/2003
' description:  Calls CreditCheckBO.GetAndDeduplicateBureauRecords outside of a transaction
'               to reduce the period of time that locks on the records are held
' pass:         vxmlRequestNode  XML Input request specifying the application number and fact find number
'               vxmlResponseNode XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    Const cstrFunctionName  As String = "GetBureauRecords"
    Dim objCCBO     As Object
    Dim strResponse As String
    Dim xmlDoc      As FreeThreadedDOMDocument40
    Dim xmlData     As IXMLDOMNode
    
On Error GoTo GetBureauRecords_VbErr

    Set objCCBO = gobjContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
    strResponse = objCCBO.GetAndDeduplicateBureauRecords(vxmlRequestNode.xml)
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, , True
    
    Set xmlData = xmlGetMandatoryNode(xmlDoc.documentElement, "BUREAUDATA")
    vxmlResponseNode.appendChild xmlData
    
GetBureauRecords_VbErr:
    Set xmlDoc = Nothing
    Set objCCBO = Nothing
    Set xmlData = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'BMIDS658 End

'BMIDS681 GHun 13/02/2003
Private Function IsMovableStage(ByVal strActivityId As String, ByVal strStageId As String) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    IsMovableStage
' author:       GHun
' date:         13/02/2003
' description:  Check if it is appropriate to move from the current stage. This validation
'               only applies when automatically moving stages, and does not apply for
'               completions or reinstatements.
' pass:         strActivityId   String containing the id of the currect activity
'               strStageId      String containing the id of the current stage
' return:       Boolean indicating if the stage can be moved from or not
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "IsMovableStage"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequestElem      As IXMLDOMElement
    Dim xmlStageElem        As IXMLDOMElement
    Dim xmlCompStage        As IXMLDOMNode
    Dim xmlPreCompStage     As IXMLDOMElement
    Dim xmlLastStageElem    As IXMLDOMElement
    Dim xmlElem             As IXMLDOMElement
    Dim strCompStageId      As String
    Dim strLastStageId      As String
    Dim strPreCompStageId   As String
    Dim strResponse         As String
    Dim blnResult           As Boolean
    
On Error GoTo IsMovableStageExit
    
    blnResult = True
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    'Get the stage list
    Set xmlRequestElem = xmlDoc.createElement("REQUEST")
    xmlRequestElem.setAttribute "OPERATION", "GetStageList"
    
    Set xmlStageElem = xmlDoc.createElement("STAGE")
    xmlStageElem.setAttribute "ACTIVITYID", strActivityId
    xmlStageElem.setAttribute "DELETEFLAG", "0"
    xmlRequestElem.appendChild xmlStageElem
        
    strResponse = gobjMsgTmBO.TmRequest(xmlRequestElem.xml)
    xmlDoc.loadXML strResponse
    errCheckXMLResponseNode xmlDoc.documentElement, , True
            
    'Check if the current stage is an exception stage
    For Each xmlStageElem In xmlDoc.selectNodes("RESPONSE/STAGE[@EXCEPTIONSTAGEINDICATOR='1']")
        If xmlStageElem.getAttribute("STAGEID") = strStageId Then
            blnResult = False
        End If
    Next
            
    If blnResult Then
        'Check if the current stage is the precompletion stage
        strCompStageId = GetMandatoryGlobalParamString("TMCompletionsStageId")
                
        Set xmlCompStage = xmlDoc.selectSingleNode("RESPONSE/STAGE[@STAGEID='" & strCompStageId & "']")
        
        If Not xmlCompStage Is Nothing Then
            Set xmlPreCompStage = xmlCompStage.previousSibling
            If Not xmlPreCompStage Is Nothing Then
                strPreCompStageId = xmlPreCompStage.getAttribute("STAGEID")
            End If
        End If
        
        If strStageId = strPreCompStageId Then
            blnResult = False
        End If
    
        If blnResult Then
    
            'Check if the current stage is the last stage
            Set xmlLastStageElem = xmlDoc.selectSingleNode("RESPONSE/STAGE[last()]")
            strLastStageId = xmlLastStageElem.getAttribute("STAGEID")
            
            If strStageId = strLastStageId Then
                blnResult = False
            End If
        End If
    End If
    
    IsMovableStage = blnResult
    
IsMovableStageExit:
    Set xmlDoc = Nothing
    Set xmlRequestElem = Nothing
    Set xmlStageElem = Nothing
    Set xmlCompStage = Nothing
    Set xmlPreCompStage = Nothing
    Set xmlLastStageElem = Nothing
    Set xmlElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'BMIDS681 End

'BMIDS681 GHun 16/02/2004
Private Function HasIncompleteTasks(ByVal vxmlCaseStageNode As IXMLDOMNode) As Boolean
' header ----------------------------------------------------------------------------------
' procedure:    HasIncompleteTasks
' author:       GHun
' date:         13/02/2003
' description:  Check if any of the tasks on the CaseStage node passed in are incomplete.
' pass:         vxmlCaseStageNode   XML input containing a list of tasks for the current CaseStage
' return:       Boolean indicating if any tasks are incomplete
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "HasIncompleteTasks"
    Dim xmlNode As IXMLDOMNode
    Dim blnCarryForward As Boolean  'MAR1198 GHun
    
On Error GoTo HasIncompleteTasksExit
    
    HasIncompleteTasks = False
    
    For Each xmlNode In vxmlCaseStageNode.childNodes
        blnCarryForward = xmlGetAttributeAsBoolean(xmlNode, "CARRYFORWARDINDICATOR")    'MAR1198 GHun
        If IsIncompleteTask(xmlNode) And Not (blnCarryForward) Then                     'MAR1198 GHun
            HasIncompleteTasks = True
            Exit For
        End If
    Next
    
HasIncompleteTasksExit:
    
    Set xmlNode = Nothing
    errCheckError cstrFunctionName
    
End Function
'BMIDS681 End

'BMIDS697 GHun 04/03/2004 Moved here from omTmBO
Private Sub DeclineApplication(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo DeclineApplicationExit

    Const cstrFunctionName As String = "DeclineApplication"
        
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCaseStageRequestNode As IXMLDOMNode
    Dim xmlCaseTaskRequestNode As IXMLDOMNode
    Dim xmlUpdateCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlMoveToStageDoc As FreeThreadedDOMDocument40
    Dim xmlMoveToStageRequestNode As IXMLDOMNode
    Dim xmlUpdateCaseTaskRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim xmlPrinterNode As IXMLDOMNode
        
    ' Validate Request
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    Set xmlCaseTaskRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlPrinterNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTER")
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONPRIORITY"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "TASKID"
    
    ' Call UpdateCaseTask
    Set xmlUpdateCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlUpdateCaseTaskDoc.validateOnParse = False
    xmlUpdateCaseTaskDoc.setProperty "NewParser", True
    Set xmlUpdateCaseTaskRequestNode = _
        xmlUpdateCaseTaskDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlUpdateCaseTaskRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CaseTask element
    Set xmlElem = xmlUpdateCaseTaskDoc.createElement("CASETASK")
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "TASKINSTANCE"
    xmlSetAttributeValue xmlElem, "TASKSTATUS", omiga4TASKSTATUSComplete
    xmlUpdateCaseTaskRequestNode.appendChild xmlElem
    RunTmRequest xmlUpdateCaseTaskDoc
    
    ' Call MoveToStage
    Set xmlMoveToStageDoc = New FreeThreadedDOMDocument40
    xmlMoveToStageDoc.validateOnParse = False
    xmlMoveToStageDoc.setProperty "NewParser", True
    Set xmlMoveToStageRequestNode = _
        xmlMoveToStageDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlMoveToStageRequestNode.Attributes.getNamedItem("OPERATION").Text = "MoveToStage"
    xmlMoveToStageRequestNode.appendChild xmlCaseStageRequestNode.cloneNode(True)
    xmlMoveToStageRequestNode.appendChild xmlApplicationRequestNode.cloneNode(True)
    xmlMoveToStageRequestNode.appendChild xmlPrinterNode.cloneNode(True)
    
    Set xmlResponse = xmlMoveToStageDoc.createElement("RESPONSE")
    MoveToStage xmlMoveToStageRequestNode, xmlResponse

DeclineApplicationExit:

    Set xmlApplicationRequestNode = Nothing
    Set xmlCaseStageRequestNode = Nothing
    Set xmlCaseTaskRequestNode = Nothing
    Set xmlUpdateCaseTaskDoc = Nothing
    Set xmlUpdateCaseTaskRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlMoveToStageDoc = Nothing
    Set xmlMoveToStageRequestNode = Nothing
    Set xmlResponse = Nothing
    Set xmlPrinterNode = Nothing
       
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'BMIDS697 End


Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

' PSC 15/09/2005 MAR32 - Start
Private Sub GetAdditionalCaseTaskData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo GetAdditionalCaseTaskDataExit

    Const cstrFunctionName As String = "GetAdditionalCaseTaskData"

    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseTaskOut As IXMLDOMNode
    Dim xmlTaskOut As IXMLDOMNode
    Dim xmlTask As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40

    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strResponse As String

    Set xmlCaseTaskIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETCASETASKS"
    strResponse = gobjMsgTmBO.TmRequest(xmlTempRequest.xml)
    Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempResponse.documentElement, vxmlResponseNode, True
    Set xmlCaseTaskOut = xmlTempResponse.selectSingleNode("RESPONSE/CASETASK")

    vxmlResponseNode.appendChild xmlCaseTaskOut

GetAdditionalCaseTaskDataExit:

    Set xmlCaseTaskIn = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTaskOut = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTaskOut = Nothing
    Set xmlTask = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub PrepareSetSLADateRequest( _
    ByVal vxmlAppNode As IXMLDOMNode, _
    ByVal vxmlNextStageNode As IXMLDOMNode, _
    ByVal vxmlGeneratedRequestDoc As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' procedure:    PrepareSetSLADateRequest
' author:       Paul Carter
' date:         22/09/2005
' description:
' pass:         vxmlAppNode  XML Application data
'               vxmlNextCaseStageNode XML Next Stage data
'               vxmlGeneratedRequestDoc XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo PrepareSetSLADateRequest_Exit

Dim xmlRequest As IXMLDOMNode
Dim xmlElem As IXMLDOMElement
Dim xmlApplicationNode As IXMLDOMNode

Dim lngAppSubStageId As Long
Dim lngNextStageId As Long

Const cstrFunctionName As String = "PrepareSetSLADateRequest"
    
    lngAppSubStageId = GetGlobalParamAmount("TMApplicationSubmissionStageID")
    lngNextStageId = xmlGetAttributeAsLong(vxmlNextStageNode, "STAGEID", "0")
    
    If lngNextStageId >= lngAppSubStageId Then
        Set xmlRequest = vxmlGeneratedRequestDoc.createElement("REQUEST")
        vxmlGeneratedRequestDoc.appendChild xmlRequest

        Set xmlApplicationNode = vxmlGeneratedRequestDoc.createElement("APPLICATION")
        xmlRequest.appendChild xmlApplicationNode
    
        Set xmlElem = vxmlGeneratedRequestDoc.createElement("APPLICATIONNUMBER")
        xmlElem.Text = xmlGetAttributeText(vxmlAppNode, "APPLICATIONNUMBER")
        xmlApplicationNode.appendChild xmlElem
    
    End If

PrepareSetSLADateRequest_Exit:
    Set xmlRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

' PSC 15/09/2005 MAR32 - End

'MAR1622 GHun Rewritten to call omTMBO.CreateAdhocCaseTask instead of MsgTMBO.CreateCaseTask
'MAR1048 Add new routine
Private Sub CreateFundsReleaseFailedTask(ByVal vxmlRequestNode As IXMLDOMNode)

    Const cstrFunctionName As String = "CreateFundReleaseFailedTask"

    Dim xmlRequest      As IXMLDOMElement
    Dim xmlDoc          As FreeThreadedDOMDocument40
    Dim xmlCaseTaskIn   As IXMLDOMElement
    Dim xmlCaseTask     As IXMLDOMElement
    Dim xmlElem         As IXMLDOMElement
    
    Dim strTaskId       As String
    Dim strResponse     As String
    Dim omTMBO          As omTMBO
    
On Error GoTo CreateFundReleaseFailedTaskExit
    
    'Get Global Parameter holding Task ID
    strTaskId = GetGlobalParamString("TMFTFundsRelFailTaskID")
    If Len(strTaskId) = 0 Then
        errThrowError cstrFunctionName, oeMissingParameter, "TMFTFundsRelFailTaskID"
    End If
    
    Set xmlDoc = xmlCreateDOMObject
    Set xmlRequest = vxmlRequestNode.cloneNode(False)
    xmlRequest.setAttribute "OPERATION", "CREATEADHOCCASETASK"
    
    Set xmlCaseTaskIn = vxmlRequestNode.selectSingleNode("//CASETASK")
    
    Set xmlCaseTask = xmlDoc.createElement("CASETASK")
    xmlCopyMandatoryAttribute xmlCaseTaskIn, xmlCaseTask, "SOURCEAPPLICATION"
    xmlCopyMandatoryAttribute xmlCaseTaskIn, xmlCaseTask, "CASEID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "CASEACTIVITYGUID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "STAGEID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "CASESTAGESEQUENCENO"
        
    xmlCaseTask.setAttribute "TASKID", strTaskId
    xmlCaseTask.setAttribute "TASKSTATUS", "10"
    
    xmlRequest.appendChild xmlCaseTask
    
    Set xmlElem = vxmlRequestNode.selectSingleNode("TASKNOTE")
    xmlCaseTask.appendChild xmlElem
    
    Set xmlElem = vxmlRequestNode.selectSingleNode("APPLICATION")
    xmlRequest.appendChild xmlElem.cloneNode(False)
    
    Set omTMBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".omTMBO")
    strResponse = omTMBO.OmTmRequest(xmlRequest.xml)
    errCheckXMLResponse strResponse, True

CreateFundReleaseFailedTaskExit:
    Set xmlRequest = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseTaskIn = Nothing
    Set xmlCaseTask = Nothing
    Set xmlElem = Nothing
    Set omTMBO = Nothing

    errCheckError cstrFunctionName

End Sub
'MAR1622 End

'MAR1848 GHun
Private Sub GetAndSynchroniseCustomerDetails(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName  As String = "GetAndSynchroniseCustomerDetails"
    Dim objCustomerBO       As Object
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument
    Dim xmlElem             As IXMLDOMElement
    
On Error GoTo GetAndSynchroniseCustomerDetailsExit
    
    Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    strResponse = objCustomerBO.GetAndSynchroniseCustomerDetails(vxmlRequestNode.xml)
    
    Set xmlDoc = xmlCreateDOMObject
    xmlDoc.loadXML (strResponse)
    
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
    For Each xmlElem In xmlDoc.documentElement.childNodes
        vxmlResponseNode.appendChild xmlElem
    Next
    
GetAndSynchroniseCustomerDetailsExit:
    Set xmlElem = Nothing
    Set objCustomerBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'MAR1848 End

'MAR1848 GHun
Private Sub GetCustomerNumbers(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName  As String = "GetCustomerNumbers"
    Dim objCustomerBO       As Object
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument
    Dim xmlElem             As IXMLDOMElement
    
On Error GoTo GetCustomerNumbersExit
    
    Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    strResponse = objCustomerBO.GetCustomerNumbers(vxmlRequestNode.xml)
    
    Set xmlDoc = xmlCreateDOMObject
    xmlDoc.loadXML (strResponse)
    
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
    For Each xmlElem In xmlDoc.documentElement.childNodes
        vxmlResponseNode.appendChild xmlElem
    Next
    
GetCustomerNumbersExit:
    Set xmlElem = Nothing
    Set objCustomerBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'MAR1848 End
