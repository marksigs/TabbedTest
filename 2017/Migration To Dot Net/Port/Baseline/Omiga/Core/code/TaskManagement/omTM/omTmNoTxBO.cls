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
'GD     09/07/2003  BM0340 Amended AddDefaultValuesToCaseTask to look at vxmlTaskNode, not vxmlCaseTaskNode for Task Owner info.
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------
'BBG History:
'
'Prog   Date        Description
'MV     05/09/2004  E2EM00001206 - Amended MoveToNextStage()
'MV     19/10/2004  BBG1658 - Amended MoveToNextStage(); MoveToStage()
'TK     20/10/2004  E2EM00002505 - Amended RunUserRules
'PSC    09/12/2004  E2EM00003099 - Amend ProcessCarryForwardTasks to copy correct attributes and
'SDS    10/12/2004  BBG1741 - Task Management - Move Stage is taking a long time when performed against a large DB
'MV     17/01/2005  E2EM00001288 - Amended MoveToStage, MoveToNextStage, ProcessCarryForwardTasks
'MV     27/01/2005  E2EM00003322 - Amended ProcessCarryForwardTasks to Include TaskName and CaseTaskName
'MV     14/02/2005  BBG1931 - Amended MoveToNextStage()
'HMA    23/02/2005  E2EM00001855 - Added GetLatestRiskAssessment.
'AS     25/07/2005  CORE174 omTM cannot create an ad hoc task for anyone except the logged in user.
'                           AddDefaultValuesToCaseTask moved to omTMFunctions.bas.
'                           SetCaseTaskDueDateAndTime moved to omTMFunctions.bas.
'------------------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl
Private gobjContext As ObjectContext
Private gobjMsgTmBO As Object
Private Const cstrTaskComplete = 40
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
Dim strAttribValue As String
Const cstrFunctionName As String = "OmTmNoTxRequest"
    Set gobjMsgTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    If gobjMsgTmBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, "MsgTm.MsgTmBO"
    End If
    Set xmlInDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlInDoc.async = False
    xmlOutDoc.async = False
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.validateOnParse = False
    xmlInDoc.loadXML vstrXmlIn
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    xmlCheckMandatoryAttribute xmlRequestNode, "OPERATION"
    strOperation = xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    strOperation = UCase(strOperation)
    Select Case strOperation
            
        Case "PROCESSAUTOMATICTASKS"
            ProcessAutomaticTasks xmlRequestNode
        Case "MOVETONEXTSTAGE"
            MoveToNextStage xmlRequestNode, xmlResponseNode
        Case "MOVETOSTAGE"
            MoveToStage xmlRequestNode, xmlResponseNode
        'SDS  BBG1741  10/12/2004__START    Moved here from omTmBO
        Case "DECLINEAPPLICATION"
            DeclineApplication xmlRequestNode
        'SDS  BBG1741  10/12/2004__END
        ' E2EM00001855  add new operation
        Case "GETLATESTRISKASSESSMENT"
            GetLatestRiskAssessment xmlRequestNode, xmlResponseNode
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
Private Sub ProcessAutomaticTasks(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo ProcessAutomaticTasksExit
    
Dim objNTxBO As IOmTmNTxBO
Dim xmlAutoRequestDoc As FreeThreadedDOMDocument40
Dim xmlAutoRequestNode As IXMLDOMNode
Dim xmlAutoCaseTaskNode As IXMLDOMNode
Dim xmlCurrentStageNode As IXMLDOMNode
Dim xmlApplicationNode As IXMLDOMNode
Dim xmlPrinterNode As IXMLDOMNode
Dim xmlCaseActivityNode As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlNodeList As IXMLDOMNodeList
Dim xmlTempNode As IXMLDOMNode
Dim sTempTaskName As String
Dim iAutoErrorCount As Integer
Dim strResponse As String
Const cstrFunctionName As String = "ProcessAutomaticTasks"
    
    
    
    '==================================================================================================
    'Initialise
    '==================================================================================================
    'Validate Request
    '   CURRENTSTAGE for 'MoveToNextStage'
    '   CASESTAGE for 'MoveToStage'
    '   CASEACTIVITY for 'ReinstatePreviousStage'
    Set xmlCaseActivityNode = xmlGetNode(vxmlRequestNode, "CASEACTIVITY")
    If xmlCaseActivityNode Is Nothing Then
        Set xmlTempNode = xmlGetNode(vxmlRequestNode, "CURRENTSTAGE")
        If xmlTempNode Is Nothing Then
            Set xmlTempNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
            xmlChangeNodeName xmlTempNode, "CASESTAGE", "CASEACTIVITY"
        Else
            xmlChangeNodeName xmlTempNode, "CURRENTSTAGE", "CASEACTIVITY"
        End If
        Set xmlCaseActivityNode = xmlTempNode
    End If
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseActivityNode, "ACTIVITYINSTANCE"
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
    Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("./PRINTER")
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GetCurrentStage"
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    '==================================================================================================
    'Get the current stage
    '==================================================================================================
    strResponse = gobjMsgTmBO.TmRequest(vxmlRequestNode.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    Set xmlCurrentStageNode = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "CASESTAGE")
    '==================================================================================================
    'Process each automatic task (in its own transaction)
    '==================================================================================================
    Set xmlNodeList = xmlCurrentStageNode.selectNodes("CASETASK[@AUTOMATICTASKIND='1']")
    sTempTaskName = ""
    On Error Resume Next
    For Each xmlAutoCaseTaskNode In xmlNodeList
        Set xmlAutoRequestDoc = New FreeThreadedDOMDocument40
        xmlAutoRequestDoc.async = False
        Set xmlAutoRequestNode = xmlAutoRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlAutoRequestNode.Attributes.getNamedItem("OPERATION").Text = "ProcessAutomaticTasks"
        xmlAutoRequestNode.appendChild xmlAutoCaseTaskNode
        xmlAutoRequestNode.appendChild xmlApplicationNode
        xmlAutoRequestNode.appendChild xmlPrinterNode
                
        objNTxBO.ProcessAutomaticTasks xmlAutoRequestNode
        If Err.Number <> 0 Then
            Err.Clear
            iAutoErrorCount = iAutoErrorCount + 1
            sTempTaskName = sTempTaskName & xmlAutoCaseTaskNode.Attributes.getNamedItem("TASKID").xml
        End If
    Next
    On Error GoTo ProcessAutomaticTasksExit
    If iAutoErrorCount > 0 Then
        errThrowError cstrFunctionName, oeTmAutomaticTaskDidntComplete, "", iAutoErrorCount, sTempTaskName
    End If
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
    errCheckError cstrFunctionName, TypeName(Me)
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
    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
    Dim xmlRbResponseNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlNextCaseStageNode As IXMLDOMNode
    Dim xmlRequestCaseStageNode As IXMLDOMNode
    Dim xmlComboDoc As FreeThreadedDOMDocument40
    Dim xmlComboNode As IXMLDOMNode
    Dim objNTxBO As IOmTmNTxBO
    Dim objCommonBO As IomTmCommonBO
    ' validate request
    Set xmlRequestCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
    ' get requested CASESTAGE detail =====================================================
    Set xmlNextStageDoc = New FreeThreadedDOMDocument40
    GetRequestedStage vxmlRequestNode, xmlNextStageDoc
    Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
    If xmlNextStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoNextStage
    End If
    ' remove non DEFAULT tasks
    For Each xmlNode In xmlNextStageNode.childNodes
        If xmlGetAttributeAsBoolean(xmlNode, "DEFAULTFLAG") = False Then
            xmlNextStageNode.removeChild xmlNode
        End If
    Next
    ' get current CASESTAGE detail =====================================================
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
        If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
            errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
        End If
    End If
    If AnyRulesForStage(xmlNextStageNode) = True Then
        ' get APPLICATION data =============================================================
        Set xmlRbResponseNode = GetApplicationData(xmlApplicationRequestNode, "TMRBTemplate")
        'Get required combos
        Set xmlComboDoc = GetComboList()
        For Each xmlComboNode In xmlComboDoc.selectNodes(".//COMBO")
            xmlRbResponseNode.appendChild xmlComboNode
        Next
        ' Run rules
        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
        xmlNextCaseStageDoc.async = False
        RunUserRules _
            vxmlRequestNode, _
            xmlNextStageNode, _
            xmlRbResponseNode, _
            xmlNextCaseStageDoc
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    Else
        ' create CASESTAGE from STAGE
        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
        xmlNextCaseStageDoc.async = False
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
    xmlCreateCaseStageDoc.async = False
    PrepareCreateCaseStageRequest _
        vxmlRequestNode, _
        xmlCurrentStageNode, _
        xmlNextCaseStageNode, _
        xmlCreateCaseStageDoc
    AddDefaultValuesToCaseStage _
        vxmlRequestNode, _
        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
        xmlNextStageNode
    ' prepare for update to current stage ==============================================
    If Not xmlCurrentStageNode Is Nothing Then
        ' first check we are auth'd to create the next stage SYS1788 JLD
        Dim xmlValidateDoc As FreeThreadedDOMDocument40
        Set xmlValidateDoc = New FreeThreadedDOMDocument40
        Dim xmlElem As IXMLDOMElement
        ' clone initial request node
        Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation name
        xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
        Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
        Set xmlNode = xmlNode.appendChild(xmlElem)
        ' copy STAGEID from next case stage
        xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
        RunTmRequest xmlValidateDoc
        Set xmlUpdateCurrentStageDoc = New FreeThreadedDOMDocument40
        xmlUpdateCurrentStageDoc.async = False
        PrepareCurrentStageUpdateRequest _
            vxmlRequestNode, _
            xmlCurrentStageNode, _
            xmlUpdateCurrentStageDoc
        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False _
        Then
            ProcessCarryForwardTasks _
                xmlCurrentStageNode, _
                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                xmlGetRequestNode(xmlCreateCaseStageDoc.documentElement)
        Else
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
'    Set xmlCurrentStageDoc = New FreeThreadedDomDocument40
'    xmlCurrentStageDoc.async = False
'    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
'
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    vxmlResponseNode.appendChild xmlCurrentStageNode
    Dim xmlApplicationRequestDoc As FreeThreadedDOMDocument40
    Set xmlApplicationRequestDoc = New FreeThreadedDOMDocument40
    'BM0340 GHun 13/06/2003 Need to pass in xmlNextStageNode, NOT xmlNextCaseStageNode
    'PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextCaseStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    'BM0340 End
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc
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
    Set xmlApplicationDataDoc = Nothing
    Set xmlValidateDoc = Nothing
    Set xmlElem = Nothing
    Set xmlRbResponseNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
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
    Dim xmlExceptionStageDoc As New FreeThreadedDOMDocument40
    Dim xmlTaskNode As IXMLDOMNode
    Dim xmlValidateDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplicationRequestDoc As FreeThreadedDOMDocument40
                    
    Dim objNTxBO As IOmTmNTxBO
    Dim strStageID As String
    Dim strStageSequenceNo As String
    Dim strStageName As String
    Dim strStageRuleReference As String
    Dim strTaskId As String
    Dim strTaskName As String
    
    Dim xmlStageTask As IXMLDOMNode
            
    Set xmlValidateDoc = New FreeThreadedDOMDocument40
    Set xmlApplicationRequestDoc = New FreeThreadedDOMDocument40
    
    ' validate request
    Set xmlCurrentStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTSTAGE")
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
    
    ' Get current CASESTAGE detail
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
        errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
    End If
        
    ' Get next CASESTAGE detail
    Set xmlNextStageDoc = New FreeThreadedDOMDocument40
    xmlNextStageDoc.async = False
    Set xmlNextCaseStageNode = Nothing
    Do While xmlNextCaseStageNode Is Nothing
            
        ' Get next STAGE in sequence
        GetNextStageInSequence vxmlRequestNode, xmlCurrentStageNode, xmlNextStageNode, xmlNextStageDoc
        Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
        
        'If the StageRuleReference is 'AAAA' then update the CurrentStageNode with NextStageNode Details and
        'Move to the next stage in the sequence.
        If xmlGetAttributeText(xmlNextStageNode, "STAGERULEREFERENCE") = "AAAA" Then
            
            strStageID = xmlGetAttributeText(xmlNextStageNode, "STAGEID")
            strStageSequenceNo = xmlGetAttributeText(xmlNextStageNode, "STAGESEQUENCENO")
            strStageSequenceNo = xmlGetAttributeText(xmlNextStageNode, "STAGESEQUENCENO")
            strStageName = xmlGetAttributeText(xmlNextStageNode, "STAGENAME")
            strStageRuleReference = xmlGetAttributeText(xmlNextStageNode, "STAGERULEREFERENCE")
            
            Set xmlStageTask = xmlNextStageNode.selectSingleNode("STAGETASK")
            
            strTaskId = xmlGetAttributeText(xmlStageTask, "TASKID")
            strTaskName = xmlGetAttributeText(xmlStageTask, "TASKNAME")
            
            xmlSetAttributeValue xmlCurrentStageNode, "STAGEID", strStageID
            xmlSetAttributeValue xmlCurrentStageNode, "STAGESEQUENCENO", strStageSequenceNo
            xmlSetAttributeValue xmlCurrentStageNode, "STAGENAME", strStageName
            xmlSetAttributeValue xmlCurrentStageNode, "STAGERULEREFERENCE", strStageRuleReference
            xmlSetAttributeValue xmlCurrentStageNode, "TASKID", strTaskId
            xmlSetAttributeValue xmlCurrentStageNode, "TASKNAME", strTaskName
            
            GetNextStageInSequence vxmlRequestNode, xmlCurrentStageNode, xmlNextStageNode, xmlNextStageDoc
            Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
            
        End If
        
        If AnyRulesForStage(xmlNextStageNode) = True Then
            If xmlRbResponseNode Is Nothing Then
                
                ' get APPLICATION data
                Set xmlRbResponseNode = GetApplicationData(xmlApplicationRequestNode, "TMRBTemplate")
                
                'Get required combos
                Set xmlComboDoc = GetComboList()
                For Each xmlComboNode In xmlComboDoc.selectNodes(".//COMBO")
                    xmlRbResponseNode.appendChild xmlComboNode
                Next
                
            End If
            
            ' Run Rules
            If xmlNextCaseStageDoc Is Nothing Then
                Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                xmlNextCaseStageDoc.async = False
            Else
                xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
            End If
            
            RunUserRules vxmlRequestNode, xmlNextStageNode, xmlRbResponseNode, xmlNextCaseStageDoc
            
            Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
            
            'If the case has been auto-declined/cancelled then the exception stageid and exception reason
            'attributes will be returned in the CASESTAGE node of xmlNextCaseStageDoc
            If Not xmlNextCaseStageNode Is Nothing Then
                
                'Next stage is set. If it is an exception stage then change the CURRENTSTAGE node to CASESTAGE
                'and copy the STAGEID and EXCEPTIONREASON attributes on the requestNode, re-populate
                'xmlNextStageNode and xmlNextCaseStageNode with the exception stage/task details
                'Is it an exception stage
                If xmlAttributeValueExists(xmlNextCaseStageNode, "EXCEPTIONREASON") Then
                    
                    'Rename the CURRENTSTAGE node on the RequestNode to CASESTAGE and copy the
                    'STAGEID and EXCEPTIONREASON attributes
                    xmlSetAttributeValue vxmlRequestNode.selectSingleNode("CURRENTSTAGE"), "STAGEID", xmlGetAttributeText(xmlNextCaseStageNode, "STAGEID")
                    xmlSetAttributeValue vxmlRequestNode.selectSingleNode("CURRENTSTAGE"), "EXCEPTIONREASON", xmlGetAttributeText(xmlNextCaseStageNode, "EXCEPTIONREASON")
                    xmlChangeNodeName vxmlRequestNode, "CURRENTSTAGE", "CASESTAGE"
                    
                    'Get exception stage/task details
                    Set xmlExceptionStageDoc = New FreeThreadedDOMDocument40
                    
                    GetRequestedStage vxmlRequestNode, xmlExceptionStageDoc
                    
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
                    For Each xmlTaskNode In xmlNextStageNode.childNodes
                        If xmlGetAttributeAsBoolean(xmlTaskNode, "DEFAULTFLAG") = False Then
                            xmlNextStageNode.removeChild xmlTaskNode
                        End If
                    Next
                    
                    'Create CaseStage/CaseTask details
                    If xmlNextCaseStageDoc Is Nothing Then
                        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                        xmlNextCaseStageDoc.async = False
                    Else
                        xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
                    End If
                    
                    'CreateCaseStageFromStage xmlNextStageNode, xmlNextCaseStageDoc
                    
                    Set xmlNextCaseStageNode = Nothing
                    
                    RunUserRules vxmlRequestNode, xmlNextStageNode, xmlRbResponseNode, xmlNextCaseStageDoc
                    
                    Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode(".//CASESTAGE")
                    
                End If
            End If
            'BS BM0136 End 07/02/2003
        Else
            If xmlNextCaseStageDoc Is Nothing Then
                Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
                xmlNextCaseStageDoc.async = False
            Else
                xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
            End If
            CreateCaseStageFromStage xmlNextStageNode, xmlNextCaseStageDoc
            Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
        End If
    Loop
    
    ' find next CASESTAGE in rules response
    If xmlNextCaseStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoNextStage
    End If
    
    ' prepare CreateCaseStage request
    Set xmlCreateCaseStageDoc = New FreeThreadedDOMDocument40
    PrepareCreateCaseStageRequest vxmlRequestNode, xmlCurrentStageNode, xmlNextCaseStageNode, xmlCreateCaseStageDoc
    
    AddDefaultValuesToCaseStage _
        vxmlRequestNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlNextStageNode
    
    ' prepare for update to current stage
    ' add incomplete tasks from current stage to next stage
    If Not xmlCurrentStageNode Is Nothing Then
        
        'BS BM0136 07/02/2003
        'Do not check authority level if case auto-declined/cancelled
        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
            ' first check we are auth'd to create the next stage   SYS1788 JLD
            
            ' clone initial request node
            Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
            
            ' set operation name
            xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
            Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
            Set xmlNode = xmlNode.appendChild(xmlElem)
            
            'Copy STAGEID from next case stage
            xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
            RunTmRequest xmlValidateDoc
            
        End If
        'BS BM0136 End 07/02/2003
          
        Set xmlUpdateCurrentStageDoc = New FreeThreadedDOMDocument40
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
            ProcessCarryForwardTasks _
                        xmlCurrentStageNode, _
                        xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
                        xmlGetRequestNode(xmlCreateCaseStageDoc.documentElement)
                    
            'ProcessCarryForwardTasks _
            '            xmlCurrentStageNode, _
            '            xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
            '            xmlCreateCaseStageDoc.documentElement
            '            'xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
                        
        Else
            ProcessCancelTasks _
                        xmlCurrentStageNode, _
                        xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE")
        End If
        'BS BM0136 End 07/02/2003
            
    End If
    ' PSC 27/11/01 SYS3221 - End
    
    'BM0340 GHun 13/06/2003 Need to pass in xmlNextStageNode, NOT xmlNextCaseStageNode
    'PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextCaseStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    PrepareCreateApplicationStageRequest xmlApplicationRequestNode, xmlNextStageNode, xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), xmlApplicationRequestDoc
    'BM0340 End
    
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    
    objNTxBO.MoveCaseToNewStage xmlApplicationRequestDoc, xmlCreateCaseStageDoc, xmlUpdateCurrentStageDoc
             
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
    Dim intStatus As Integer
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
    xmlRequestDoc.async = False
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
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
    vxmlNextCaseStageDoc.setProperty "NewParser", True
    vxmlNextCaseStageDoc.validateOnParse = False
    'TK 19/10/2004 E2EM00002505
    Dim xmlCaseStage As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute

    Set xmlCaseStage = vxmlRequestNode.selectSingleNode("//REQUEST/CASESTAGE")
    If Not xmlCaseStage Is Nothing Then
        If xmlAttributeValueExists(xmlCaseStage, "EXCEPTIONREASON") Then
            Set xmlAttrib = xmlRequestNode.selectSingleNode("//REQUEST/STAGE").ownerDocument.createAttribute("EXCEPTIONREASON")
            xmlAttrib.Text = xmlCaseStage.Attributes.getNamedItem("EXCEPTIONREASON").Text
            xmlRequestNode.selectSingleNode("//REQUEST/STAGE").Attributes.setNamedItem xmlAttrib
        End If
    End If
    'TK 19/10/2004 E2EM00002505 End
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
    Set xmlCaseStage = Nothing
    Set xmlAttrib = Nothing
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
Private Sub AddDefaultValuesToCaseStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
    ByVal vxmlNextStageNode As IXMLDOMNode)
    On Error GoTo AddDefaultValuesToCaseStageExit
    Const cstrFunctionName As String = "AddDefaultValuesToCaseStage"
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlTaskNode As IXMLDOMNode
    Dim strPattern As String
    Dim strUserId As String
    Dim strUnitId As String
    Call GetApplicationOwners(vxmlRequestNode, strUserId, strUnitId)
    For Each xmlCaseTaskNode In vxmlNextCaseStageNode.childNodes
        strPattern = _
            "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlCaseTaskNode, "TASKID") & "']"
        Set xmlTaskNode = vxmlNextStageNode.selectSingleNode(strPattern)
        'CORE174 AS 25/07/2005 AddDefaultValuesToCaseTask moved to omTMFunctions.bas; pass gobjContext.
        AddDefaultValuesToCaseTask gobjContext, vxmlRequestNode, xmlCaseTaskNode, xmlTaskNode, strUserId, strUnitId
        ' DRC AQR SYS2266 This Call Moved
        '  ProcessAutomaticTasks vxmlRequestNode, xmlCaseTaskNode
               
        ' AQR SYS1791
        ' add ORIGINATINGSTAGEID
        xmlCopyAttributeValue xmlTaskNode, xmlCaseTaskNode, "STAGEID", "ORIGINATINGSTAGEID"
                                
    Next
AddDefaultValuesToCaseStageExit:
    
    Set xmlCaseTaskNode = Nothing
    Set xmlTaskNode = Nothing
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
        xmlThisResponseDoc.async = False
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.loadXML gobjMsgTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print xmlThisResponseDoc.xml
        Set xmlThisResponseDoc = Nothing
    Else
        vxmlResponseDoc.async = False
        vxmlResponseDoc.setProperty "NewParser", True
        vxmlResponseDoc.validateOnParse = False
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
Private Sub ProcessCarryForwardTasks(ByVal vxmlCurrentStageNode As IXMLDOMNode, _
                                     ByVal vxmlCurrentStageUpdateNode As IXMLDOMNode, _
                                     ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
                                     ByVal xmlRequestNode As IXMLDOMNode)
                                     
    On Error GoTo ProcessCarryForwardTasksExit
    
    Const cstrFunctionName As String = "ProcessCarryForwardTasks"
    
    Dim xmlSrceNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim strPattern As String
    Dim enumUpdateStatus As TASKSTATUS
    Dim xmlNewStageAllTasksDoc As FreeThreadedDOMDocument40
    Dim xmlNewStageAllTasks As IXMLDOMNode
    Dim xmlStageTask As IXMLDOMNode
    Dim xmlTaskNotes As IXMLDOMNode
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim vxmlResponseDoc As FreeThreadedDOMDocument40
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    Dim xmlTempTaskNoteNodeList As IXMLDOMNodeList
    Dim xmlTempTaskNode As IXMLDOMNode
    Dim xmlNewStageTaskNoteDoc As FreeThreadedDOMDocument40
    
    Set xmlNewStageTaskNoteDoc = New FreeThreadedDOMDocument40
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    Set vxmlResponseDoc = New FreeThreadedDOMDocument40
    
    'Get all the tasks for the new stage we are moving too.
    'Create GetStageDetail request
    Set xmlNewStageAllTasksDoc = New FreeThreadedDOMDocument40
    Set xmlTaskNotes = xmlOutDoc.createElement("TASKNOTES")
    
    'clone the request
    'xmlNewStageAllTasksDoc.appendChild vxmlCurrentStageUpdateNode.ownerDocument.documentElement.cloneNode(False)
    xmlNewStageAllTasksDoc.appendChild xmlRequestNode.cloneNode(True)
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
    
    'loop through the tasks
    For Each xmlSrceNode In vxmlCurrentStageNode.childNodes
        enumUpdateStatus = omiga4TASKSTATUSUndefined
        If IsIncompleteTask(xmlSrceNode) = True Then
            
            If xmlGetAttributeAsBoolean(xmlSrceNode, "CARRYFORWARDINDICATOR") = True Then
                enumUpdateStatus = omiga4TASKSTATUSCarriedForward
            Else
                enumUpdateStatus = omiga4TASKSTATUSNotApplicable
            End If
            
            Set xmlElem = vxmlCurrentStageUpdateNode.ownerDocument.createElement("CASETASK")
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
                
                'MV Start
                'Get all the TaskNotes Associated with the tasks in the CurrentStage
                Set xmlNewStageTaskNoteDoc = New FreeThreadedDOMDocument40
                Set vxmlResponseDoc = New FreeThreadedDOMDocument40
    
                xmlNewStageTaskNoteDoc.appendChild xmlRequestNode.cloneNode(True)
                xmlSetAttributeValue xmlNewStageTaskNoteDoc.documentElement, "OPERATION", "FINDTASKNOTELIST"
                
                Set xmlElem = xmlOutDoc.createElement("TASKNOTE")
                xmlNewStageTaskNoteDoc.documentElement.appendChild xmlElem
                xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
                xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASEID"
                xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYID"
                xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
                xmlCopyAttribute xmlSrceNode, xmlElem, "STAGEID"
                xmlCopyAttribute xmlSrceNode, xmlElem, "CASESTAGESEQUENCENO"
                xmlCopyAttribute xmlSrceNode, xmlElem, "TASKID"
                xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "TASKINSTANCE"
                xmlCopyAttribute xmlSrceNode, xmlElem, "CASEACTIVITYGUID"
                
                On Error Resume Next
                RunTmRequest xmlNewStageTaskNoteDoc, vxmlResponseDoc
                
                lngErrNo = Err.Number
                strSource = Err.Source
                strDescription = Err.Description
            
                On Error GoTo ProcessCarryForwardTasksExit
   
                If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
                    Err.Raise lngErrNo, strSource, strDescription
                End If
                
                Set xmlTempTaskNoteNodeList = vxmlResponseDoc.selectNodes(".//TASKNOTE")
                If xmlTempTaskNoteNodeList.length > 0 Then
                    For Each xmlTempTaskNode In xmlTempTaskNoteNodeList
                        xmlSetAttributeValue xmlTempTaskNode, "STAGEID", xmlGetAttributeText(vxmlNextCaseStageNode, "STAGEID")
                        xmlTaskNotes.appendChild xmlTempTaskNode
                    Next
                End If
                'End
                Set xmlNewStageTaskNoteDoc = Nothing
                Set vxmlResponseDoc = Nothing
                Set xmlElem = Nothing
                
                strPattern = "CASETASK[@TASKID='" & xmlGetAttributeText(xmlSrceNode, "TASKID") & "'"
                
                If xmlAttributeValueExists(xmlSrceNode, "CUSTOMERIDENTIFIER") Then
                    
                    strPattern = strPattern & "  and  @CUSTOMERIDENTIFIER='" & _
                                 xmlGetAttributeText(xmlSrceNode, "CUSTOMERIDENTIFIER") & "'"
                End If
                                    
                If xmlAttributeValueExists(xmlSrceNode, "CONTEXT") Then
                    strPattern = strPattern & "  and  @CONTEXT='" & _
                                 xmlGetAttributeText(xmlSrceNode, "CONTEXT") & "'"
                End If
                
                strPattern = strPattern & "]"
                                   
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
                        xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKNAME", "TASKNAME"
                        xmlCopyAttribute xmlSrceNode, xmlNode, "CASETASKNAME"
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
                        
                        'PSC 09/12/2004 E2EM00003099 - Start
                        xmlCopyAttribute xmlSrceNode, xmlNode, "TASKSTATUS"
                        xmlCopyAttribute xmlSrceNode, xmlNode, "CONTEXT"
                        xmlCopyAttribute xmlSrceNode, xmlNode, "CUSTOMERIDENTIFIER"
                        'PSC 09/12/2004 E2EM00003099 - End
                        
                        vxmlNextCaseStageNode.appendChild xmlNode
                    End If
                End If
            End If
        End If
    Next
    
    If xmlTaskNotes.selectNodes("TASKNOTE").length > 0 Then
        vxmlNextCaseStageNode.appendChild xmlTaskNotes
    End If
        
ProcessCarryForwardTasksExit:

    Set xmlSrceNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlNewStageAllTasksDoc = Nothing
    Set xmlNewStageAllTasks = Nothing
    Set xmlStageTask = Nothing
    Set xmlTaskNotes = Nothing
    Set xmlOutDoc = Nothing
    Set vxmlResponseDoc = Nothing
    Set xmlTempTaskNoteNodeList = Nothing
    Set xmlTempTaskNode = Nothing
    Set xmlNewStageTaskNoteDoc = Nothing
    
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
        intStatus = omiga4TASKSTATUSPending _
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
    xmlFindAppOwnerShipListDoc.setProperty "NewParser", True
    xmlFindAppOwnerShipListDoc.validateOnParse = False
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
    Dim strPattern As String
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
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
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
    Set xmlNode = xmlDoc.createElement("COMBOLIST")
    xmlDoc.appendChild xmlNode
    'Read in combo tags from TMRBTemplateCombos ==============================================================
    strFileName = App.Path & "\TMRBTEMPLATECOMBOS.XML"
    strFileName = Replace(strFileName, "DLL", "XML", 1, 1, vbTextCompare)
    Set objFSO = New FileSystemObject
    If objFSO.FileExists(strFileName) Then
        Set xmlTemplateDoc = New FreeThreadedDOMDocument40
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
'SDS  BBG1741  10/12/2004__START    Moved here from omTmBO
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
    
    ' Validate Request
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    Set xmlCaseTaskRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
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
    Set xmlUpdateCaseTaskDoc = xmlCreateDOMObject
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
    Set xmlMoveToStageDoc = xmlCreateDOMObject
    Set xmlMoveToStageRequestNode = _
        xmlMoveToStageDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlMoveToStageRequestNode.Attributes.getNamedItem("OPERATION").Text = "MoveToStage"
    xmlMoveToStageRequestNode.appendChild xmlCaseStageRequestNode.cloneNode(True)
    xmlMoveToStageRequestNode.appendChild xmlApplicationRequestNode.cloneNode(True)
        
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
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'E2EM00001855  Add new function
Private Sub GetLatestRiskAssessment(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)
                                    
    ' This function gets the latest Product Eligibility Risk Assessment
    ' Returns <RESPONSE>
    '             <RISKASSESSMENT @RAEXISTS   @UNDERWRITERDECISION />
    '         </RESPONSE>
    '
                              
                                    
    On Error GoTo GetLatestRiskAssessmentExit
    Const cstrFunctionName As String = "GetLatestRiskAssessment"
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode

    Dim xmlRiskNodes As IXMLDOMNodeList
    Dim xmlRiskNode As IXMLDOMNode
    
    Dim intRAExists As Integer
    Dim intUnderwriterDecision As Integer
    Dim dteLastRisk As Date
    
    ' validate request
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
    
    ' get APPLICATION data =============================================================
    Set xmlApplicationNode = GetApplicationData(xmlApplicationRequestNode, "TMRBTemplate")
       
    ' Find the latest Product eligibility Risk Assessment (Type = 10) node
    Set xmlRiskNodes = xmlApplicationNode.selectNodes("//RISKASSESSMENT[@RISKASSESSMENTTYPE='10']")
            
    intUnderwriterDecision = 0
    intRAExists = 0
            
    If xmlRiskNodes.length > 0 Then
    
        intRAExists = 1

        ' Save the Underwriter Decision for the latest risk assessment

        For Each xmlRiskNode In xmlRiskNodes
            If xmlGetAttributeAsDate(xmlRiskNode, "RISKASSESSMENTDATETIME") > dteLastRisk Then
                dteLastRisk = xmlGetAttributeAsDate(xmlRiskNode, "RISKASSESSMENTDATETIME")
                intUnderwriterDecision = xmlGetAttributeAsInteger(xmlRiskNode, "UNDERWRITERDECISION")
            End If
        Next
    
    End If
      
    Set xmlElem = vxmlResponseNode.ownerDocument.createElement("RISKASSESSMENT")
    xmlElem.setAttribute "RAEXISTS", intRAExists
    xmlElem.setAttribute "UNDERWRITERDECISION", intUnderwriterDecision
    Set xmlNode = vxmlResponseNode.appendChild(xmlElem)
  
 
GetLatestRiskAssessmentExit:
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlRiskNodes = Nothing
    Set xmlRiskNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Sub
'SDS  BBG1741  10/12/2004__END
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
