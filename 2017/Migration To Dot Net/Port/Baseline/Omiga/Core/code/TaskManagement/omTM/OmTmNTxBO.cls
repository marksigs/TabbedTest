VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "OmTmNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OmTmBOBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:   Task Manager New Transaction Business Object.
'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DRC    14/11/01    Added Methods RunCompositeRequest & TmRequest to ensure a new transaction
'                   for Moveto... methods before calling ProcessAutomaticTasks
'SA     29/11/01    SYS3270 Added new method GetApplicationData (moved from omTmBo)
'JLD    29/01/02    SYS3734 updated ForwardBORequest to work for phase2 components as well as phase1.
'PSC    15/02/02    SYS3997 Move GetApplicationData functionality to omTmCommonBO
'MV     17/01/2005  E2EM00001288 - Amended IOmTmNTxBO_MoveCaseToNewStage()
'-------------------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements IOmTmNTxBO
Private gobjContext As ObjectContext
Private gobjTmBO As Object
'----------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     14/08/2002  BMIDS00333  Core Ref AQr: SYS4919 ; Modified IOmTmNTxBO_GetApplicationData
'MDC    22/08/2002  BMIDS00336  New method RunExperianCreditCheck
'MDC    27/09/2002  BMIDS00521  Return error messages from Experian
'PSC    06/01/2003  BM0089      New Methods: MoveToStage, MoveToNextStage & ProcessAutomaticTasks
'GD     05/03/2003  BM0373      RunExperianCreditCheck, ImportBureauData - pass DATEOFBIRTH with Customer details.
'GHun   10/03/2003  BM0376      CC026 Added method RunIncomeCalcs
'MDC    24/04/2003  BM0340      Added MoveCaseToNewStage
'----------------------------------------------------------------------------------
Private Sub IOmTmNTxBO_GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                                    ByVal vxmlApplicationDestDoc As FreeThreadedDOMDocument40, _
                                    Optional ByVal vstrRBTemplateName As String)
    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "GetApplicationData"
    ' PSC 15/02/02 SYS3997 - Start
    Dim objomTmCommonBO As IomTmCommonBO
    gobjContext.SetComplete
    Set objomTmCommonBO = gobjContext.CreateInstance(App.Title & ".omTmCommonBO")
    objomTmCommonBO.GetApplicationData vxmlApplicationSrceNode, vxmlApplicationDestDoc, vstrRBTemplateName
GetApplicationDataExit:
    Set objomTmCommonBO = Nothing
    ' PSC 15/02/02 SYS3997 - End
    errCheckError cstrFunctionName
End Sub
'BMIDS00336 MDC 23/08/2002
Private Sub IOmTmNTxBO_ImportBureauData(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    IOmTmNTxBO_ImportBureauData
' author:       Mark Coleman
' date:         23/08/2002
' description:  Import data from a full bureau download
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo ImportBureauDataErr
Const cstrFunctionName As String = "ImportBureauData"
Dim objAppManagerBO As Object
Dim objAIPBO As Object
Dim objBO As Object
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlApplicationNode As IXMLDOMNode
Dim xmlCustomerListNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlCaseTask As IXMLDOMNode
Dim xmlCustomerNodeList As IXMLDOMNodeList
Dim xmlCustomerNode As IXMLDOMNode
Dim xmlNewCustomerNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim strResponse As String
Dim blnImportComplete As Boolean
    'Initialise
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    Set objAIPBO = gobjContext.CreateInstance(gstrAIP_COMPONENT & ".AgreementInPrincipleBO")
    'Get a list of customers on the current application ===========================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION" _
                    & "[(@APPLICATIONNUMBER  !=  '')  and  (@APPLICATIONFACTFINDNUMBER  !=  '')]")
    Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.appendChild xmlApplicationNode
    strResponse = objAppManagerBO.FindCustomersForApplication(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseElem, True
    Set xmlCustomerListNode = xmlResponseDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerNodeList = xmlResponseDoc.documentElement.selectNodes("CUSTOMER/CUSTOMERVERSION")
    For Each xmlCustomerNode In xmlCustomerNodeList
        Set xmlNewCustomerNode = xmlResponseDoc.createElement("CUSTOMER")
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "CUSTOMERNUMBER")
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
        xmlCustomerListNode.appendChild xmlNewCustomerNode
    Next
    'Set xmlCustomerListNode = xmlGetMandatoryNode(xmlResponseDoc.documentElement, ".//CUSTOMERLIST")
    'Import the bureau data =======================================================================
    xmlTempRequest.appendChild xmlCustomerListNode
    strResponse = objAIPBO.ImportBureauData(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True, xmlResponseElem
    'If necessary, update task status =============================================================
    If Not xmlGetAttributeAsBoolean(vxmlRequestNode, "CREDITCHECK") Then
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
        Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
        xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
        xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
        xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
        xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
        xmlTempRequest.appendChild xmlCaseTask
        Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
        strResponse = objBO.TmRequest(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
    End If
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    gobjContext.SetComplete
ImportBureauDataExit:
    Set objAppManagerBO = Nothing
    Set objAIPBO = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTask = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlCustomerNodeList = Nothing
    Set objBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
ImportBureauDataErr:
    App.LogEvent vbCrLf & "Error: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError
    'Abort this transaction
    gobjContext.SetAbort
    GoTo ImportBureauDataExit
End Sub
Private Sub IOmTmNTxBO_MoveCaseToNewStage(ByVal vxmlApplicationRequestDoc As FreeThreadedDOMDocument40, _
                                            ByVal vxmlCreateCaseStageDoc As FreeThreadedDOMDocument40, _
                                            ByVal vxmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' procedure:    MoveCaseToNewStage
' author:       Mark Coleman
' date:         24/04/2003
' description:
' pass:         vxmlApplicationRequestDoc   XML request for CreateApplicationStage
'               vxmlCreateCaseStageDoc      XML request for CreateCaseStage
'               vxmlUpdateCurrentStageDoc   XML request for UpdateCurrentStage
' return:       N/A
'------------------------------------------------------------------------------------------
    
    Const cstrFunctionName As String = "MoveCaseToNewStage"
    
    On Error GoTo MoveCaseToNewStage_Err
    
    Dim objTmBO As OmTmBO
    Dim objApplicationBO As Object
    Dim strResponse As String
    Dim xmlTaskNotesList As IXMLDOMNodeList
    Dim xmlTakNote  As IXMLDOMNode
    Dim vxmlCreateCaseTaskNoteDoc As FreeThreadedDOMDocument40
    Dim strStageSequenceNo As String
    Dim xmlElem As IXMLDOMElement
    Dim vxmlResponseDoc As FreeThreadedDOMDocument40
    Dim vxmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlTempNode As IXMLDOMNode
    
    Set vxmlCreateCaseTaskNoteDoc = New FreeThreadedDOMDocument40
    Set vxmlCurrentStageDoc = New FreeThreadedDOMDocument40
    Set vxmlResponseDoc = New FreeThreadedDOMDocument40
    
'Stage1:

    'Update the current stage (if necessary)
    If Not vxmlUpdateCurrentStageDoc Is Nothing Then
        IOmTmNTxBO_RunTmRequest vxmlUpdateCurrentStageDoc
    End If
    
'Stage2:

    'Create the new case stage
    IOmTmNTxBO_RunTmRequest vxmlCreateCaseStageDoc
    
'Stage3:
    
    'To Get the CaseStageSequenceNumber from the new stage
    'Get current CASESTAGE detail
    vxmlCurrentStageDoc.appendChild xmlGetRequestNode(vxmlCreateCaseStageDoc.documentElement)
    xmlSetAttributeValue vxmlCurrentStageDoc.documentElement, "OPERATION", "GETCURRENTSTAGE"
   
    Set xmlTempNode = vxmlUpdateCurrentStageDoc.selectSingleNode(".//CASESTAGE").cloneNode(True)
    Set xmlElem = vxmlCurrentStageDoc.createElement("CASEACTIVITY")
    xmlCopyAttribute xmlTempNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlTempNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlTempNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlTempNode, xmlElem, "ACTIVITYINSTANCE"
    
    vxmlCurrentStageDoc.documentElement.appendChild xmlElem
    
    'Call GetStageDetails
    IOmTmNTxBO_RunTmRequest vxmlCurrentStageDoc, vxmlResponseDoc
    
    strStageSequenceNo = vxmlResponseDoc.selectSingleNode(".//CASESTAGE").Attributes.getNamedItem("CASESTAGESEQUENCENO").Text
    'Create New Case Task Notes if exist any
    Set xmlTaskNotesList = vxmlCreateCaseStageDoc.selectNodes(".//TASKNOTES/TASKNOTE")

    'If there are any TaskNotesNodes then Save them
    If xmlTaskNotesList.length > 0 Then

        'Create Request Node
        vxmlCreateCaseTaskNoteDoc.appendChild xmlGetRequestNode(vxmlCreateCaseStageDoc.documentElement.cloneNode(True))

        'Set Operation Attribute to FINDTASKNOTELIST
        xmlSetAttributeValue vxmlCreateCaseTaskNoteDoc.documentElement, "OPERATION", "CREATETASKNOTE"

        'append all the taskNotes Nodes to the request element
        For Each xmlTakNote In xmlTaskNotesList
            xmlSetAttributeValue xmlTakNote, "CASESTAGESEQUENCENO", strStageSequenceNo
            vxmlCreateCaseTaskNoteDoc.documentElement.appendChild xmlTakNote.cloneNode(True)
        Next

        'Create CaseTaskNode
        IOmTmNTxBO_RunTmRequest vxmlCreateCaseTaskNoteDoc
    
    End If
    
'Stage4:

    'Call omApp.ApplicationBO to create application stage on omiga
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    strResponse = objApplicationBO.CreateApplicationStage(vxmlApplicationRequestDoc.xml)
    errCheckXMLResponse strResponse, True
    
    gobjContext.SetComplete
    
MoveCaseToNewStage_Exit:
    
    Set objTmBO = Nothing
    Set objApplicationBO = Nothing
    Set xmlTaskNotesList = Nothing
    Set xmlTakNote = Nothing
    Set vxmlCreateCaseTaskNoteDoc = Nothing
    Set xmlElem = Nothing
    Set vxmlResponseDoc = Nothing
    Set vxmlCurrentStageDoc = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
MoveCaseToNewStage_Err:
    
    gobjContext.SetAbort
    
    GoTo MoveCaseToNewStage_Exit
    
End Sub
'BM0340 MDC 24/04/2003 - End
Private Sub IOmTmNTxBO_RunExperianCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    IOmTmNTxBO_RunExperianCreditCheck
' author:       Mark Coleman
' date:         22/08/2002
' description:  Performs a Credit Score and Full Bureau Download
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo IOmTmNTxBO_RunExperianCreditCheck_VbErr
Dim strFunctionName As String
strFunctionName = "IOmTmNTxBO_RunExperianCreditCheck"
    
Dim xmlNewRequest As IXMLDOMNode
Dim xmlCaseTask As IXMLDOMElement
Dim xmlApplicationElem As IXMLDOMElement
Dim xmlElem As IXMLDOMElement
Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlCustomerNodeList As IXMLDOMNodeList
Dim xmlCustomerNode As IXMLDOMNode
Dim xmlNewCustomerNode As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlCustomerListNode As IXMLDOMNode
Dim objBO As Object
Dim objAppManagerBO As Object
Dim strResponse As String
Dim strApplicationNumber As String
Dim strAFFNumber As String
Dim lngErrNo As Long
    
    
    ' Update the Case Task status ==============================================================
    ' Build the Request and update the case task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    Set xmlCaseTask = vxmlRequestNode.selectSingleNode(".//CASETASK")
    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)
    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    'Get a list of customers on the current application ===========================================
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION" _
                    & "[(@APPLICATIONNUMBER  !=  '')  and  (@APPLICATIONFACTFINDNUMBER  !=  '')]")
    Set xmlApplicationElem = xmlMakeNodeElementBased(xmlApplicationElem, True, "APPLICATION")
    Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.appendChild xmlApplicationElem
    strResponse = objAppManagerBO.FindCustomersForApplication(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, xmlResponseDoc.documentElement, True
    Set xmlCustomerListNode = xmlResponseDoc.createElement("CUSTOMERLIST")
    Set xmlCustomerNodeList = xmlResponseDoc.documentElement.selectNodes("CUSTOMER/CUSTOMERVERSION")
    For Each xmlCustomerNode In xmlCustomerNodeList
        Set xmlNewCustomerNode = xmlResponseDoc.createElement("CUSTOMER")
        Set xmlTempNode = xmlGetMandatoryNode(xmlCustomerNode, "CUSTOMERNUMBER")
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
        xmlCustomerListNode.appendChild xmlNewCustomerNode
    Next
    ' Run the Credit Score and Full Bureau Download ============================================
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlNewRequest.appendChild xmlApplicationElem
    xmlNewRequest.appendChild xmlCustomerListNode
    Set objBO = gobjContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
    strResponse = objBO.RunCreditCheck(xmlNewRequest.xml)
    'BMIDS00521 MDC 27/09/2002
    errCheckXMLResponse strResponse, True
'    lngErrNo = errCheckXMLResponse(strResponse, False)
'    If lngErrNo <> 0 Then
'        errThrowError strFunctionName, 4812
'    End If
    'BMIDS00521 MDC 27/09/2002 - End
    gobjContext.SetComplete
IOmTmNTxBO_RunExperianCreditCheck_Exit:
    Set xmlNewRequest = Nothing
    Set xmlCaseTask = Nothing
    Set xmlApplicationElem = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    Set objBO = Nothing
    Set objAppManagerBO = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlNewCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlCustomerListNode = Nothing
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
IOmTmNTxBO_RunExperianCreditCheck_VbErr:
    App.LogEvent vbCrLf & "Error: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError
    gobjContext.SetAbort
    GoTo IOmTmNTxBO_RunExperianCreditCheck_Exit
End Sub
'BMIDS00336 MDC 23/08/2002 - End
'BM0376 GHun 10/03/2003
Private Sub IOmTmNTxBO_RunIncomeCalcs(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    IOmTmNTxBO_RunIncomeCalcs
' author:       GHun
' date:         10/03/2003
' description:  Call RunIncomeCalcsImport after a bureau download
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
'------------------------------------------------------------------------------------------
    Const cstrFunctionName As String = "RunIncomeCalcs"
    Dim objAIPBO            As Object
    Dim xmlRequest          As IXMLDOMNode
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim strResponse         As String
On Error GoTo RunIncomeCalcsErr
    Set objAIPBO = gobjContext.CreateInstance(gstrAIP_COMPONENT & ".AgreementInPrincipleBO")
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION" _
                    & "[(@APPLICATIONNUMBER  !=  '')  and  (@APPLICATIONFACTFINDNUMBER  !=  '')]")
    Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
    Set xmlRequest = vxmlRequestNode.cloneNode(False)
    xmlRequest.appendChild xmlApplicationNode
    strResponse = objAIPBO.RunIncomeCalcs(xmlRequest.xml)
    errCheckXMLResponse strResponse, False
    gobjContext.SetComplete
RunIncomeCalcsExit:
    Set objAIPBO = Nothing
    Set xmlRequest = Nothing
    Set xmlApplicationNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
RunIncomeCalcsErr:
    App.LogEvent vbCrLf & "Error: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError
    gobjContext.SetAbort
    Resume RunIncomeCalcsExit
End Sub
'BM0376 End
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Private Sub IOmTmNTxBO_ForwardBORequest(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName As String = "ForwardBORequest"
    On Error GoTo ForwardBORequestVbErr
    Dim xmlContextNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim objBO As Object
    Dim strProgId As String, _
        strMethod As String, _
        strOperation As String
    Dim varResponse As Variant
    gobjContext.SetComplete
    Set xmlContextNode = vxmlRequestNode.selectSingleNode("CRITICALDATACONTEXT")
    strProgId = xmlContextNode.Attributes.getNamedItem("COMPONENT").Text
    strMethod = xmlContextNode.Attributes.getNamedItem("METHOD").Text
    If Not xmlContextNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        strOperation = xmlContextNode.Attributes.getNamedItem("OPERATION").Text
    Else: strOperation = ""
    End If
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.async = False
    ' clone original REQUEST
    Set xmlThisRequestNode = vxmlRequestNode.cloneNode(True)
    ' set OPERATION attribute to that passed in (if any - phase2 objects and later)  - JLD SYS3734
    If Not strOperation = "" Then
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = strOperation
    Else
        xmlThisRequestNode.Attributes.removeNamedItem "OPERATION"
    End If
    ' drop CRITICALDATACONTEXT node
    xmlThisRequestNode.removeChild xmlThisRequestNode.selectSingleNode("CRITICALDATACONTEXT")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.async = False
    Set objBO = gobjContext.CreateInstance(strProgId)
    varResponse = CallByName(objBO, strMethod, VbMethod, xmlThisRequestNode.xml)
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML varResponse
        
    Set objBO = Nothing
        
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    Set xmlThisResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE").firstChild
    If Not (xmlThisResponseNode Is Nothing) Then
        vxmlResponseNode.appendChild xmlThisResponseDoc.selectSingleNode("RESPONSE").firstChild.cloneNode(True)
    End If
    Debug.Print cstrFunctionName & ": vxmlResponse is "
    Debug.Print vxmlResponseNode.xml
ForwardBORequestExit:
    
    Set xmlThisResponseDoc = Nothing
    Set xmlThisResponseNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
ForwardBORequestVbErr:
    gobjContext.SetAbort
    GoTo ForwardBORequestExit
End Sub
Private Sub IOmTmNTxBO_RunCompositeRequest(ByVal vxmlRequestNodes As Collection)
    
    Const cstrFunctionName As String = "RunCompositeRequest"
    gobjContext.SetComplete
    On Error GoTo RunCompositeRequestExit
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlInRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.async = False
    Set xmlThisRequestNode = _
        xmlThisRequestDoc.appendChild(vxmlRequestNodes.Item(1).cloneNode(False))
    xmlThisRequestNode.Attributes.removeNamedItem "OPERATION"
    For Each xmlInRequestNode In vxmlRequestNodes
        Set xmlElem = xmlThisRequestDoc.createElement("OPERATION")
        xmlElem.setAttribute _
            "NAME", _
            xmlInRequestNode.Attributes.getNamedItem("OPERATION").Text
        Set xmlNode = xmlThisRequestNode.appendChild(xmlElem)
        xmlNode.appendChild _
            xmlInRequestNode.firstChild.cloneNode(True)
    Next
        
    IOmTmNTxBO_RunTmRequest xmlThisRequestDoc
RunCompositeRequestExit:
    
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlInRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName
End Sub
Private Sub IOmTmNTxBO_RunTmRequest( _
    ByVal vxmlRequestDoc As FreeThreadedDOMDocument40, _
    Optional ByVal vxmlResponseDoc As FreeThreadedDOMDocument40 = Nothing)
    gobjContext.SetComplete
    Const cstrFunctionName As String = "RunTmRequest"
    On Error GoTo RunTmRequestExit
    Debug.Print "RunTmRequest request"
    Debug.Print vxmlRequestDoc.xml
    Set gobjTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    If vxmlResponseDoc Is Nothing Then
        Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.async = False
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.loadXML gobjTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print xmlThisResponseDoc.xml
        Set xmlThisResponseDoc = Nothing
    Else
        vxmlResponseDoc.async = False
        vxmlResponseDoc.setProperty "NewParser", True
        vxmlResponseDoc.validateOnParse = False
        vxmlResponseDoc.loadXML gobjTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode vxmlResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print vxmlResponseDoc.xml
    End If
RunTmRequestExit:
    If Err.Number <> 0 Then
        gobjContext.SetAbort
    End If
    errCheckError cstrFunctionName
End Sub
'SDS  14/12/2004  BBG1741 - Start
''PSC 06/01/2003 BM0089 - Start
'Private Function IOmTmNTxBO_MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'    Const cstrFunctionName As String = "IOmTmNTxBO_MoveToNextStage"
'    On Error GoTo MoveToNextStageVbErr
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim objTmBO As OmTmBO
'    Dim strResponse As String
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToNextStageNewTx"
'    Set objTmBO = gobjContext.CreateInstance(App.Title & ".omTmBO")
'    strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'    IOmTmNTxBO_MoveToNextStage = strResponse
'    gobjContext.SetComplete
'MoveToNextStageExit:
'
'    Set xmlTempRequest = Nothing
'    Set objTmBO = Nothing
'    errCheckError cstrFunctionName, TypeName(Me)
'    Exit Function
'MoveToNextStageVbErr:
'    gobjContext.SetAbort
'    GoTo MoveToNextStageExit
'End Function
'Private Function IOmTmNTxBO_MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'    Const cstrFunctionName As String = "IOmTmNTxBO_MoveToStage"
'    On Error GoTo MoveToStageVbErr
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim objTmBO As OmTmBO
'    Dim strResponse As String
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStageNewTx"
'    Set objTmBO = gobjContext.CreateInstance(App.Title & ".omTmBO")
'    strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'    IOmTmNTxBO_MoveToStage = strResponse
'    gobjContext.SetComplete
'MoveToStageExit:
'
'    Set xmlTempRequest = Nothing
'    Set objTmBO = Nothing
'    errCheckError cstrFunctionName, TypeName(Me)
'    Exit Function
'MoveToStageVbErr:
'    gobjContext.SetAbort
'    GoTo MoveToStageExit
'End Function
'SDS  14/12/2004  BBG1741 - End

Private Sub IOmTmNTxBO_ProcessAutomaticTasks(ByVal vxmlRequest As IXMLDOMNode)
    Const cstrFunctionName As String = "ProcessAutomaticTasks"
    On Error GoTo ProcessAutomaticTasksVbErr
        
    Dim objOmTmBo As Object
    Dim strResponse As String
    Set objOmTmBo = gobjContext.CreateInstance(App.Title & ".OmTmBo")
    'Call into OmTMBo.ProcessAutomaticTasks
    strResponse = objOmTmBo.OmTmRequest(vxmlRequest.xml)
    errCheckXMLResponse strResponse, True
    gobjContext.SetComplete
ProcessAutomaticTasksExit:
    
    errCheckError cstrFunctionName
    Exit Sub
ProcessAutomaticTasksVbErr:
    
    gobjContext.SetAbort
    GoTo ProcessAutomaticTasksExit
End Sub
'PSC 06/01/2003 BM0089 - End
