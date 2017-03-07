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
'Workfile:      OmTmNTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Task Manager New Transaction Business Object.

'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    25/08/2005  MAR32 Task Automation Changes
'PSC    29/11/2005  MAR491 Amend PerformTASMoveToNextStageProcessing correctly identify
'                   whether case has already moved stage
'GHun   17/01/2005  MAR972 Changed CreateAdhocCaseTask to also be able to create TaskNotes
'AS     22/03/2006  MAR1501 OmTMBO does not set transaction to complete.
'HMA    06/04/2006  MAR1408 Set task due date based on completion date for post completion tasks.
'GHun   24/04/2006  MAR1547 Changed CreateAdhocCaseTask to set the due dateTime for delay completion
'PSC    02/05/2006  MAR1692 Amend CanMoveStage to allow move stage if task is carry forward
'JD     07/05/2006  MAR1703 MoveCaseToNewStage - add reserveproduct functionality.
'PSC    09/05/2006  MAR1643 Amend PerformTASMoveToNextStageProcessing and PerformTASTaskProcessing so that
'                   locking and unlocking the application is done in separate transactions
'HMA    16/05/2006  MAR1787 Set up mandatory flag in CreateAdHocCaseTask.
'PSC    03/08/2006  MAR1927 Don't process TAS Task if stage is different
'PSC    19/02/2007  EP2_1409 Amend ForwardBORequest to take into account warnings
'-------------------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IOmTmNTxBO

Private gobjContext As ObjectContext
Private gobjTmBO As Object

Private Enum APPLICATIONLOCKSTATUS
    lsLocked
    lsUnlocked
    lsError
End Enum

Private Enum TASPROGRESS
    ValidateRequest
    CheckingCurrentStage
    GettingTaskDetails
    GettingLockStatus
    LockingApplication
    ProcessingTask
    MovingStage
    UpdatingTaskStatus
    DetermineMoveStage
    SendMoveStageMessage
    UnLockingApplication
End Enum

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
'GHun   26/11/2003  BMIDS658    Amend ImportBureauData and RunExperianCreditCheck to reduce locking
'GHun   22/01/2004  BMIDS697    Removed old MoveTo(Next)Stage methods
'INR    13/02/2004  BMIDS682    Address Targeting Processing
'INR    13/02/2004  BMIDS730    Address Targeting Processing
'HMA    28/07/2004  BMIDS748    Add RunUpgradeToFullAppCreditCheck
'----------------------------------------------------------------------------------
Private Sub IOmTmNTxBO_GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                                    ByVal vxmlApplicationDestDoc As FreeThreadedDOMDocument40, _
                                    Optional ByVal vstrRBTemplateName As String)

    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "IOmTmNTxBO_GetApplicationData"
    
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

Const cstrFunctionName As String = "IOmTmNTxBO_ImportBureauData"

Dim objAIPBO As Object
Dim objBO As Object

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlResponseElem As IXMLDOMElement
Dim xmlApplicationNode As IXMLDOMNode
Dim xmlCustomerListNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlCaseTask As IXMLDOMNode
Dim strResponse As String
Dim xmlBureauDataNode As IXMLDOMNode

    'Initialise
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    
    'BMIDS658
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    If Not xmlApplicationNode.hasChildNodes Then
        Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
    End If
    
    Set xmlCustomerListNode = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMERLIST")
    Set xmlBureauDataNode = xmlGetMandatoryNode(vxmlRequestNode, "BUREAUDATA")
    'BMIDS658 End
    
    'Import the bureau data =======================================================================
    'Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    'xmlTempRequest.appendChild xmlApplicationNode
    'xmlTempRequest.appendChild xmlCustomerListNode
    'xmlTempRequest.appendChild xmlBureauDataNode
    
    Set objAIPBO = gobjContext.CreateInstance(gstrAIP_COMPONENT & ".AgreementInPrincipleBO")
    strResponse = objAIPBO.ImportBureauData(vxmlRequestNode.xml)
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

    Set objAIPBO = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTask = Nothing
    Set objBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

ImportBureauDataErr:
    App.LogEvent vbCrLf & "Error: " & Err.Number & ", " & Err.Description & ", " & Err.Source, vbLogEventTypeError

    'Abort this transaction
    gobjContext.SetAbort
    GoTo ImportBureauDataExit

End Sub

'BM0340 MDC 24/04/2003
' PSC 22/09/2005 MAR32
Private Sub IOmTmNTxBO_MoveCaseToNewStage(ByVal vxmlApplicationRequestDoc As FreeThreadedDOMDocument40, _
                                            ByVal vxmlCreateCaseStageDoc As FreeThreadedDOMDocument40, _
                                            ByVal vxmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40, _
                                            Optional ByVal vxmlSetSLAExpiryDateDoc As FreeThreadedDOMDocument = Nothing)
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
On Error GoTo MoveCaseToNewStage_Err

Dim objTMBO As omTMBO
Dim objApplicationBO As Object
Dim objSubquote As Object 'MAR1703

Dim xmlReserveProductReqNode As IXMLDOMNode 'MAR1703
Dim xmlMSQNode As IXMLDOMNode 'MAR1703
Dim xmlTempNode As IXMLDOMNode 'MAR1703
Dim xmlAppStageNode As IXMLDOMNode 'MAR1703
Dim xmlThisRequestDoc As FreeThreadedDOMDocument40 'MAR1703
    

Dim strResponse As String
Const cstrFunctionName As String = "IOmTmNTxBO_MoveCaseToNewStage"
    
    'Update the current stage (if necessary)
    If Not vxmlUpdateCurrentStageDoc Is Nothing Then
        IOmTmNTxBO_RunTmRequest vxmlUpdateCurrentStageDoc
    End If
    
    'Create the new case stage
    IOmTmNTxBO_RunTmRequest vxmlCreateCaseStageDoc
    
    'Call omApp.ApplicationBO to create application stage on omiga
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    strResponse = objApplicationBO.CreateApplicationStage(vxmlApplicationRequestDoc.xml)
    errCheckXMLResponse strResponse, True
    
    ' PSC 22/09/2005 MAR32 - Start
    If Not vxmlSetSLAExpiryDateDoc Is Nothing Then
        strResponse = objApplicationBO.DeriveAndSetSLADate(vxmlSetSLAExpiryDateDoc.xml)
        errCheckXMLResponse strResponse, True
        'MAR1703 reserve the mortgageproduct if global parameter set
        If GetGlobalParamBoolean("ReserveMtgProductAppSubmission") = True Then
        
            Set xmlAppStageNode = vxmlApplicationRequestDoc.selectSingleNode(".//APPLICATIONSTAGE")
            Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
            xmlThisRequestDoc.validateOnParse = False
            xmlThisRequestDoc.setProperty "NewParser", True
            xmlThisRequestDoc.async = False
            Set xmlReserveProductReqNode = xmlThisRequestDoc.createElement("REQUEST")
            xmlThisRequestDoc.appendChild xmlReserveProductReqNode
            Set xmlMSQNode = xmlThisRequestDoc.createElement("MORTGAGESUBQUOTE")
            xmlReserveProductReqNode.appendChild xmlMSQNode
            Set xmlTempNode = xmlThisRequestDoc.createElement("APPLICATIONNUMBER")
            xmlTempNode.Text = xmlGetNodeText(xmlAppStageNode, "APPLICATIONNUMBER")
            xmlMSQNode.appendChild xmlTempNode
            Set xmlTempNode = xmlThisRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
            xmlTempNode.Text = xmlGetNodeText(xmlAppStageNode, "APPLICATIONFACTFINDNUMBER")
            xmlMSQNode.appendChild xmlTempNode
            
            Set objSubquote = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
            strResponse = objSubquote.ReserveMortgageProduct(xmlThisRequestDoc.xml)
            errCheckXMLResponse strResponse, True
        End If
    End If
    ' PSC 22/09/2005 MAR32 - End

    gobjContext.SetComplete
    
MoveCaseToNewStage_Exit:
    Set objTMBO = Nothing
    Set objApplicationBO = Nothing
    Set objSubquote = Nothing
    Set xmlReserveProductReqNode = Nothing
    Set xmlMSQNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlAppStageNode = Nothing
    Set xmlThisRequestDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
MoveCaseToNewStage_Err:
    gobjContext.SetAbort
    GoTo MoveCaseToNewStage_Exit

End Sub

'BM0340 MDC 24/04/2003 - End
'BMIDS682 Return address targeting info
Private Function IOmTmNTxBO_RunExperianCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode) As String
' header ----------------------------------------------------------------------------------
' procedure:    IOmTmNTxBO_RunExperianCreditCheck
' author:       Mark Coleman
' date:         22/08/2002
' description:  Performs a Credit Score and Full Bureau Download
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo IOmTmNTxBO_RunExperianCreditCheck_VbErr

Const strFunctionName   As String = "IOmTmNTxBO_RunExperianCreditCheck"
    
Dim xmlNewRequest       As IXMLDOMNode
Dim xmlCaseTask         As IXMLDOMElement
Dim xmlApplicationNode  As IXMLDOMNode
Dim xmlCustomerListNode As IXMLDOMNode
Dim objBO               As Object
Dim strResponse         As String
    
    'BMIDS682
    Dim sAddressTargetReq As String
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    Dim strAddressTargeting As String
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    'BMIDS730
    Dim sFromTM030 As String
    
    'BMIDS682 need casetask info regardless of which type of request
    Set xmlCaseTask = vxmlRequestNode.selectSingleNode(".//CASETASK")
   
    'BMIDS682 Don't Update the Case Task status if this is an Address Target Request, already done
    'xmlGetAttributeText will return an empty string if the attribute does not exist
    'i.e. the request is from other than TM030
    sAddressTargetReq = xmlGetAttributeText(vxmlRequestNode, "ADDRESSTARGETREQ")
    If (StrComp(sAddressTargetReq, "TRUE") <> 0) Then
    
    'BMIDS730 Don't update the case task status here. As this is now in a new transaction
    'there may be errors back in OmTmbo, omExp, omCC etc after this processing and the
    '"UPDATECASETASK" will not be rolled back. Case task status is now updated
    'in omTmBO.RunCreditCheck
'        ' Update the Case Task status ==============================================================
'        ' Build the Request and update the case task
'        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
'
'        xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
'
'        xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
'        xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
'        xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
'        xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
'        xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)
'
'        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
'        strResponse = objBO.TmRequest(xmlNewRequest.xml)
'        errCheckXMLResponse strResponse, True
        
    
        'BMIDS658
        'Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION" _
                        & "[(@APPLICATIONNUMBER != '') and (@APPLICATIONFACTFINDNUMBER != '')]")
        'Set xmlApplicationElem = xmlMakeNodeElementBased(xmlApplicationElem, True, "APPLICATION")
        
        Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
        Set xmlCustomerListNode = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMERLIST")
        'BMIDS658 End
        
        ' Run the Credit Score and Full Bureau Download ============================================
        'Set xmlDoc = New FreeThreadedDOMDocument40
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlNewRequest.appendChild xmlApplicationNode.cloneNode(True)
        xmlNewRequest.appendChild xmlCustomerListNode.cloneNode(True)
        
        Set objBO = gobjContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
        strResponse = objBO.RunCreditCheck(xmlNewRequest.xml)
    Else
        Set objBO = gobjContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
        strResponse = objBO.RunCreditCheck(vxmlRequestNode.xml)
    End If

    
    'BMIDS00521 MDC 27/09/2002
    errCheckXMLResponse strResponse, True
'    lngErrNo = errCheckXMLResponse(strResponse, False)
'    If lngErrNo <> 0 Then
'        errThrowError strFunctionName, 4812
'    End If
    'BMIDS00521 MDC 27/09/2002 - End
    'BMIDS682
    xmlThisResponseDoc.loadXML strResponse
    Set xmlThisResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE")
    strAddressTargeting = xmlGetNodeText(xmlThisResponseNode, "ADDRESSTARGETING")
    
    'if address targeting we need to keep the casetask info
    If (StrComp(strAddressTargeting, "YES") = 0) Then
        xmlThisResponseNode.appendChild xmlCaseTask.cloneNode(True)

        IOmTmNTxBO_RunExperianCreditCheck = xmlThisResponseNode.xml
    Else
        IOmTmNTxBO_RunExperianCreditCheck = strResponse
    End If
    

    gobjContext.SetComplete

IOmTmNTxBO_RunExperianCreditCheck_Exit:
    Set xmlNewRequest = Nothing
    Set xmlCaseTask = Nothing
    Set xmlApplicationNode = Nothing
    Set objBO = Nothing
    Set xmlCustomerListNode = Nothing
    'BMIDS682
    Set xmlThisResponseDoc = Nothing
    Set xmlThisResponseNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Function
    
IOmTmNTxBO_RunExperianCreditCheck_VbErr:

    'BMIDS682
    If errIsApplicationError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    
    IOmTmNTxBO_RunExperianCreditCheck = errCreateErrorResponse()
    
    gobjContext.SetAbort
    
    GoTo IOmTmNTxBO_RunExperianCreditCheck_Exit
    
End Function
'BMIDS00336 MDC 23/08/2002 - End

'BMIDS748  Add function
Private Function IOmTmNTxBO_RunUpgradeToFullAppCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode) As String
' header ----------------------------------------------------------------------------------
' procedure:    IOmTmNTxBO_RunUpgradeToFullAppCreditCheck
' author:       Helen Aldred
' date:         28/07/2004
' description:  Performs an upgrade to a full application 'footprint'
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
'------------------------------------------------------------------------------------------
    On Error GoTo IOmTmNTxBO_RunUpgradeToFullAppCreditCheck_VbErr

    Const strFunctionName   As String = "IOmTmNTxBO_RunUpgradeToFullAppCreditCheck"
    
    Dim objBO               As Object
    Dim strResponse         As String
    
    Set objBO = gobjContext.CreateInstance(gstrCREDITCHECK_COMPONENT & ".CreditCheckBO")
    strResponse = objBO.RunUpgradeToFullAppCreditCheck(vxmlRequestNode.xml)

    errCheckXMLResponse strResponse, True
    
    IOmTmNTxBO_RunUpgradeToFullAppCreditCheck = strResponse

    gobjContext.SetComplete

IOmTmNTxBO_RunUpgradeToFullAppCreditCheck_Exit:
    Set objBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Function
    
IOmTmNTxBO_RunUpgradeToFullAppCreditCheck_VbErr:

    If errIsApplicationError = True Then
        App.LogEvent Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
    End If
    
    IOmTmNTxBO_RunUpgradeToFullAppCreditCheck = errCreateErrorResponse()
    
    gobjContext.SetAbort
    
    GoTo IOmTmNTxBO_RunUpgradeToFullAppCreditCheck_Exit
    
End Function

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
    Const cstrFunctionName As String = "IOmTmNTxBO_RunIncomeCalcs"
    Dim objAIPBO            As Object
    Dim xmlRequest          As IXMLDOMNode
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim strResponse         As String

On Error GoTo RunIncomeCalcsErr
    Set objAIPBO = gobjContext.CreateInstance(gstrAIP_COMPONENT & ".AgreementInPrincipleBO")
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    '_ & "[(@APPLICATIONNUMBER != '') and (@APPLICATIONFACTFINDNUMBER != '')]")
    'BMIDS658 xmlApplicationNode will already be in the correct format when called from ImportBureauData
    If Not xmlApplicationNode.hasChildNodes Then
        Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
    End If
    
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

    Const cstrFunctionName As String = "IOmTmNTxBO_ForwardBORequest"

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
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
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
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    
    Set objBO = gobjContext.CreateInstance(strProgId)
    varResponse = CallByName(objBO, strMethod, VbMethod, xmlThisRequestNode.xml)
    
    xmlThisResponseDoc.loadXML varResponse
        
    Set objBO = Nothing
        
    ' PSC 19/02/2007 EP2_1409 - Start
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    AttachResponseData vxmlResponseNode, xmlThisResponseDoc.documentElement
    ' PSC 19/02/2007 EP2_1409 - End
    
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
    
    Const cstrFunctionName As String = "IOmTmNTxBO_RunCompositeRequest"
    gobjContext.SetComplete
    On Error GoTo RunCompositeRequestExit
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlInRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
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
    Const cstrFunctionName As String = "IOmTmNTxBO_RunTmRequest"
    On Error GoTo RunTmRequestExit
    
    Debug.Print "RunTmRequest request"
    Debug.Print vxmlRequestDoc.xml
    Set gobjTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    
    If vxmlResponseDoc Is Nothing Then
        Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.async = False

        xmlThisResponseDoc.loadXML gobjTmBO.TmRequest(vxmlRequestDoc.xml)
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        Debug.Print "RunTmRequest response"
        Debug.Print xmlThisResponseDoc.xml
        Set xmlThisResponseDoc = Nothing
    Else
        vxmlResponseDoc.async = False
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

'BMIDS697 GHun 22/01/2004 No longer used
''PSC 06/01/2003 BM0089 - Start
'Private Function IOmTmNTxBO_MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'
'    Const cstrFunctionName As String = "IOmTmNTxBO_MoveToNextStage"
'
'    On Error GoTo MoveToNextStageVbErr
'
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim objTmBO As OmTmBO
'
'    Dim strResponse As String
'
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
'
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToNextStageNewTx"
'
'    Set objTmBO = gobjContext.CreateInstance(App.Title & ".omTmBO")
'
'    strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'
'    IOmTmNTxBO_MoveToNextStage = strResponse
'
'    gobjContext.SetComplete
'
'MoveToNextStageExit:
'
'    Set xmlTempRequest = Nothing
'    Set objTmBO = Nothing
'
'    errCheckError cstrFunctionName, TypeName(Me)
'
'    Exit Function
'
'MoveToNextStageVbErr:
'    gobjContext.SetAbort
'
'    GoTo MoveToNextStageExit
'
'End Function
'BMIDS697 End

'BMIDS697 GHun 22/01/2004 No longer used
'Private Function IOmTmNTxBO_MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode) As String
'
'    Const cstrFunctionName As String = "IOmTmNTxBO_MoveToStage"
'
'    On Error GoTo MoveToStageVbErr
'
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim objTmBO As OmTmBO
'
'    Dim strResponse As String
'
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
'
'    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStageNewTx"
'    Set objTmBO = gobjContext.CreateInstance(App.Title & ".omTmBO")
'
'    strResponse = objTmBO.OmTmRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True
'
'    IOmTmNTxBO_MoveToStage = strResponse
'
'    gobjContext.SetComplete
'
'MoveToStageExit:
'
'    Set xmlTempRequest = Nothing
'    Set objTmBO = Nothing
'
'    errCheckError cstrFunctionName, TypeName(Me)
'
'    Exit Function
'
'MoveToStageVbErr:
'    gobjContext.SetAbort
'
'    GoTo MoveToStageExit
'
'End Function
'BMIDS697 End

Private Sub IOmTmNTxBO_ProcessTask(ByVal vxmlRequest As IXMLDOMNode)
    Const cstrFunctionName As String = "IOmTmNTxBO_ProcessTask"
    
    On Error GoTo ProcessTaskVbErr
        
    Dim objOmTmBo As Object
    Dim strResponse As String
        
    gobjContext.SetComplete 'AS 22/03/2006 MAR1501
    
    Set objOmTmBo = gobjContext.CreateInstance(App.Title & ".OmTmBo")
    
    'Call into OmTMBo.ProcessAutomaticTasks
    strResponse = objOmTmBo.OmTmRequest(vxmlRequest.xml)
    errCheckXMLResponse strResponse, True
        
ProcessTaskExit:

    errCheckError cstrFunctionName
    
    Exit Sub
    
ProcessTaskVbErr:
    
    gobjContext.SetAbort
    GoTo ProcessTaskExit

End Sub
'PSC 06/01/2003 BM0089 - End
'PSC 25/08/2005 MAR32 - Start
Private Function IOmTmNTxBO_CreateAdhocCaseTask(ByVal vxmlRequestNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo CreateAdhocCaseTaskError

    Const cstrFunctionName As String = "IOmTmNTxBO_CreateAdhocCaseTask"
    
    Dim xmlCaseTaskInputList As IXMLDOMNodeList
    
    Dim xmlCreateCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlCreateCaseTaskRequestNode As IXMLDOMNode
    Dim xmlCreateCaseTaskNode As IXMLDOMNode
    
    Dim xmlGetTaskDetailDoc As FreeThreadedDOMDocument40
    Dim xmlGetTaskNode As IXMLDOMNode
    
    Dim xmlTaskDetailDoc As FreeThreadedDOMDocument40
    Dim xmlTaskDetailNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim xmlProcessTasks As IXMLDOMNode  ' PSC 26/09/2005 MAR32
    
    'MAR1408
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlROTResponseDoc As FreeThreadedDOMDocument40
    Dim bCompletion As Boolean
    Dim strTaskType As String
    Dim strCaseId As String
    Dim colValueIDs As Collection
    Dim varValueId As Variant
    Dim objROTBO As Object
    Dim strCompletionDate As String
    
    'MAR1787
    Dim strCreateTaskID As String
    Dim strStageId As String
    Dim strMandatoryFlag As String
       
    ' validate request ================================================================
    Set xmlCaseTaskInputList = vxmlRequestNode.selectNodes("CASETASK")
    xmlCheckMandatoryNode vxmlRequestNode, "APPLICATION/@APPLICATIONPRIORITY"

    'MAR1408
    Set xmlNewRequest = vxmlRequestNode.selectSingleNode("//REQUEST")
    Set xmlROTResponseDoc = New FreeThreadedDOMDocument40

    ' create CreateCaseTask request ====================================================
    ' attributes from original casetask request
    Set xmlCreateCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlCreateCaseTaskDoc.validateOnParse = False
    xmlCreateCaseTaskDoc.setProperty "NewParser", True
    
    Set xmlProcessTasks = xmlCreateCaseTaskDoc.createElement("TASKSTOPROCESS")
    
    Set xmlCreateCaseTaskRequestNode = _
        xmlCreateCaseTaskDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlCreateCaseTaskRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"
    
    ' create GetTaskDetail request =====================================================
    Set xmlGetTaskDetailDoc = New FreeThreadedDOMDocument40
    xmlGetTaskDetailDoc.validateOnParse = False
    xmlGetTaskDetailDoc.setProperty "NewParser", True
    xmlGetTaskDetailDoc.async = False
    ' clone original REQUEST node
    Set xmlNode = xmlGetTaskDetailDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskDetail"
    
    Set xmlElem = xmlGetTaskDetailDoc.createElement("TASK")
    
    xmlSetAttributeValue _
        xmlElem, _
        "CASEPRIORITY", _
        vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONPRIORITY").Text
    
    Set xmlGetTaskNode = xmlNode.appendChild(xmlElem)
    
    ' create GetTaskDetail response doc. ===============================================
    Set xmlTaskDetailDoc = New FreeThreadedDOMDocument40
    xmlTaskDetailDoc.validateOnParse = False
    xmlTaskDetailDoc.setProperty "NewParser", True
    xmlTaskDetailDoc.async = False
    
    'BM0340 MDC 17/04/2003
    Dim strUserId As String
    Dim strUnitId As String
    Call GetApplicationOwners(gobjContext, vxmlRequestNode, strUserId, strUnitId)
    'BM0340 MDC 17/04/2003 - End
    
    ' create CaseTask elements
    For Each xmlNode In xmlCaseTaskInputList
        
        xmlCheckMandatoryAttribute xmlNode, "SOURCEAPPLICATION"
        xmlCheckMandatoryAttribute xmlNode, "CASEID"
        xmlCheckMandatoryAttribute xmlNode, "ACTIVITYID"
        xmlCheckMandatoryAttribute xmlNode, "ACTIVITYINSTANCE"
        xmlCheckMandatoryAttribute xmlNode, "STAGEID"
        xmlCheckMandatoryAttribute xmlNode, "TASKID"

        strCreateTaskID = xmlGetAttributeText(xmlNode, "TASKID")               ' MAR1787
        strStageId = xmlGetAttributeText(xmlNode, "STAGEID")                   ' MAR1787
        strMandatoryFlag = xmlGetAttributeText(xmlNode, "MANDATORYFLAG")       ' MAR1787

        Set xmlCreateCaseTaskNode = _
            xmlCreateCaseTaskRequestNode.appendChild(xmlNode.cloneNode(True))
            
        xmlCopyAttribute xmlNode, xmlGetTaskNode, "TASKID"
        'BM0553
        xmlCopyAttribute xmlNode, xmlGetTaskNode, "STAGEID"
            
        ' run GetTaskDetail request
        IOmTmNTxBO_RunTmRequest xmlGetTaskDetailDoc, xmlTaskDetailDoc
        
        Set xmlTaskDetailNode = xmlTaskDetailDoc.selectSingleNode("RESPONSE/TASK")
        
        If Not xmlTaskDetailNode Is Nothing Then
            'BM0340 MDC 17/04/2003
            'AddDefaultValuesToCaseTask vxmlRequestNode, xmlCreateCaseTaskNode, xmlTaskDetailNode
            'MAR7 GHun - CORE174 AddDefaultValuesToCaseTask moved to omTMFunctions.bas; pass gobjContext.
            
            'MAR1408 If this is a Post Completion task, set up the Due Date based on Completion Date
            Set colValueIDs = New Collection
            GetValueIdsForValidationType "TaskType", "PC", colValueIDs
                                        
            strTaskType = xmlGetAttributeText(xmlTaskDetailNode, "TASKTYPE")
            bCompletion = False
            For Each varValueId In colValueIDs
                If strTaskType = varValueId Then
                    bCompletion = True
                End If
            Next
                        
            If (bCompletion = True) Then
                
                strCaseId = xmlGetAttributeText(xmlNode, "CASEID")
            
                'Get the Completion Date from Report On Title
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETREPORTONTITLEDATA"
    
                Set xmlApplication = xmlNewRequest.selectSingleNode("APPLICATION").cloneNode(True)
                xmlChangeNodeName xmlApplication, "APPLICATION", "REPORTONTITLE"
                xmlSetAttributeValue xmlApplication, "APPLICATIONNUMBER", strCaseId
                xmlSetAttributeValue xmlApplication, "APPLICATIONFACTFINDNUMBER", "1"
                xmlNewRequest.appendChild xmlApplication
                
                Set objROTBO = gobjContext.CreateInstance("omROT.omRotBO")
                xmlROTResponseDoc.loadXML objROTBO.OmRotRequest(xmlNewRequest.xml)
                                
                Set xmlTempNode = xmlROTResponseDoc.selectSingleNode("//REPORTONTITLE")
                                
                If Not (xmlTempNode Is Nothing) Then
                    strCompletionDate = xmlGetAttributeText(xmlTempNode, "COMPLETIONDATE")
                    'MAR1547 GHun
                    If xmlGetAttributeText(xmlNode, "TASKID") = GetGlobalParamString("TMDelayCompletionTaskID") Then
                        strCompletionDate = strCompletionDate & " " & GetGlobalParamString("TMDlyCompTaskDueTime")
                    End If
                    'MAR1547 End
                    'Add the due date to the Case Task node
                    SetCaseTaskDueDateAndTime xmlCreateCaseTaskNode, xmlTaskDetailNode, strCompletionDate
                              
                End If
                            
            End If
         
            AddDefaultValuesToCaseTask gobjContext, vxmlRequestNode, xmlCreateCaseTaskNode, xmlTaskDetailNode, strUserId, strUnitId
            
            ' PSC 26/09/2005 MAR32 - Start
            If xmlGetAttributeAsBoolean(xmlTaskDetailNode, "ALWAYSAUTOMATICONCREATION") Then
                xmlProcessTasks.appendChild xmlTaskDetailNode.cloneNode(True)
            End If
            ' PSC 26/09/2005 MAR32 - End
            
            'BM0340 MDC 17/04/2003 - End
        End If
    
        ' AQR SYS1791
        ' add ORIGINATINGSTAGEID
        xmlCopyAttributeValue _
            xmlNode, xmlCreateCaseTaskNode, "STAGEID", "ORIGINATINGSTAGEID"
            
        'DB BM0248 - Add mandatory flag.
        'MAR1787  Get the mandatory flag if it has not been passed in.
        If (Len(strMandatoryFlag) = 0) Then
            strMandatoryFlag = GetTaskMandatoryFlag(strCreateTaskID, strStageId)
            xmlSetAttributeValue xmlCreateCaseTaskNode, "MANDATORYINDICATOR", strMandatoryFlag
        Else
            xmlCopyAttributeValue xmlNode, xmlCreateCaseTaskNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
        End If
        'DB End
    Next
    
    'PSC 26/09/2005 MAR32 - Start
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlReturnedCaseTask As IXMLDOMNode
    Dim xmlReturnedCaseTaskList As IXMLDOMNodeList
    Dim strTaskId As String
    Dim intIndex As Integer
    'MAR972 GHun
    Dim xmlTaskNote As IXMLDOMElement
    Dim strResponse As String
    'MAR1972 End
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlNode = xmlResponseDoc.createElement("RESPONSE")
    xmlResponseDoc.appendChild xmlNode
            
    ' run CreateCaseTask request
    IOmTmNTxBO_RunTmRequest xmlCreateCaseTaskDoc, xmlResponseDoc
        
    'MAR972 GHun
    For Each xmlTaskNote In xmlResponseDoc.selectNodes("RESPONSE/CASETASK/TASKNOTE")
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "CreateTaskNote"
        
        'copy the primary key from parent casetask
        Set xmlElem = xmlTaskNote.parentNode
        xmlTaskNote.setAttribute "CASEACTIVITYGUID", xmlElem.getAttribute("CASEACTIVITYGUID")
        xmlTaskNote.setAttribute "STAGEID", xmlElem.getAttribute("STAGEID")
        xmlTaskNote.setAttribute "CASESTAGESEQUENCENO", xmlElem.getAttribute("CASESTAGESEQUENCENO")
        xmlTaskNote.setAttribute "TASKID", xmlElem.getAttribute("TASKID")
        xmlTaskNote.setAttribute "TASKINSTANCE", xmlElem.getAttribute("TASKINSTANCE")
        
        xmlNewRequest.appendChild xmlTaskNote.cloneNode(True)
        strResponse = gobjTmBO.TmRequest(xmlNewRequest.xml)
        errCheckXMLResponse strResponse, True, xmlResponseDoc.documentElement
    Next
    'MAR972 End
    
    ' Get the returned CaseTaskList
    Set xmlReturnedCaseTaskList = xmlResponseDoc.selectNodes("RESPONSE/CASETASK")
    
    ' For each one check to see if it is to be processed on creation.
    For Each xmlReturnedCaseTask In xmlReturnedCaseTaskList
        strTaskId = xmlGetAttributeText(xmlReturnedCaseTask, "TASKID")
        Set xmlTaskDetailNode = xmlProcessTasks.selectSingleNode("TASK[@TASKID='" & strTaskId & "']")
    
        ' If to be processed on creation copy the task details else remove it from the response
        If Not xmlTaskDetailNode Is Nothing Then
            For intIndex = 0 To xmlTaskDetailNode.Attributes.length - 1
                xmlCopyAttribIfMissingFromDest xmlTaskDetailNode, xmlReturnedCaseTask, xmlTaskDetailNode.Attributes.Item(intIndex).nodeName
            Next
        Else
            xmlReturnedCaseTask.parentNode.removeChild xmlReturnedCaseTask
        End If
    Next
    
    Set IOmTmNTxBO_CreateAdhocCaseTask = xmlResponseDoc.documentElement
    'PSC 26/09/2005 MAR32 - End
    gobjContext.SetComplete
    
CreateAdhocCaseTaskExit:
    
    ' PSC 26/09/2005 MAR32 - Start
    Set xmlCaseTaskInputList = Nothing
    Set xmlCreateCaseTaskDoc = Nothing
    Set xmlCreateCaseTaskRequestNode = Nothing
    Set xmlCreateCaseTaskNode = Nothing
    Set xmlGetTaskDetailDoc = Nothing
    Set xmlGetTaskNode = Nothing
    Set xmlTaskDetailDoc = Nothing
    Set xmlTaskDetailNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlProcessTasks = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlReturnedCaseTask = Nothing
    Set xmlReturnedCaseTaskList = Nothing
    ' PSC 26/09/2005 MAR32 - End
    Set xmlTaskNote = Nothing   'MAR972 GHun
    Set xmlNewRequest = Nothing 'MAR972 GHun

    errCheckError cstrFunctionName, TypeName(Me)
    Exit Function
    
CreateAdhocCaseTaskError:
    gobjContext.SetAbort
    GoTo CreateAdhocCaseTaskExit

End Function

Private Sub IOmTmNTxBO_ProcessTASTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ProcessTASTaskExit
    Const cstrFunctionName As String = "IOmTmNTxBO_ProcessTASTask"
                
    Dim objNTTxBO As IOmTmNTxBO
    Dim xmlCaseTask As IXMLDOMNode
    
    Dim strTASStatus As String
    Dim strTaskStatus As String
        
    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"

    Set objNTTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBo")
    objNTTxBO.PerformTASTaskProcessing vxmlRequestNode, vxmlResponseNode
    errCheckXMLResponseNode vxmlResponseNode, , True

    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "SUCCESS"

ProcessTASTaskExit:
    
    If Err.Number <> 0 Then
    
        DoRetry vxmlRequestNode, vxmlResponseNode
        
        ' Set status to success
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
        
    End If
    
    Set objNTTxBO = Nothing
    Set xmlCaseTask = Nothing
    
    gobjContext.SetComplete

    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'PSC 09/05/2006 MAR1643
Private Sub IOmTmNTxBO_LockApplication(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String, ByVal strAppFFNo As String)
    
    On Error GoTo LockApplicationError
    Const cstrFunctionName As String = "LockApplication"
    
    Dim objAppManBO As Object
    
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
    
    gobjContext.SetComplete

LockApplicationExit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
LockApplicationError:
    gobjContext.SetAbort
    GoTo LockApplicationExit

End Sub

'PSC 09/05/2006 MAR1643
Private Sub IOmTmNTxBO_UnlockApplication(ByVal xmlTempRequest As IXMLDOMNode, ByVal strAppNo As String)

    On Error GoTo UnlockApplicationError
    Const cstrFunctionName As String = "UnlockApplication"

    Dim objAppManBO As Object
    
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set objAppManBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    'Create Request
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlAppNode.Text = strAppNo
    xmlTempRequest.appendChild xmlAppNode
    
    'Unlock the Application and Customers
    strResponse = objAppManBO.UnlockApplicationAndCustomers(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    
    gobjContext.SetComplete

UnlockApplicationExit:
    Set xmlTempRequest = Nothing
    Set objAppManBO = Nothing
    Set xmlAppNode = Nothing
    Set xmlDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
   
UnlockApplicationError:
    gobjContext.SetAbort
    GoTo UnlockApplicationExit

End Sub

Private Sub DoRetry(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo DoRetryExit
    Const cstrFunctionName As String = "DoRetry"
    
    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlCaseTaskDetails As IXMLDOMNode
    Dim xmlCurrentStage As IXMLDOMNode

    Dim strQueueName As String
    Dim strResponse As String
    Dim strTASStatus As String
    Dim strErrorMessage As String
  
    Dim lngAllowableRetries As Long
    Dim lngCurrentRetries As Long
    Dim lngDelayInterval As Long
   
    Set xmlCaseTaskDetails = vxmlResponseNode.selectSingleNode("CASETASK")
    Set xmlCaseTaskIn = vxmlRequestNode.selectSingleNode("CASETASK")
    Set xmlCurrentStage = vxmlRequestNode.selectSingleNode("CURRENTSTAGE")
    
    If Not xmlCaseTaskIn Is Nothing Then
        strQueueName = xmlGetMandatoryAttributeText(xmlCaseTaskIn, "TARGETQUEUENAME")
    Else
        strQueueName = xmlGetMandatoryAttributeText(xmlCurrentStage, "TARGETQUEUENAME")
    End If
    
    strTASStatus = xmlGetAttributeText(vxmlResponseNode, "TASSTATUS")
    
    Select Case strTASStatus
        ' PSC 03/08/2006 MAR1927
        Case "RETRYAPPLICATIONLOCK"
            lngAllowableRetries = xmlGetAttributeAsLong(xmlCaseTaskIn, "MAXAPPLICATIONLOCKRETRIES")
            lngCurrentRetries = xmlGetAttributeAsLong(xmlCaseTaskIn, "APPLICATIONLOCKRETRIES", "0")
            lngDelayInterval = xmlGetAttributeAsLong(xmlCaseTaskIn, "APPLICATIONLOCKRETRYINTERVAL", "0")
            
            ' If the number of retries hasn't been reached increment it to be put back on the message queue
            If lngCurrentRetries < lngAllowableRetries Then
                xmlSetAttributeValue xmlCaseTaskIn, "APPLICATIONLOCKRETRIES", CStr(lngCurrentRetries + 1)
            End If
        ' PSC 03/08/2006 MAR1927
        Case "RETRYTASK", "RETRYGETTASKDETAILS", "RETRYGETCURRENTSTAGE"
            lngAllowableRetries = xmlGetAttributeAsInteger(xmlCaseTaskIn, "MAXPROCESSRETRIES")
            lngCurrentRetries = xmlGetAttributeAsInteger(xmlCaseTaskDetails, "TASRETRYCOUNT", "0")
            lngDelayInterval = xmlGetAttributeAsLong(xmlCaseTaskIn, "PROCESSRETRYINTERVAL", "0")
            
            ' Have been able to lock application so reset APPLICATIONLOCKRETRIES
            xmlSetAttributeValue xmlCaseTaskIn, "APPLICATIONLOCKRETRIES", "0"
        Case "MOVESTAGERETRYAPPLICATIONLOCK"
            lngAllowableRetries = xmlGetAttributeAsLong(xmlCurrentStage, "MAXAPPLICATIONLOCKRETRIES")
            lngCurrentRetries = xmlGetAttributeAsLong(xmlCurrentStage, "APPLICATIONLOCKRETRIES", "0")
            lngDelayInterval = xmlGetAttributeAsLong(xmlCurrentStage, "APPLICATIONLOCKRETRYINTERVAL", "0")
            
            ' If the number of retries hasn't been reached increment it to be put back on the message queue
            If lngCurrentRetries < lngAllowableRetries Then
                xmlSetAttributeValue xmlCurrentStage, "APPLICATIONLOCKRETRIES", CStr(lngCurrentRetries + 1)
            End If
        Case "MOVESTAGERETRY"
            lngAllowableRetries = xmlGetAttributeAsLong(xmlCurrentStage, "MAXPROCESSRETRIES")
            lngCurrentRetries = xmlGetAttributeAsLong(xmlCurrentStage, "MOVETOSTAGERETRIES", "0")
            lngDelayInterval = xmlGetAttributeAsLong(xmlCurrentStage, "PROCESSRETRYINTERVAL", "0")
            
            ' If the number of retries hasn't been reached increment it to be put back on the message queue
            If lngCurrentRetries < lngAllowableRetries Then
                xmlSetAttributeValue xmlCurrentStage, "MOVETOSTAGERETRIES", CStr(lngCurrentRetries + 1)
            End If
    End Select
    
    ' If allowable retries has not been reached put the message back on the queue
    If lngCurrentRetries < lngAllowableRetries Then
        IOmTmNTxBO_SendMessageToQueue vxmlRequestNode, vxmlRequestNode, strQueueName, lngDelayInterval
    Else
        ' Set the TASSTATUS to failure so that message is put onto dead queue
        If strTASStatus <> "WARNING" Then
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"
        End If
        
        'Log to the application log
        strErrorMessage = "Task Automation Error." & vbCrLf & _
        "Request: " & vxmlRequestNode.xml & vbCrLf & _
        "Number: " & xmlGetNodeText(vxmlResponseNode, "ERROR/NUMBER") & vbCrLf & _
        "Source: " & xmlGetNodeText(vxmlResponseNode, "ERROR/SOURCE") & vbCrLf & _
        "Description: " & xmlGetNodeText(vxmlResponseNode, "ERROR/DESCRIPTION")
        App.LogEvent strErrorMessage, vbLogEventTypeError

        ' Log to error table
    End If
    
DoRetryExit:

    Set xmlCaseTaskIn = Nothing
    Set xmlCaseTaskDetails = Nothing
    Set xmlCurrentStage = Nothing
    
    If Err.Number <> 0 Then
        ' Log error
    End If

End Sub

Private Function GetApplicationLockStatus(vstrApplicationNumber As String) As Integer
On Error GoTo GetApplicationLockStatusExit

Const strFunctionName As String = "GetApplicationLockStatus"

    Dim objAppManBO As Object
    
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    
    Dim lngErr As Long

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'Create Request
    Set xmlTempRequest = xmlDoc.createElement("REQUEST")
    Set xmlAppNode = xmlDoc.createElement("APPLICATIONLOCK")
    Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = vstrApplicationNumber
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
    
GetApplicationLockStatusExit:
    Set objAppManBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlAppNode = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function CanMoveStage(vxmlRequestNode As IXMLDOMNode) As Boolean
On Error GoTo CanMoveStageExit
    
    Const cstrFunctionName = "CanMoveStage"

    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlCaseTasks As IXMLDOMNodeList
    Dim objTMBO As Object
    
    Dim blnOkToMoveStage As Boolean
    Dim intNoOfTasks As Integer
    Dim intIndex As Integer
    Dim intTaskStatus As Integer
    Dim strResponse As String
    Dim blnAllowMoveStage As Boolean
 
    blnOkToMoveStage = False
     
    Set xmlCaseTaskIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    blnAllowMoveStage = xmlGetAttributeAsBoolean(xmlCaseTaskIn, "TASLASTTASKALLOWSTAGEMOVE")

    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETCASETASKS"
    
    If blnAllowMoveStage Then
        
        blnOkToMoveStage = True
    
        Set xmlCaseTask = vxmlRequestNode.ownerDocument.createElement("CASETASK")
        xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "CASEACTIVITYGUID"
        xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "STAGEID"
        xmlCopyAttribute xmlCaseTaskIn, xmlCaseTask, "CASESTAGESEQUENCENO"
        xmlTempRequest.appendChild xmlCaseTask
        
        Set objTMBO = gobjContext.CreateInstance("MsgTM.msgTMBO")
        strResponse = objTMBO.TmRequest(xmlTempRequest.xml)
        Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        Set xmlCaseTasks = xmlTempResponse.selectNodes("RESPONSE/CASETASK")
        intNoOfTasks = xmlCaseTasks.length
        
        intIndex = 0
        
        While intIndex < intNoOfTasks And blnOkToMoveStage
            Set xmlCaseTask = xmlCaseTasks.Item(intIndex)
            intTaskStatus = xmlGetAttributeAsInteger(xmlCaseTask, "TASKSTATUS", 0)
            
            ' PSC 02/05/2006 MAR1692 - Start
            If IsValidationType("TaskStatus", intTaskStatus, "I") And Not xmlGetAttributeAsBoolean(xmlCaseTask, "CARRYFORWARDINDICATOR") Then
                blnOkToMoveStage = False
            End If
            ' PSC 02/05/2006 MAR1692 - End
            intIndex = intIndex + 1
        Wend
    End If
    
    CanMoveStage = blnOkToMoveStage
   
CanMoveStageExit:

    Set xmlCaseTaskIn = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTask = Nothing
    Set xmlCaseTasks = Nothing
    Set objTMBO = Nothing

    
    errCheckError cstrFunctionName, TypeName(Me)
  
End Function

Private Sub IOmTmNTxBO_UpdateCaseTask(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo UpdateCaseTaskExit
    Const cstrFunctionName As String = "IOmTmNTxBO_UpdateCaseTask"
    
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim objTMBO As Object
    
    Dim strResponse As String
    
    Set objTMBO = gobjContext.CreateInstance("MsgTM.msgTMBO")
    
    strResponse = objTMBO.TmRequest(vxmlRequestNode.xml)
    Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempResponse.documentElement, , True

UpdateCaseTaskExit:

    Set xmlTempResponse = Nothing
    Set objTMBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IOmTmNTxBO_SendMessageToQueue(ByVal vxmlRequestNode As IXMLDOMNode, _
                                          ByVal vxmlMessageToSend As IXMLDOMNode, _
                                          vstrQueueName As String, _
                                          Optional ByVal lngDelayInterval As Long = 0)

    On Error GoTo SendMessageToQueueExit
    Const cstrFunctionName As String = "IOmTmNTxBO_SendMessageToQueue"
    
    Dim objMessageQueue As Object
    Dim xmlSendToQueueRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    Dim dteExecutionDate As Date
   
    Set xmlTempResponse = New FreeThreadedDOMDocument40
    xmlTempResponse.validateOnParse = False
    xmlTempResponse.setProperty "NewParser", True

    Set objMessageQueue = gobjContext.CreateInstance("omMQ.omMQBO")
    Set xmlSendToQueueRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlSendToQueueRequest, "OPERATION", "SendToQueue"
    Set xmlNode = xmlTempResponse.createElement("MESSAGEQUEUE")
    xmlSetAttributeValue xmlNode, "QUEUENAME", vstrQueueName
    xmlSetAttributeValue xmlNode, "PROGID", "omTM.omTMBO"
    
    If lngDelayInterval > 0 Then
        dteExecutionDate = Now
        dteExecutionDate = DateAdd("n", lngDelayInterval, dteExecutionDate)
        xmlSetAttributeValue xmlNode, "EXECUTEAFTERDATE", Format$(dteExecutionDate, "yyyy-mm-dd hh:nn:ss")
    End If
    
    Set xmlDataNode = xmlTempResponse.createElement("XML")
    xmlNode.appendChild xmlDataNode
    xmlDataNode.appendChild vxmlMessageToSend.cloneNode(True)
    xmlSendToQueueRequest.appendChild xmlNode
    strResponse = objMessageQueue.omRequest(xmlSendToQueueRequest.xml)
    Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempResponse.documentElement, , True
    
SendMessageToQueueExit:

    Set objMessageQueue = Nothing
    Set xmlSendToQueueRequest = Nothing
    Set xmlNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlTempResponse = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub IOmTmNTxBO_PerformTASTaskProcessing(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo PerformTASTaskProcessingError
    Const cstrFunctionName As String = "IOmTmNTxBO_PerformTASTaskProcessing"
    
    Const lngAppErrorNo As Long = vbObjectError + 512 + 110
                
    Dim objNTTxBO As IOmTmNTxBO
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlMoveToStageRequest As IXMLDOMNode
    Dim xmlCurrentStage As IXMLDOMNode
    Dim xmlDescription As IXMLDOMNode
    Dim xmlSource As IXMLDOMNode
    Dim xmlTempError As IXMLDOMNode
    Dim objNoTxBO As omTmNoTxBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlStoredCaseTask As IXMLDOMNode
    Dim xmlRevisedRequest As IXMLDOMNode
    Dim xmlCaseTaskUpdateRequest As IXMLDOMNode
    Dim xmlCaseTaskUpdate As IXMLDOMNode
    Dim xmlCaseActivity As IXMLDOMNode  ' PSC 03/08/2006 MAR1927
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strResponse As String
    Dim strQueueName As String
    Dim strMessage As String
    Dim strDescription As String
    Dim strSource As String
    Dim strValueId As String
    Dim strMessageSource As String
   
    Dim intLockStatus As APPLICATIONLOCKSTATUS
    Dim intCurrentProgress As TASPROGRESS
   
    Dim intTaskStatus As Integer
    
    Dim lngErrorNo As Long
    Dim lngCurrentRetries As Long
    Dim lngAllowableRetries As Long
    
    Dim blnCaseSuccessfullyLocked As Boolean  ' PSC 09/05/2008 MAR1643

    Set xmlTempResponse = New FreeThreadedDOMDocument40
    xmlTempResponse.validateOnParse = False
    xmlTempResponse.setProperty "NewParser", True
    
    ' Validate Request
    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"
    intCurrentProgress = TASPROGRESS.ValidateRequest
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    strAppNo = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAppFFNo = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    strQueueName = xmlGetMandatoryAttributeText(xmlCaseTask, "TARGETQUEUENAME")
    
    intCurrentProgress = TASPROGRESS.GettingLockStatus
    intLockStatus = GetApplicationLockStatus(strAppNo)
    intCurrentProgress = TASPROGRESS.LockingApplication
    
    ' PSC 09/05/2008 MAR1643
    Set objNTTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBo")
    
    If intLockStatus = APPLICATIONLOCKSTATUS.lsUnlocked Then
          
        ' PSC 09/05/2008 MAR1643 - Start
        objNTTxBO.LockApplication vxmlRequestNode.cloneNode(False), strAppNo, strAppFFNo
        blnCaseSuccessfullyLocked = True
        ' PSC 09/05/2008 MAR1643 - End
        
        ' PSC 03/08/2006 MAR1927 - Start
        intCurrentProgress = TASPROGRESS.CheckingCurrentStage
        
        ' Get the current stage to make sure it is still the same
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetCurrentStage"
        Set xmlCaseActivity = xmlTempResponse.createElement("CASEACTIVITY")
        xmlTempRequest.appendChild xmlCaseActivity
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "CASEACTIVITYGUID"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "CASEID"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "ACTIVITYID"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "ACTIVITYINSTANCE"
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        Set xmlCurrentStage = xmlTempResponse.selectSingleNode("RESPONSE/CASESTAGE")
        
        ' Process task if case is still at the same stage
        If xmlGetAttributeText(xmlCaseTask, "STAGEID") = xmlGetAttributeText(xmlCurrentStage, "STAGEID") Then
                
            ' Get the additional data required for processing
            intCurrentProgress = TASPROGRESS.GettingTaskDetails
            Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
            xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetAdditionalCaseTaskData"
            strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlTempResponse.documentElement, , True
            
            Set xmlStoredCaseTask = xmlTempResponse.selectSingleNode("RESPONSE/CASETASK")
                     
            ' Check task still has the correct status
            intTaskStatus = xmlGetAttributeAsInteger(xmlStoredCaseTask, "TASKSTATUS")
            
            If IsValidationType("TaskStatus", intTaskStatus, "I") Then
                               
                ' Create a new request with all the required attributes
                Set xmlRevisedRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlRevisedRequest, "OPERATION", "ProcessTask"
                xmlRevisedRequest.appendChild xmlStoredCaseTask
                xmlSetAttributeValue xmlStoredCaseTask, "CASEID", strAppNo
                xmlSetAttributeValue xmlStoredCaseTask, "COMPLETETASKONPRINT", "1"
                xmlRevisedRequest.appendChild xmlApplication.cloneNode(True)
                
                ' Add the case task to the response for processing later
                vxmlResponseNode.appendChild xmlStoredCaseTask.cloneNode(False)
                
                intCurrentProgress = TASPROGRESS.ProcessingTask
                
                On Error Resume Next
                objNTTxBO.ProcessTask xmlRevisedRequest
                
                lngErrorNo = Err.Number
                strSource = Err.Source
                strDescription = Err.Description
                
                On Error GoTo PerformTASTaskProcessingError
                
                ' If there was an error processing the task then update the case task with the
                ' appropriate status
                If lngErrorNo <> 0 Then
                
                    ' Set the progress to UpdatingTaskStatus for error processing
                    intCurrentProgress = TASPROGRESS.UpdatingTaskStatus
                    lngAllowableRetries = xmlGetAttributeAsLong(xmlCaseTask, "MAXPROCESSRETRIES")
                    lngCurrentRetries = xmlGetAttributeAsLong(xmlStoredCaseTask, "TASRETRYCOUNT")
                
                    Set xmlCaseTaskUpdateRequest = vxmlRequestNode.cloneNode(False)
                    xmlSetAttributeValue xmlCaseTaskUpdateRequest, "OPERATION", "UpdateCaseTask"
                    
                    Set xmlCaseTaskUpdate = xmlTempResponse.createElement("CASETASK")
                    xmlCaseTaskUpdateRequest.appendChild xmlCaseTaskUpdate
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "CASEACTIVITYGUID"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "STAGEID"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "CASESTAGESEQUENCENO"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "TASKID"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "TASKINSTANCE"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "SOURCEAPPLICATION"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "ACTIVITYID"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "ACTIVITYINSTANCE"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "CASEID"
                    xmlCopyAttribute xmlCaseTask, xmlCaseTaskUpdate, "CASEACTIVITYGUID"
                    
                    If lngCurrentRetries < lngAllowableRetries Then
                        strValueId = GetFirstComboValueId("TaskStatus", "TR")
                        xmlSetAttributeValue xmlCaseTaskUpdate, "TASRETRYCOUNT", CStr(lngCurrentRetries + 1)
                    Else
                        strMessageSource = xmlGetAttributeText(xmlCaseTask, "MESSAGESOURCE")
                        
                        If strMessageSource <> "TAR" Then
                            strValueId = GetFirstComboValueId("TaskStatus", "TF")
                        Else
                            strValueId = GetFirstComboValueId("TaskStatus", "TRF")
                        End If
                        xmlSetAttributeValue xmlCaseTaskUpdate, "TASRETRYCOUNT", "0"
                        xmlSetAttributeValue xmlCaseTaskUpdate, "TASINPROGRESS", "0"
                    End If
                               
                    xmlSetAttributeValue xmlCaseTaskUpdate, "TASKSTATUS", strValueId
                    objNTTxBO.UpdateCaseTask xmlCaseTaskUpdateRequest
                    
                    ' Set the progress back to ProcessingTask so that the original error will be
                    ' handled correctly
                    intCurrentProgress = TASPROGRESS.ProcessingTask
                    Err.Raise lngErrorNo, strSource, strDescription
                End If
                
                intCurrentProgress = TASPROGRESS.DetermineMoveStage
                              
                If CanMoveStage(xmlRevisedRequest) Then
                    intCurrentProgress = TASPROGRESS.SendMoveStageMessage
                    Set xmlMoveToStageRequest = vxmlRequestNode.cloneNode(False)
                    xmlSetAttributeValue xmlMoveToStageRequest, "OPERATION", "TASMoveToNextStage"
                    Set xmlCurrentStage = xmlTempResponse.createElement("CURRENTSTAGE")
                    xmlMoveToStageRequest.appendChild xmlCurrentStage
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "CASEACTIVITYGUID"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "STAGEID"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "CASESTAGESEQUENCENO"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "SOURCEAPPLICATION"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "CASEID"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "ACTIVITYID"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "ACTIVITYINSTANCE"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "TARGETQUEUENAME"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "MAXAPPLICATIONLOCKRETRIES"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "APPLICATIONLOCKRETRYINTERVAL"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "PROCESSRETRYINTERVAL"
                    xmlCopyAttribute xmlCaseTask, xmlCurrentStage, "MAXPROCESSRETRIES"
                    xmlMoveToStageRequest.appendChild xmlApplication.cloneNode(False)
                    objNTTxBO.SendMessageToQueue vxmlRequestNode, xmlMoveToStageRequest, strQueueName, 1
                End If
            End If
        End If
        ' PSC 03/08/2006 MAR1927 - End
   
        intCurrentProgress = TASPROGRESS.UnLockingApplication
    ElseIf intLockStatus = APPLICATIONLOCKSTATUS.lsLocked Then
        Err.Raise lngAppErrorNo, cstrFunctionName, "Application is already locked"
    Else
        Err.Raise lngAppErrorNo, cstrFunctionName, "Cannot determine application lock status"
    End If
    
    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    
    gobjContext.SetComplete
    
UnlockApplication:

    ' PSC 09/05/2006 MAR1643 - Start
    If blnCaseSuccessfullyLocked Then
        On Error Resume Next
        
        ' Try to unlock application
        objNTTxBO.UnlockApplication vxmlRequestNode.cloneNode(False), strAppNo
        
        lngErrorNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
        
        On Error GoTo PerformTASTaskProcessingError
        
        If lngErrorNo <> 0 Then
            App.LogEvent "Unable to unlock application " & strAppNo & ". " & strDescription, vbLogEventTypeWarning
            
            ' If error was caused during unlockin operation of TAS processing raise an error
            If intCurrentProgress = TASPROGRESS.UnLockingApplication Then
                Err.Raise lngErrorNo, strSource, strDescription
            End If
        End If
    End If
    ' PSC 09/05/2006 MAR1643 - End
    
PerformTASTaskProcessingExit:
        
    Set objNTTxBO = Nothing
    Set xmlTempResponse = Nothing
    Set xmlCaseTask = Nothing
    Set xmlApplication = Nothing
    Set xmlNode = Nothing
    Set xmlMoveToStageRequest = Nothing
    Set xmlCurrentStage = Nothing
    Set xmlDescription = Nothing
    Set xmlSource = Nothing
    Set xmlTempError = Nothing
    Set objNoTxBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlStoredCaseTask = Nothing
    Set xmlRevisedRequest = Nothing
    Set xmlCaseTaskUpdateRequest = Nothing
    Set xmlCaseTaskUpdate = Nothing
    Set xmlCaseActivity = Nothing  ' PSC 03/08/2006 MAR1927

    Exit Sub
        
PerformTASTaskProcessingError:
    
    gobjContext.SetAbort
    Set xmlTempError = CreateErrorResponseNode()
    xmlCopyAttribute xmlTempError, vxmlResponseNode, "TYPE"
    vxmlResponseNode.appendChild xmlTempError.firstChild
    
    Set xmlDescription = vxmlResponseNode.selectSingleNode("ERROR/DESCRIPTION")
    Set xmlSource = vxmlResponseNode.selectSingleNode("ERROR/SOURCE")
    
    Select Case intCurrentProgress
        Case TASPROGRESS.ValidateRequest
            strMessage = "The supplied request is invalid."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"
        Case TASPROGRESS.GettingLockStatus
            strMessage = "Unable to determine if the application is already locked."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYAPPLICATIONLOCK"
        Case TASPROGRESS.LockingApplication
            strMessage = "Unable to lock application."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYAPPLICATIONLOCK"
        ' PSC 03/08/2006 MAR1927 - Start
        Case TASPROGRESS.CheckingCurrentStage
            strMessage = "Unable to check current stage."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYGETCURRENTSTAGE"
        ' PSC 03/08/2006 MAR1927 - End
        Case TASPROGRESS.GettingTaskDetails
            strMessage = "Unable to get details for the task."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYGETTASKDETAILS"
        Case TASPROGRESS.ProcessingTask
            strMessage = "Unable to process the task."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYTASK"
        Case TASPROGRESS.UpdatingTaskStatus
            strMessage = "Unable to update the task status."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "RETRYTASK"
        Case TASPROGRESS.DetermineMoveStage
            strMessage = "Unable to determine whether to move to next stage."
            'PSC 09/05/2006 MAR1643
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "WARNING"
        Case TASPROGRESS.SendMoveStageMessage
            strMessage = "Unable to send MoveToNextStage message."
            'PSC 09/05/2006 MAR1643
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "WARNING"
        Case TASPROGRESS.UnLockingApplication
            strMessage = "Unable to unlock application."
            'PSC 09/05/2006 MAR1643
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "WARNING"
    End Select
    
    xmlDescription.Text = strMessage & " " & xmlDescription.Text
    xmlSource.Text = cstrFunctionName & "." & xmlSource.Text
    
    'PSC 09/05/2006 MAR1643 - Start
    ' If error occurred  before unlock stage unlock the application else cleanup
    If intCurrentProgress <> TASPROGRESS.UnLockingApplication Then
        Resume UnlockApplication
    Else
        Resume PerformTASTaskProcessingExit
    End If
    'PSC 09/05/2006 MAR1643 - End
End Sub

Private Sub IOmTmNTxBO_TASMoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ProcessTASTaskExit
    Const cstrFunctionName As String = "IOmTmNTxBO_TASMoveToNextStage"
                
    Dim objNTTxBO As IOmTmNTxBO
    Dim xmlCaseTask As IXMLDOMNode
    
    Dim strTASStatus As String
    Dim strTaskStatus As String
        
    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"

    Set objNTTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBo")
    objNTTxBO.PerformTASMoveToNextStageProcessing vxmlRequestNode, vxmlResponseNode
    errCheckXMLResponseNode vxmlResponseNode, , True

    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "SUCCESS"

ProcessTASTaskExit:
    
    If Err.Number <> 0 Then
    
        DoRetry vxmlRequestNode, vxmlResponseNode
        
        ' Set status to success
        xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
        
    End If
    
    Set objNTTxBO = Nothing
    Set xmlCaseTask = Nothing
    
    gobjContext.SetComplete

    errCheckError cstrFunctionName, TypeName(Me)
End Sub


Private Sub IOmTmNTxBO_PerformTASMoveToNextStageProcessing(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo PerformTASMoveToNextStageProcessingError
    Const cstrFunctionName As String = "IOmTmNTxBO_PerformTASMoveToNextStageProcessing"
    
    Const lngAppErrorNo As Long = vbObjectError + 512 + 110     ' PSC 09/05/2006 MAR1643
    
    Dim objNTTxBO As IOmTmNTxBO     ' PSC 09/05/2006 MAR1643
    Dim objNoTxBO As omTmNoTxBO
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlApplication As IXMLDOMNode
    Dim xmlTempError As IXMLDOMNode
    Dim xmlCurrentStageIn As IXMLDOMNode
    Dim xmlCurrentStage As IXMLDOMNode
    Dim xmlCaseActivity As IXMLDOMNode
    Dim xmlDescription As IXMLDOMNode
    Dim xmlSource As IXMLDOMNode

    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strResponse As String
    Dim strMessage As String
    Dim intLockStatus As APPLICATIONLOCKSTATUS
    Dim intCurrentProgress As TASPROGRESS
    
    ' PSC 09/05/2006 MAR1643 - Start
    Dim lngErrorNo As Long
    Dim strDescription As String
    Dim strSource As String
    ' PSC 09/05/2006 MAR1643 - End
    
    Dim blnCaseSuccessfullyLocked As Boolean  ' PSC 09/05/2008 MAR1643

    Set xmlTempResponse = New FreeThreadedDOMDocument40
    xmlTempResponse.validateOnParse = False
    xmlTempResponse.setProperty "NewParser", True

    intCurrentProgress = TASPROGRESS.ValidateRequest
    Set xmlCurrentStageIn = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTSTAGE")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    strAppNo = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAppFFNo = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")

    intCurrentProgress = TASPROGRESS.GettingLockStatus
    intLockStatus = GetApplicationLockStatus(strAppNo)

    ' PSC 09/05/2008 MAR1643
    Set objNTTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBo")
    
    If intLockStatus = APPLICATIONLOCKSTATUS.lsUnlocked Then

        intCurrentProgress = TASPROGRESS.LockingApplication
        ' PSC 09/05/2006 MAR1643 - Start
        objNTTxBO.LockApplication vxmlRequestNode.cloneNode(False), strAppNo, strAppFFNo
        blnCaseSuccessfullyLocked = True
        ' PSC 09/05/2006 MAR1643 - End
        intCurrentProgress = TASPROGRESS.MovingStage
        ' Get the current stage to make sure it is still the same
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetCurrentStage"
        Set xmlCaseActivity = xmlTempResponse.createElement("CASEACTIVITY")
        xmlTempRequest.appendChild xmlCaseActivity
        xmlCopyAttribute xmlCurrentStageIn, xmlCaseActivity, "CASEACTIVITYGUID"
        xmlCopyAttribute xmlCurrentStageIn, xmlCaseActivity, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCurrentStageIn, xmlCaseActivity, "CASEID"
        xmlCopyAttribute xmlCurrentStageIn, xmlCaseActivity, "ACTIVITYID"
        xmlCopyAttribute xmlCurrentStageIn, xmlCaseActivity, "ACTIVITYINSTANCE"
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        Set xmlCurrentStage = xmlTempResponse.selectSingleNode("RESPONSE/CASESTAGE")
        
        ' Process the message if the case is still at the the same stage
        ' PSC 29/11/2005 MAR491
        If xmlGetAttributeText(xmlCurrentStageIn, "STAGEID") = xmlGetAttributeText(xmlCurrentStage, "STAGEID") Then
            Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
            xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToNextStage"
            strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        End If
        
        intCurrentProgress = TASPROGRESS.UnLockingApplication
    ElseIf intLockStatus = APPLICATIONLOCKSTATUS.lsLocked Then
        Err.Raise lngAppErrorNo, cstrFunctionName, "Application is already locked"
    Else
        Err.Raise lngAppErrorNo, cstrFunctionName, "Cannot determine application lock status"
    End If

    xmlSetAttributeValue vxmlResponseNode, "TYPE", "SUCCESS"
    xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "SUCCESS"

    gobjContext.SetComplete
    
UnlockApplication:

    ' PSC 09/05/2006 MAR1643 - Start
    If blnCaseSuccessfullyLocked Then
        On Error Resume Next
        
        ' Try to unlock application
        objNTTxBO.UnlockApplication vxmlRequestNode.cloneNode(False), strAppNo
        
        lngErrorNo = Err.Number
        strSource = Err.Source
        strDescription = Err.Description
        
        On Error GoTo PerformTASMoveToNextStageProcessingError
        
        If lngErrorNo <> 0 Then
            App.LogEvent "Unable to unlock application " & strAppNo & ". " & strDescription, vbLogEventTypeWarning
            
            ' If error was caused during unlockin operation of TAS processing raise an error
            If intCurrentProgress = TASPROGRESS.UnLockingApplication Then
                Err.Raise lngErrorNo, strSource, strDescription
            End If
        End If
    End If
    ' PSC 09/05/2006 MAR1643 - End
    
PerformTASMoveToNextStageProcessingExit:
            
    Set objNoTxBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlApplication = Nothing
    Set xmlTempError = Nothing
    Set xmlCurrentStageIn = Nothing
    Set xmlCurrentStage = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlDescription = Nothing
    Set xmlSource = Nothing
    Set objNTTxBO = Nothing     ' PSC 09/05/2006 MAR1643

    Exit Sub
        
PerformTASMoveToNextStageProcessingError:
    
    gobjContext.SetAbort
    Set xmlTempError = CreateErrorResponseNode()
    xmlCopyAttribute xmlTempError, vxmlResponseNode, "TYPE"
    vxmlResponseNode.appendChild xmlTempError.firstChild
    
    Set xmlDescription = vxmlResponseNode.selectSingleNode("ERROR/DESCRIPTION")
    Set xmlSource = vxmlResponseNode.selectSingleNode("ERROR/SOURCE")
    
    Select Case intCurrentProgress
        Case TASPROGRESS.ValidateRequest
            strMessage = "The supplied request is invalid."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"
        Case TASPROGRESS.GettingLockStatus
            strMessage = "Unable to determine if the application is already locked."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "MOVESTAGERETRYAPPLICATIONLOCK"
        Case TASPROGRESS.LockingApplication
            strMessage = "Unable to lock application."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "MOVESTAGERETRYAPPLICATIONLOCK"
        Case TASPROGRESS.MovingStage
            strMessage = "Unable to move to the next stage."
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "MOVESTAGERETRY"
        Case TASPROGRESS.UnLockingApplication
            strMessage = "Unable to unlock application."
            'PSC 09/05/2006 MAR1643
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "WARNING"
        Case Else
            xmlSetAttributeValue vxmlResponseNode, "TASSTATUS", "FAILURE"
    End Select
    
    xmlDescription.Text = strMessage & " " & xmlDescription.Text
    xmlSource.Text = cstrFunctionName & "." & xmlSource.Text
    
    'PSC 09/05/2006 MAR1643 - Start
    ' If error occurred  before unlock stage unlock the application else cleanup
    If intCurrentProgress <> TASPROGRESS.UnLockingApplication Then
        Resume UnlockApplication
    Else
        Resume PerformTASMoveToNextStageProcessingExit
    End If
    'PSC 09/05/2006 MAR1643 - End
    
End Sub
'PSC 25/08/2005 MAR32 - End

'MAR1787 Add sub to get Mandatory flag.
Private Function GetTaskMandatoryFlag(ByVal vstrTaskId As String, ByVal vstrStageId As String) As String
    
    Const cstrFunction   As String = "GetTaskMandatoryFlag"
    Dim cmd              As ADODB.Command
    Dim rst              As ADODB.Recordset
    Dim strSQL           As String
    Dim strMandatoryFlag As String

On Error GoTo GetTaskMandatoryFlagVbErr

    strMandatoryFlag = "0"

    strSQL = "SELECT MANDATORYFLAG FROM STAGETASK WHERE TASKID = '" & vstrTaskId & _
                                                 "' AND STAGEID = '" & vstrStageId & "'"
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    
    If Not rst Is Nothing Then
        If Not (IsNull(rst.Fields(0).Value)) Then
            If CSafeBool(rst.Fields(0).Value) Then
                strMandatoryFlag = "1"
            End If
        End If
    End If

    GetTaskMandatoryFlag = strMandatoryFlag

GetTaskMandatoryFlagVbErr:
    Set cmd = Nothing
    Set rst = Nothing
    errCheckError cstrFunction, TypeName(Me)
End Function




