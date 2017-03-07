VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OmTmBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OmTmBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Task Manager Business Object.

'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    04/07/2005  MAR5 - KFI Integration
'GHun   29/07/2005  MAR7 Integrate local printing and apply CORE174
'PSC    24/08/2005  MAR32 Task Automation Changes
'AM     30/09/2005  MAR57 Added REGULATIONINDICATOR to CreateApplication
'PSC    03/10/2005  MAR32 Add printer to CreateAdHocCaseTask
'MV     03/10/2005  MAR35 Amended IssueOffer();CreateDefaultApplicationConditions(); CreateLegalRepTask()
'MV     07/10/2005  MAR128  Amended CreateLegalRepTask() Core Review Fixes
'MV     10/10/2005  MAR22   Amended omTmRequest and Added RunFTInterfaceCalls()
'MV     13/10/2005  MAR22   Amended IssueOffer();
'MV     17/10/2005  MAR182  Created RunFirstTitleFND(), RunFirstTitleMSFTOE() ,RunFirstTitleTTL() ,GetApplicationFirstTitle()
'                           Amended omTMRequest()
'MV     17/10/2005  MAR182  Created omTMRequest(),GetTaskStatusFromInterfaceMessage(),GetRecordsAsXMLForSQL(),
'                           HandleInterfaceResponse()
'DRC    18/10/2005  MAR65   Changes to Critical Data functions for Rescore creditcheck
'HM     18/10/2005  MAR41   Amended omTmRequest and Added RunHomeTrackValuation()
'MV     19/10/2005  MAR231  Amended IssueOffer()
'GHun   20/10/2005  MAR184  Change omTmRequest and added RunXMLCreditCheck
'JD     20/10/2005  MAR40   Added RunESurvValuation interface method.
'HMA    11/10/2005  MAR23   Add KnowYourCustomer
'SR     21/10/2005  MAR268  new method 'RunHunterExperianInterface'
'GHun   22/10/2005  MAR231  Changed KnowYourCustomer
'IK     26/10/2005  MAR232  HandleInterfaceResponse errors:
'                           should not be dependant on TASKINTERFACESUBMESSAGE nodes
'                           call to CreateAdHocCaseTask needs APPLICATION node
'                           call to CreateAdHocCaseTask needs APPLICATIONPRIORITY on APPLICATION node
'                           error in GetTaskStatusFromInterfaceMessage - test for null value
'PSC    26/10/2005  MAR284  Amend FindCaseTaskListLite to correct SLA search
'MV     27/10/2005  MAR281  Amended KnowYourCustomer()
'SAB    28/10/2005  MAR245  Amended RunXMLCreditCheck()
'HM     27/10/2005  MAR263  Amend RunHomeTrackValuation
'HM     27/10/2005  MAR316  New method RunDelayedCompletionInterface
'JD     28/10/2005  MAR342  Fixes to HandleInterfaceResponse
'MV     02/11/2005  MAR182  Amended RunFTInterfaceCall()
'MV     03/11/2005  MAR204  Amended RunFirstTitleFND(); RunFirstTitleTTL();HandleInterfaceResponse()
'TW     03/11/2005  MAR211  ProcessTask for PACK handling
'SR     04/11/2005  MAR329  Add RunPinNumberTriggerInterface and RunLegalAddressTriggerInterface
'SR     04/11/2005  MAR269  Amend RunRiskassessment
'PSC    04/11/2005  MAR329  Amend RunRiskassessment
'HMA    07/11/2005  MAR339  Added error handling to RunHunterExperianInterface
'GHun   09/11/2005  MAR182  Changed RunFTInterfaceCall and GetApplicationFirstTitle
'JD     11/11/2005  MAR503  Amended RunHometrackValuation to use passed in customers.
'IK     11/11/2005  MAR452  Amend RunXMLCreditCheck, RunRiskassessment
'DRC    14/11/2005  MAR329  Amend RunRiskassessment need CaseActivity if running from Task Manager
'INR    15/11/2005  MAR533 RETURNCODE for RunHunterExperianInterface is returned as mixed case
'INR    15/11/2005  MAR563 rename GenerateEMailFulfillmentExtract & GenerateSMSFulfillmentExtract, too long
'DRC    16/11/2005  MAR33 & 329 RunRiskassessment needed call to HandleInterfaceResponse
'DRC    16/11/2005  MAR599  No riskassessment for LIM rescore
'DRC    16/11/2005  MAR512  Use default Omiga CaseActivity in HandleInterfaceResponse
'HMA    16/11/2005  MAR582  Add RUNFTREISSUEOFFERINTERFACE
'GHun   16/11/2005  MAR467  Changed KnowYourCustomer
'GHun   16/11/2005  MAR584  Changed RunHomeTrackValuation
'GHun   16/11/2005  MAR461  Changed IssueOffer to fix auto approval
'HMA    17/11/2005  MAR467  Changed KnowYourCustomer
'HMA    17/11/2005  MAR630  Change RunXMLCreditCheck
'GHun   17/11/2005  MAR512  Changed GetTaskStatusFromInterfaceMessage
'HMA    21/11/2005  MAR630  Further changes for RunXMLCreditCheck/RunRiskAssessment
'HMA    24/11/2005  MAR630  Changes to HandleInterfaceResponse
'HMA    24/11/2005  MAR664  Changes to HandleInterfaceResponse and RunHomeTrackValuation for Task Note
'GHun   28/11/2005  MAR609  Changed HandleInterfaceResponse
'HMA    28/11/2005  MAR712  Changes to HandleInterfaceResponse and GetRecordsAsXMLForSQL
'INR    29/11/2005  MAR731  Implement RunPinNumberTriggerInterface and RunLegalAddressTriggerInterface
'DRC    30/11/2005  MAR727  Critical Data call without database update
'GHun   30/11/2005  MAR609  Changed RunFirstTitleMSFTOE and omTmRequest
'GHun   01/12/2005  MAR727  Changed CriticalDataCheck
'PSC    01/12/2005  MAR770  Amend CreateKYCTask to use copies of nodes
'HMA    01/12/2005  MAR750  Changed RunDelayedCompletionInterface, RunPinNumberTriggerInterface,
'                           RunLegalAddressTriggerInterface and KnowYourCustomer
'GHun   01/12/2005  MAR780  Change IssueOffer to complete the task
'PSC    02/12/2005  MAR770  Amend CreateKYCTask to copy nodes earlier
'HMA    08/12/2005  MAR750  Changed RunPinNumberTriggerInterface and RunLegalAddressTriggerInterface
'GHun   12/12/2005  MAR847  Changed GetTaskStatusFromInterfaceMessage and HandleInterfaceResponse
'HMA    12/12/2005  MAR853  Changed RunPinNumberTriggerInterface, RunLegalAddressTriggerInterface and RunDelayedCompletionInterface
'DRC    12/12/2005  MAR856  Get the task names that critical data can create from global parameters
'PJO    14/12 2005  MAR850  Set Direct / Infirect value on Create Application
'GHun   14/12/2005  MAR865  Changed IssueOffer
'GHun   14/12/2005  MAR865  Changed IssueOffer (again)
'PSC    14/12/2005  MAR802  Send customer id to omPack for SMSFulfilmentExtract and EmailFulfilmentExtract
'HMA    15/12/2005  MAR841  Changed RunXMLCreditCheck for Reprocess and Rescore Credit Check
'BC     14/12/2005  MAR842  Incorrect transactioning when calling GetpplicationData in omTMBO/CriticalDataChaeck
'GHun   15/12/2005  MAR881  Changed RunFirstTitleTTL
'DRC    15/12/2005  MAR856  Revisited - needed to move the Global param call for DAwebtop comparison
'DRC    15/12/2005  MAR856  Revisited again - needed to move the Global param call for Update only
'PE     16/12/2005  MAR898  Modified RunFirstTitleFND/RunFirstTitleTTL to allow for
'                           GetApplicationFirstTitle raising an error when record not found.
'HMA    19/12/2005  MAR550  Modifed IssueOffer to check risk assessment rules, strategy manager reasons
'                           and KYC status
'GHun   16/01/2006  MAR972  Changed HandleInterfaceResponse to create tasks in a single transaction
'PSC    18/01/2006  MAR1074 Amend CheckTasksOnCurrentStage to use GlobalParameter for credit score task
'HMA    18/01/2006  MAR1048 Changed RunFirstTitleFND
'GHun   19/01/2006  MAR1088 Added RunAutoApplicationDecline
'GHun   24/01/2006  MAR1084 Changed RunXMLCreditCheck to raise an error if address targetting is not supported
'GHun   24/01/2006  MAR885  Changed CreateAdhocCaseTask
'HMA    25/01/2006  MAR1110 Changed RunFirstTitleFND to create task note
'HMA    26/01/2006  MAR1126 Changes to CriticalDataCheck
'HMA    27/01/2006  MAR1137 Changed RunFirstTitleFND to set up amount in DisbursementPayment record
'DRC    30/01/2006  MAR1142 Changes to RunRiskAssessment - no Message Subtypes for Decline Tasks
'HMA    30/01/2006  MAR1110 Changes to RunFirstTitleFND for Task Notes
'GHun   31/01/2006  MAR1173 Changed RunRiskAssessment to return DECISION as it did before
'DRC    01/02/2006  MAR1172 Added function InitialCreditCheckDone to see if a credit check has been run yet - used in critical data check
'PSC    03/02/2006  MAR1156 Amend process task to only do Pack Processing if there is no interface specified
'                           Amend IssueOffer to perform pack processing
'RF     08/02/2006  MAR1222 The MSI (NOR) message is generated in the system, with the error message ACA is
'                           not found and the reason 'Error searching for ACA First Title Message Record'
'JD     09/02/2006  MAR1220 in RunFirstTitleFND - when creating the disbursementpayment set the completiondate field to the calculated advance date.
'RF     11/02/2006  MAR1250 Hometrack and other interfaces don't run when submitted from web
'PSC    13/02/2006  MAR1250 Correct Hometrack
'HMA    13/02/2006  MAR1244 Changes for ProcessTask and RunPinNumberTriggerInterface
'HMA    13/02/2006  MAR1122 Changes to RunLegalAddressTriggerInterface
'RF     16/02/2006  MAR1251 Allow template destination of DMS only
'DRC    28/02/2006  MAR1323 Rationalise calls into CDRules
'JD     28/02/2006  MAR1188 add more detail to error message 4804/4805
'HMA    01/03/2006  MAR1265 Changed RunFirstTitleFND
'PSC    02/03/2006  MAR1341 Amend UpdateCaseTaskPriority to take into account ADJUSTMENTMINUTES
'GHun   06/03/2006  MAR1143 Changed RunRiskAssessment
'DRC    28/02/2006  MAR1323 Backed out
'DRC    08/03/2006  MAR1323 Put back in again
'GHun   10/03/2006  MAR1300 Changed ProcessTask to call ValidateProcessTaskAuthority
'RF     15/03/2006  MAR1392 Writing MOF message to ApplicationFirstTitle table is timing out
'JD     17/03/2006  MAR1396 CreateValuationReportNoInst - copy some values from the previous valuerinstruction
'AS     22/03/2006  MAR1501 OmTMBO does not set transaction to complete.
'JD     24/03/2006  MAR1434 expose getmaxcasetaskinstance as public function
'                           Call omRB to pass data to valuationRules in ValidateValuationReport
'JD     28/03/2006  MAR1533 IssueOffer -  check for tasks in combos not globals
'DRC    29/03/2006  MAR1477 KnowYourCustomer - Don't set customerlock in call to getandsyncronise
'AS     31/03/2006  MAR1507 PERFORMANCE - SetComplete is done in the wrong place for KnowYourCustomer.
'BC     31/03/2006  MAR1549 Run USP_DELETEUNMODIFIEDCONDITIONS before CREATEDEFAULTAPPLICATIONCONDITIONS in IssueOffer function
'GHun   04/04/2006  MAR1300 Added SetChangeOfProperty
'GHun   05/04/2006  MAR1300 Changed SetChangeOfProperty
'BC     06/04/2006  MAR1425 Remove call to USP_DELETEUNMODIFIEDCONDITIONS. Processing now done in CondtionsTxBO
'GHun   10/04/2006  MAR1599 Change IssueOffer to set offer expiry task user authority level
'GHun   11/04/2006  MAR1607 Changed SetChangeOfProperty to raise an error if NewProperty does not exist
'HMA    19/04/2006  MAR1527 Added new interface RunFirstTitleACA
'JD     20/04/2006  MAR1615 RunRiskAssessment - get FRAUDCOUNT correctly
'JD     20/04/2006  MAR1567 HandleInterfaceResponse - if there are no UW refers set flag accordingly
'GHun   20/04/2006  MAR1632 Changed ProcessTask to pass ChannelId
'GHun   20/04/2006  MAR1646 Changed RunFirstTitleFND to remove DIR checking
'GHun   20/04/2006  MAR1646 Changed RunFirstTitleFND to prevent processing multiple FNDs
'JD     30/04/2006  MAR1617 added interface RUNXMLEXPUPDATELIM
'PSC    05/05/2006  MAR1593 Add RUNSMSEMAILINTERFACE
'HMA    08/05/2006  MAR1259 Change to First Title task production in Issue Offer.
'GHun   08/05/2006  MAR1717 Changed InitialCreditCheckDone & CriticalDataCheck
'GHun   09/05/2006  MAR1759 Changed CriticalDataCheck
'GHun   10/05/2006  MAR1715 Changed HandleInterfaceResponse and RunFirstTitleFND and added GetApplicationFirstTitleFND
'GHun   11/05/2006  MAR1715 Renamed GetApplicationFirstTitleFND to GetApplicationFirstTitleFromId
'-------------------------------------------------------------------------------------------------------
'EPSOM Specific History:
'
'IK     17/032006   EP2     changes for epsom ingestion
'SAB    20/03/2006  EP17    Added RunRealHunterExperianInterface
'LDM    23/03/2006  EP6     New interface to experian. RunXMLExpQuotationConversionCC
'LDM    04/04/2006  EP16    New procedure for epsom KYC check(RunEpsomKYC)
'LDM    11/04/2006  EP377   New interfaces for epsom for credit score reprocess and rescore
'IK     26/04/2006  EP465   error in HandleInterfaceResponse when running DIP Risk Assessment
'IK     12/05/2006  EP510   merge MARS defects, add compile flag
'IK     18/05/2006  EP574   add BuildHandoverDoc interface
'IK     31/05/2006  EP626   fix automation error
'IK     18/05/2006  EP672   BuildHandoverDoc interface, parameterise .xslt file
'PB     29/06/2006  EP435   Do not select 'Default' valuer
'PE     25/07/2006  EP974   Mars merge/MAR1848 - TAS deadlocking issue
'AW     28/09/06    EP1135  Set TASKDUEDATEANDTIME in CreateCaseStageTrigger
'AW     05/10/06    EP1196  Copy valuer details from previous instruction instance
'------------------------------------------------------------------------------------------
'EPSOM Phase 2 Specific History:
'
'PE     09/11/2006  EP2_67  Xit2 Interface - New Xit2 valuation tasks
'PE     04/01/2007  EP2_682 omTM - HandleInterfaceResponse - Attempts to process deleted tasks
'IK     20/01/2007  EP2_902 AddDefaultValuesToCaseStage moved to omTmFunctions.bas
'AShaw  13/02/2007  EP2_1293 - Check ISLEGALREPTOBEUSED flag in IssueOffer.
'IK     16/03/2007  EP2_1433 - printer destination via combovalidation look-up
'GHun   15/03/2007  EP2_1903 Added HandleInboundDocument and changed PrintDocumentForTask
'AW     19/03/07    EP2_1897  Save Accepted quote number on ApplicationOffer table
'PSC    21/03/2007  EP2_1619 Amend RunEpsomCaseAssess to create correct referal task
'GHun   24/03/2007  EP2_2262 Changed ValidateReportOnTitle
'IK     11/04/2007  EP2_2264 HandleInterfaceResponse only when tasks identified
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements MessageQueueComponentVC2

Private gobjContext As ObjectContext
Private gobjTmBO As Object
Private Const cstrTaskComplete As String = "40"

'MAR563 Interface name too long - renamed
Private Sub EMailFulfilmentExtract(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' TW 03/11/2005 MAR211
Const strFunctionName As String = "EMailFulfilmentExtract"

Dim xmlApplicationNode As IXMLDOMNode

Dim xmlTempResponse As FreeThreadedDOMDocument40

Dim strResponse As String

Dim xmlPackAttribute As IXMLDOMAttribute
Dim xmlPackRequestDoc As FreeThreadedDOMDocument40
Dim xmlPackRequestNode As IXMLDOMNode
Dim gobjPack As Object
   
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlCaseTaskNode As IXMLDOMNode
Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
Dim xmlThisCaseTaskElem As IXMLDOMElement
Dim xmlThisCaseTaskNode As IXMLDOMNode

'PSC 05/05/2006 MAR1593 - Start
Dim blnBypassTaskUpdate As Boolean
Dim strTaskId As String
Dim strPackId As String
Dim strValueId As String
Dim intIndex As Integer
Dim colValueIDs As Collection
'PSC 05/05/2006 MAR1593 - End

    On Error GoTo EMailFulfilmentExtractExit:
    
    Set xmlPackRequestDoc = New FreeThreadedDOMDocument40
    xmlPackRequestDoc.async = False
    xmlPackRequestDoc.setProperty "NewParser", True
        
    Set xmlApplicationNode = vxmlRequest.selectSingleNode("APPLICATION")
    
    ' PSC 14/12/2005 MAR802
    Set xmlCaseTaskNode = vxmlRequest.selectSingleNode("./CASETASK")

    'PSC 05/05/2006 MAR1593 - Start
    strTaskId = xmlGetMandatoryAttributeText(xmlCaseTaskNode, "TASKID")
    Set colValueIDs = New Collection
    GetValueIdsForValidationType "EmailSMSPackType", strTaskId, colValueIDs
    
    intIndex = 1
    
    While intIndex <= colValueIDs.Count And Len(strPackId) = 0
        strValueId = colValueIDs.Item(intIndex)
        
        If IsValidationType("EmailSMSPackType", strValueId, "E") Then
            strPackId = GetComboText("EmailSMSPackType", strValueId)
        End If

        intIndex = intIndex + 1
    Wend
    
    If Len(strPackId) = 0 Then
        errThrowError strFunctionName, 4702, "EmailSMSPackType does not contain an SMS entry for taskid " & strTaskId
    End If
    'PSC 05/05/2006 MAR1593 - End

    Set xmlPackRequestNode = xmlPackRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
    xmlPackRequestDoc.appendChild xmlPackRequestNode
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("OPERATION")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = "SENDPACKTOEMAIL"
    
    'PSC 05/05/2006 MAR1593
    xmlSetAttributeValue xmlPackRequestNode, "PACKCONTROLNUMBER", strPackId
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("APPLICATIONNUMBER")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("APPLICATIONFACTFINDNUMBER")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    ' PSC 14/12/2005 MAR802
    xmlCopyAttributeValue xmlCaseTaskNode, xmlPackRequestNode, "CUSTOMERIDENTIFIER", "CUSTOMERNUMBER"
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("UNITID")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "UNITID")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("USERID")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "USERID")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("USERAUTHORITYLEVEL")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "USERAUTHORITYLEVEL")

    On Error Resume Next
    Set gobjPack = gobjContext.CreateInstance("omPack.PackManagerBO")
    On Error GoTo EMailFulfilmentExtractExit:

    If Not gobjPack Is Nothing Then
        strResponse = gobjPack.omRequest(xmlPackRequestNode.xml)
        Set xmlTempResponse = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        ' PSC 05/05/2006 MAR1593
        blnBypassTaskUpdate = xmlGetAttributeAsBoolean(xmlCaseTaskNode, "BYPASSTASKSTATUSUPDATE")

        ' PSC 05/05/2006 MAR1593 - Start
        If blnBypassTaskUpdate = False Then
            Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
            xmlThisCaseTaskDoc.setProperty "NewParser", True
            xmlThisCaseTaskDoc.async = False
            
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "UpdateCaseTask"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "UNITID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERAUTHORITYLEVEL"
            Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True) ' PSC 05/05/2006 MAR1593
            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "40"
            xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
            
            Set xmlOut = New FreeThreadedDOMDocument40
            xmlOut.async = False
            xmlOut.setProperty "NewParser", True
            
            ' run Tm UpdateCaseTask request
            RunTmRequest xmlThisCaseTaskDoc, xmlOut
        
            Debug.Print "UpdateCaseTask response:"
            Debug.Print xmlOut.xml
            ' parse response
            errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True
        End If
        ' PSC 05/05/2006 MAR1593 - End
    End If
    
EMailFulfilmentExtractExit:

    Set xmlApplicationNode = Nothing
    
    Set xmlTempResponse = Nothing

    Set xmlPackAttribute = Nothing
    Set xmlPackRequestDoc = Nothing
    Set xmlPackRequestNode = Nothing

    Set xmlOut = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing

    errCheckError strFunctionName
' TW 03/11/2005 MAR211 End

End Sub
       
'MAR563 Interface name too long - renamed
Private Sub SMSFulfilmentExtract(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' TW 03/11/2005 MAR211
Const strFunctionName As String = "SMSFulfilmentExtract"

Dim xmlApplicationNode As IXMLDOMNode
Dim xmlTempResponse As FreeThreadedDOMDocument40

Dim strResponse As String

Dim xmlPackAttribute As IXMLDOMAttribute
Dim xmlPackRequestDoc As FreeThreadedDOMDocument40
Dim xmlPackRequestNode As IXMLDOMNode
Dim gobjPack As Object
   
Dim xmlOut As FreeThreadedDOMDocument40
Dim xmlCaseTaskNode As IXMLDOMNode
Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
Dim xmlThisCaseTaskElem As IXMLDOMElement
Dim xmlThisCaseTaskNode As IXMLDOMNode

'PSC 05/05/2006 MAR1593 - Start
Dim blnBypassTaskUpdate As Boolean
Dim strTaskId As String
Dim strPackId As String
Dim strValueId As String
Dim intIndex As Integer
Dim colValueIDs As Collection
'PSC 05/05/2006 MAR1593 - End
 
    On Error GoTo SMSFulfilmentExtractExit:
    
    Set xmlPackRequestDoc = New FreeThreadedDOMDocument40
    xmlPackRequestDoc.async = False
    xmlPackRequestDoc.setProperty "NewParser", True
        
    Set xmlApplicationNode = vxmlRequest.selectSingleNode("APPLICATION")
    
    ' PSC 14/12/2005 MAR802
    Set xmlCaseTaskNode = vxmlRequest.selectSingleNode("./CASETASK")
    
    'PSC 05/05/2006 MAR1593 - Start
    strTaskId = xmlGetMandatoryAttributeText(xmlCaseTaskNode, "TASKID")
    Set colValueIDs = New Collection
    GetValueIdsForValidationType "EmailSMSPackType", strTaskId, colValueIDs
    
    intIndex = 1
    
    While intIndex <= colValueIDs.Count And Len(strPackId) = 0
        strValueId = colValueIDs.Item(intIndex)
        
        If IsValidationType("EmailSMSPackType", strValueId, "S") Then
            strPackId = GetComboText("EmailSMSPackType", strValueId)
        End If

        intIndex = intIndex + 1
    Wend
    
    If Len(strPackId) = 0 Then
        errThrowError strFunctionName, 4702, "EmailSMSPackType does not contain an SMS entry for taskid " & strTaskId
    End If
    'PSC 05/05/2006 MAR1593 - End
    
    Set xmlPackRequestNode = xmlPackRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
    xmlPackRequestDoc.appendChild xmlPackRequestNode
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("OPERATION")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = "SENDPACKTOSMS"
    
    'PSC 05/05/2006 MAR1593
    xmlSetAttributeValue xmlPackRequestNode, "PACKCONTROLNUMBER", strPackId
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("APPLICATIONNUMBER")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("APPLICATIONFACTFINDNUMBER")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    ' PSC 14/12/2005 MAR802
    xmlCopyAttributeValue xmlCaseTaskNode, xmlPackRequestNode, "CUSTOMERIDENTIFIER", "CUSTOMERNUMBER"
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("UNITID")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "UNITID")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("USERID")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "USERID")
    
    Set xmlPackAttribute = xmlPackRequestDoc.createAttribute("USERAUTHORITYLEVEL")
    xmlPackRequestNode.Attributes.setNamedItem xmlPackAttribute
    xmlPackAttribute.Text = xmlGetAttributeText(vxmlRequest, "USERAUTHORITYLEVEL")

    On Error Resume Next
    Set gobjPack = gobjContext.CreateInstance("omPack.PackManagerBO")
    On Error GoTo SMSFulfilmentExtractExit:

    If Not gobjPack Is Nothing Then
        strResponse = gobjPack.omRequest(xmlPackRequestNode.xml)
        Set xmlTempResponse = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        ' PSC 05/05/2006 MAR1593
        blnBypassTaskUpdate = xmlGetAttributeAsBoolean(xmlCaseTaskNode, "BYPASSTASKSTATUSUPDATE")
        
        ' PSC 05/05/2006 MAR1593 - Start
        If blnBypassTaskUpdate = False Then
            Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
            xmlThisCaseTaskDoc.setProperty "NewParser", True
            xmlThisCaseTaskDoc.async = False
            
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "UpdateCaseTask"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "UNITID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERAUTHORITYLEVEL"
            Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True) ' PSC 05/05/2006 MAR1593
            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "40"
            xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
            
            Set xmlOut = New FreeThreadedDOMDocument40
            xmlOut.async = False
            xmlOut.setProperty "NewParser", True
            
            ' run Tm UpdateCaseTask request
            RunTmRequest xmlThisCaseTaskDoc, xmlOut
        
            Debug.Print "UpdateCaseTask response:"
            Debug.Print xmlOut.xml
            ' parse response
            errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True
        End If
        ' PSC 05/05/2006 MAR1593 - End

    End If
    
SMSFulfilmentExtractExit:

    Set xmlApplicationNode = Nothing
    
    Set xmlTempResponse = Nothing

    Set xmlPackAttribute = Nothing
    Set xmlPackRequestDoc = Nothing
    Set xmlPackRequestNode = Nothing

    Set xmlOut = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing

    errCheckError strFunctionName
' TW 03/11/2005 MAR211 End
End Sub

Private Sub GetApplicationFirstTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetApplicationFirstTitleVbErr

    Const cstrFunctionName As String = "GetApplicationFirstTitle"
    
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlAppNode As IXMLDOMNode
    
    Dim strData As String
    Dim strAppNumber As String
    Dim strAppFactFindNo As String
    Dim strMsgType As String
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    
    Set xmlAppNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    strAppNumber = xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONNUMBER")
    strAppFactFindNo = xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONFACTFINDNUMBER")
    strMsgType = xmlGetMandatoryAttributeText(xmlAppNode, "MESSAGETYPE")
    
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.Open
    
    With adoCommand
        .CommandText = "USP_GETLATESTAPPLICATIONFIRSTTITLE"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
    End With
    
    Set adoParameter = adoCommand.CreateParameter("@pApplicationNumber", adBSTR, adParamInput, Len(strAppNumber), strAppNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@pApplicationFactFindNumber", adInteger, adParamInput, , strAppFactFindNo)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@pMessageType", adBSTR, adParamInput, Len(strMsgType), strMsgType)
    adoCommand.Parameters.Append adoParameter
    
    With adoCommand
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = xmlLoad(strData, cstrFunctionName)
        
    Dim xmlAppFTNode As IXMLDOMNode
    Set xmlAppFTNode = xmlGetNode(xmlOut.documentElement, "APPLICATIONFIRSTTITLE")
    
    If (xmlAppFTNode Is Nothing) Then
        'MAR182 GHun
        errThrowError cstrFunctionName, oeRecordNotFound
        'xmlSetAttributeValue xmlOut.documentElement, "TYPE", "APPERR"
    Else
        xmlSetAttributeValue xmlOut.documentElement, "TYPE", "SUCCESS"
    End If
    
    Set vxmlResponseNode = xmlOut.documentElement
    
GetApplicationFirstTitleExit:
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
    
    Exit Sub
    
GetApplicationFirstTitleVbErr:
    
    gobjContext.SetComplete
    Set vxmlResponseNode = CreateErrorResponseNode
    GoTo GetApplicationFirstTitleExit
    
End Sub

Private Sub RunFirstTitleFND(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo RunFirstTitleFNDErr
    
    Const cstrFunctionName As String = "RunFirstTitleFND"
    
    Dim xmlFirstTitleRoot As IXMLDOMNode
    Dim AppFirstTitleNode  As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc  As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlRequest  As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlAdvanceDate   As IXMLDOMNode
    Dim xmlPaymentMethod  As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlROTNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode          ' MAR1048
    Dim xmlDisbursementNode As IXMLDOMNode  ' MAR1048
    Dim xmlSaveAppNode As IXMLDOMNode       ' MAR1048
    Dim xmlPaymentSummary As IXMLDOMNode    ' MAR1048
    Dim xmlPayeeList As IXMLDOMNodeList     ' MAR1048
    Dim xmlThirdParty As IXMLDOMNode        ' MAR1048
    Dim strBalance As String                ' MAR1048
    Dim strApplicationNumber As String      ' MAR1048
    Dim strApplicationFFNumber As String    ' MAR1048
    Dim strLegalRep As String               ' MAR1048
    Dim strSequenceNo As String             ' MAR1048
    Dim strTaskNote As String               ' MAR1110
    Dim lngOmigaError As Long               ' MAR1110
    Dim strNetAmount As String              ' MAR1137
    Dim strAmount As String                 ' MAR1137
    Dim xmlFeeList As IXMLDOMNodeList       ' MAR1137
    Dim xmlFee As IXMLDOMNode               ' MAR1137
    
    Dim objNoTxBO  As Object
    Dim objPayProc As Object
    Dim objROTBO  As Object
    Dim objMsgTmBO As Object
    Dim objComboDO As Object
    
    Dim strFNDCompletionDate  As String
    Dim strResponse As String
    Dim strROTGUID  As String
    Dim strAdvanceDate As String
    Dim strPaymentMethod  As String
    Dim strUserId As String
    Dim strUnitId  As String
    Dim strUserAuthorityLevel  As String
    Dim strTMPreCompletionStage As String
    Dim strCurrentStage As String                   ' MAR1265
    
    Dim dtACADate As Date
    'Dim dtDIRDate As Date
    Dim dtMOFDate As Date
    'Dim dtFNDCompletionDate  As Date
    
    Dim lngErrorNumber As Long
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlTempResponse = New FreeThreadedDOMDocument40
    
    Set xmlRequest = vxmlRequestNode.selectSingleNode("//REQUEST")
    Set xmlCaseTaskIn = vxmlRequestNode.selectSingleNode("//CASETASK")
    Set xmlAppNode = vxmlRequestNode.selectSingleNode("//APPLICATION")
    Set xmlSaveAppNode = vxmlRequestNode.selectSingleNode("//APPLICATION").cloneNode(True)
    
    strApplicationNumber = xmlGetAttributeText(xmlAppNode, "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(xmlAppNode, "APPLICATIONFACTFINDNUMBER")
    
    strCurrentStage = xmlGetAttributeText(xmlCaseTaskIn, "STAGEID")   ' MAR1265
    
    'MAR1048  Correct error handling throughout
    'If any error occurs, generate the Funds Release Failed task
    
    'Get ApplicationFirstTitle record for MessageType "ACA"
    xmlSetAttributeValue xmlAppNode, "MESSAGETYPE", "ACA"
    
    On Error Resume Next
    GetApplicationFirstTitle vxmlRequestNode, xmlFirstTitleRoot
    On Error GoTo RunFirstTitleFNDErr
        
    lngErrorNumber = errCheckXMLResponseNode(xmlFirstTitleRoot, , False)
    
    If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
        strTaskNote = errGetMessageText(8541)
        'strTaskNote = "Missing ACA First Title Message Record"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errThrowError cstrFunctionName, oeRecordNotFound, strTaskNote
    ElseIf lngErrorNumber <> 0 Then
        strTaskNote = "Error searching for ACA First Title Message Record"
        'RF 08/02/2006 MAR1222 Start - Improve error message
        If Not xmlFirstTitleRoot Is Nothing Then
            strTaskNote = strTaskNote & ". Details: " & xmlFirstTitleRoot.xml
        End If
        'RF 08/02/2006 MAR1222 End
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errRaiseXMLResponseNode xmlFirstTitleRoot
    Else
        'Get the Date
        Set AppFirstTitleNode = xmlFirstTitleRoot.selectSingleNode("//APPLICATIONFIRSTTITLE")
        dtACADate = CSafeDate(xmlGetAttributeText(AppFirstTitleNode, "DATETIME"))
    End If
     
    'MAR1646 GHun Commented out DIR checking
    ''Get ApplicationFirstTitle record for MessageType "DIR"
    'xmlSetAttributeValue xmlAppNode, "MESSAGETYPE", "DIR"
    '
    'On Error Resume Next
    'GetApplicationFirstTitle vxmlRequestNode, xmlFirstTitleRoot
    'On Error GoTo RunFirstTitleFNDErr
    '
    'lngErrorNumber = errCheckXMLResponseNode(xmlFirstTitleRoot, , False)
    '
    'If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
    '    strTaskNote = errGetMessageText(8542)
    '    'strTaskNote = "Missing DIR First Title Message Record"
    '    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
    '    errThrowError cstrFunctionName, oeRecordNotFound, strTaskNote
    'ElseIf lngErrorNumber <> 0 Then
    '    strTaskNote = "Error searching for DIR First Title Message Record"
    '    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
    '    errRaiseXMLResponseNode xmlFirstTitleRoot
    'Else
    '    'Get the completionDate
    '    Set AppFirstTitleNode = xmlFirstTitleRoot.selectSingleNode("//APPLICATIONFIRSTTITLE")
    '    dtDIRDate = CSafeDate(xmlGetAttributeText(AppFirstTitleNode, "DATETIME"))
    'End If
    'MAR1646

    'Get ApplicationFirstTitle record for MessageType "MOF"
    xmlSetAttributeValue xmlAppNode, "MESSAGETYPE", "MOF"
    
    On Error Resume Next
    GetApplicationFirstTitle vxmlRequestNode, xmlFirstTitleRoot
    On Error GoTo RunFirstTitleFNDErr
    
    lngErrorNumber = errCheckXMLResponseNode(xmlFirstTitleRoot, , False)
        
    If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
        strTaskNote = errGetMessageText(8543)
        'strTaskNote = "Missing MOF First Title Message Record"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errThrowError cstrFunctionName, oeRecordNotFound, strTaskNote
    ElseIf lngErrorNumber <> 0 Then
        strTaskNote = "Error searching for MOF First Title Message Record"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errRaiseXMLResponseNode xmlFirstTitleRoot
    Else
        'Get the completionDate
        Set AppFirstTitleNode = xmlFirstTitleRoot.selectSingleNode("//APPLICATIONFIRSTTITLE")
        dtMOFDate = CSafeDate(xmlGetAttributeText(AppFirstTitleNode, "DATETIME"))
    End If

    'MAR1715 GHun
    'Get ApplicationFirstTitle FND record using ApplicationFirstTitleId from CaseTask.Context
    xmlSetAttributeValue xmlAppNode, "APPLICATIONFIRSTITLEID", xmlGetAttributeText(xmlCaseTaskIn, "CONTEXT")
    
    On Error Resume Next
    GetApplicationFirstTitleFromId vxmlRequestNode, xmlFirstTitleRoot
    'MAR1715 End
    On Error GoTo RunFirstTitleFNDErr
    
    lngErrorNumber = errCheckXMLResponseNode(xmlFirstTitleRoot, , False)
    
    If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
        strTaskNote = errGetMessageText(8544)
        'strTaskNote = "Missing FND First Title Message Record"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errThrowError cstrFunctionName, oeRecordNotFound, strTaskNote
    ElseIf lngErrorNumber <> 0 Then
        strTaskNote = "Error searching for FND First Title Message Record"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errRaiseXMLResponseNode xmlFirstTitleRoot
    Else
        Set AppFirstTitleNode = xmlFirstTitleRoot.selectSingleNode("//APPLICATIONFIRSTTITLE")
        
        'Get the completionDate
        strFNDCompletionDate = xmlGetAttributeText(AppFirstTitleNode, "COMPLETIONDATE")
        'dtFNDCompletionDate = CSafeDate(strFNDCompletionDate)
        strROTGUID = xmlGetAttributeText(AppFirstTitleNode, "ROTGUID")
        
    End If
  
    'MAR1646 GHun
    ''Check the date and time on the DIR and MOF to ensure that they are after the latest ACA
    'If ((dtDIRDate > dtACADate) And (dtMOFDate > dtACADate)) Then
    
    'Check the date and time on the MOF to ensure that it is after the latest ACA
    If (dtMOFDate > dtACADate) Then
    'MAR1646 End
    
        'Validate CompletionDate
        Set xmlNewRequest = vxmlRequestNode.selectSingleNode("//REQUEST").cloneNode(True)
        Set xmlAppNode = xmlNewRequest.selectSingleNode("//APPLICATION")
        
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "VALIDATECOMPLETIONDATE"
        xmlChangeNodeName xmlAppNode, "APPLICATION", "VALIDATECOMPLETIONDATE"
        xmlSetAttributeValue xmlAppNode, "COMPLETIONDATE", strFNDCompletionDate
        
        Set objPayProc = gobjContext.CreateInstance("omPayProc.PaymentProcessingBO")
        
        xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))

        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
        
        If lngErrorNumber = 0 Then
            
           'If Validate CompletionDate is success then ValidateSolicitor
            xmlChangeNodeName xmlAppNode, "VALIDATECOMPLETIONDATE", "SOLICITOR"
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "VALIDATESOLICITOR"
            
            xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))
            
            lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
            
            If lngErrorNumber = 0 Then

                'Get ROTTitleData
                'Find ReportOnTitle Record
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETREPORTONTITlEDATA"
                xmlChangeNodeName xmlAppNode, "SOLICITOR", "REPORTONTITLE"
                'Do not check Completion Date
                xmlAppNode.Attributes.removeNamedItem ("COMPLETIONDATE")
                                
                Set objROTBO = gobjContext.CreateInstance("omROT.omRotBO")
                xmlThisResponseDoc.loadXML (objROTBO.OmRotRequest(xmlNewRequest.xml))
                
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
            
                If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
                        
                    'Create a ROT Record
                    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEREPORTONTITLE"
                    xmlSetAttributeValue xmlAppNode, "COMPLETIONDATE", strFNDCompletionDate
                        
                    xmlThisResponseDoc.loadXML (objROTBO.OmRotRequest(xmlNewRequest.xml))
                    
                    lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                    If (lngErrorNumber <> 0) Then
                        strTaskNote = "Error creating ROT record"
                        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                        errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                    End If
                    
                ElseIf lngErrorNumber <> 0 Then
                        
                    strTaskNote = "Error looking for ROT record"
                    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                    
                Else
                    'Update existing ROT Record
                            
                    Set xmlROTNode = xmlThisResponseDoc.selectSingleNode("//REPORTONTITLE")
                    strROTGUID = xmlGetAttributeText(xmlROTNode, "ROTGUID")
                    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATEREPORTONTITLE"
                    
                    xmlSetAttributeValue xmlAppNode, "COMPLETIONDATE", strFNDCompletionDate
                    xmlSetAttributeValue xmlAppNode, "ROTGUID", strROTGUID
                    
                    xmlThisResponseDoc.loadXML (objROTBO.OmRotRequest(xmlNewRequest.xml))
                    
                    lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                    If (lngErrorNumber <> 0) Then
                        strTaskNote = "Error updating ROT record"
                        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                        errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                    End If
                    
                End If
                
                'SetPayment Method
                xmlChangeNodeName xmlAppNode, "REPORTONTITLE", "SETPAYMENTMETHOD"
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "SETPAYMENTMETHOD"
                xmlSetAttributeValue xmlAppNode, "COMPLETIONDATE", strFNDCompletionDate
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                
                If (lngErrorNumber = 0) Then
                    Set xmlPaymentMethod = xmlThisResponseDoc.selectSingleNode("//PAYMENTMETHODRETURNED")
                    strPaymentMethod = xmlGetAttributeText(xmlPaymentMethod, "PAYMENTMETHOD")
                Else
                    lngOmigaError = errGetOmigaErrorNumber(lngErrorNumber)
                    If (lngOmigaError = 8526) Then
                        strTaskNote = errGetMessageText(8526)
                        'strTaskNote = "Completion cannot occur today as it is now after the CHAPS cut-off time"
                        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                        errThrowError cstrFunctionName, lngOmigaError, strTaskNote
                    Else
                        strTaskNote = "Unable to create disbursement"
                        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                        errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                    End If
                End If
                
                'Set Advance Date
                xmlChangeNodeName xmlAppNode, "SETPAYMENTMETHOD", "SETADVANCEDATE"
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "SETADVANCEDATE"
                xmlSetAttributeValue xmlAppNode, "COMPLETIONDATE", strFNDCompletionDate
                xmlSetAttributeValue xmlAppNode, "PAYMENTMETHOD", strPaymentMethod
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                If (lngErrorNumber <> 0) Then
                    strTaskNote = "Unable to create disbursement"
                    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                Else
                    Set xmlAdvanceDate = xmlThisResponseDoc.selectSingleNode("//ADVANCEDATERETURNED")
                    strAdvanceDate = xmlGetAttributeText(xmlAdvanceDate, "ADVANCEDATE")
                End If
                
                'Get Payment Summary Record
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETPAYMENTSUMMARY"
                
                Set xmlTempNode = xmlTempResponse.createElement("PAYMENTRECORD")
                xmlNewRequest.appendChild xmlTempNode
                xmlSetAttributeValue xmlTempNode, "APPLICATIONNUMBER", strApplicationNumber
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                If (lngErrorNumber <> 0) Then
                    strTaskNote = "Unable to create disbursement"
                    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                Else
                    Set xmlPaymentSummary = xmlThisResponseDoc.selectSingleNode("//CALCULATIONS")
                    strBalance = xmlGetAttributeText(xmlPaymentSummary, "BALANCE")
                    strNetAmount = strBalance - xmlGetAttributeText(xmlPaymentSummary, "FEESTOBEDEDUCTEDFROMADVANCE")  ' MAR1137
                    
                    'MAR1646 GHun If there is a balance value and it is zero then fail
                    If Len(strBalance) > 0 And CSafeDbl(strBalance) = 0 Then
                        strTaskNote = errGetMessageText(8551)
                        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                        errThrowError cstrFunctionName, 8551
                    End If
                    'MAR1646 End
                End If
       
                'Get PayeeHistory Record
                Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
                
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "FINDPAYEEHISTORYLIST"
                
                Set xmlTempNode = xmlTempResponse.createElement("PAYEEHISTORY")
                xmlNewRequest.appendChild xmlTempNode
                xmlSetAttributeValue xmlTempNode, "APPLICATIONNUMBER", strApplicationNumber
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlNewRequest.xml))
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                If (lngErrorNumber <> 0) Then
                    strTaskNote = "Unable to create disbursement"
                    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                Else
                    'Get sequence number for Legal Rep
                    strLegalRep = objComboDO.GetFirstComboValueId("ThirdPartyType", "S")

                    Set xmlPayeeList = xmlThisResponseDoc.selectNodes("//PAYEEHISTORY")
    
                    If xmlPayeeList.length > 0 Then
                        For Each xmlTempNode In xmlPayeeList
                            Set xmlThirdParty = xmlTempNode.selectSingleNode("THIRDPARTY")
                            If (xmlGetAttributeText(xmlThirdParty, "THIRDPARTYTYPE") = strLegalRep) Then
                                strSequenceNo = xmlGetAttributeText(xmlTempNode, "PAYEEHISTORYSEQNO")
                                Exit For
                            End If
                        Next
                    End If
                 
                End If
  
                'MAR1137  Get Outstanding Fees
                Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlTempRequest, "OPERATION", "FINDFEEPAYMENTTOTALS"
                Set xmlTempNode = xmlTempResponse.createElement("APPLICATIONFEETYPE")
                xmlTempRequest.appendChild xmlTempNode
                xmlSetAttributeValue xmlTempNode, "APPLICATIONNUMBER", strApplicationNumber
                xmlSetAttributeValue xmlTempNode, "_COMBOLOOKUP_", "1"
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlTempRequest.xml))

                Set xmlFeeList = xmlThisResponseDoc.selectNodes(".//APPLICATIONFEETYPE")
  
                'CreateDisbursement
                Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                xmlSetAttributeValue xmlTempRequest, "OPERATION", "CREATEDISBURSEMENT"
                
                Set xmlTempNode = xmlTempResponse.createElement("PAYMENTRECORD")
                xmlTempRequest.appendChild xmlTempNode
                xmlSetAttributeValue xmlTempNode, "APPLICATIONNUMBER", strApplicationNumber
                xmlSetAttributeValue xmlTempNode, "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
                xmlSetAttributeValue xmlTempNode, "AMOUNT", strBalance
                xmlSetAttributeValue xmlTempNode, "PAYMENTMETHOD", objComboDO.GetFirstComboValueId("PaymentMethod", strPaymentMethod)
                xmlSetAttributeValue xmlTempNode, "CREATIONDATETIME", Format$(Now, "dd/mm/yy hh:mm:ss")

                Set xmlDisbursementNode = xmlTempResponse.createElement("DISBURSEMENTPAYMENT")
                xmlSetAttributeValue xmlDisbursementNode, "ISSUEDATE", Now()
                'JD MAR1220 use the advancedate    xmlSetAttributeValue xmlDisbursementNode, "COMPLETIONDATE", strFNDCompletionDate
                xmlSetAttributeValue xmlDisbursementNode, "COMPLETIONDATE", strAdvanceDate
                xmlSetAttributeValue xmlDisbursementNode, "PAYMENTTYPE", objComboDO.GetFirstComboValueId("PaymentType", "I")
                xmlSetAttributeValue xmlDisbursementNode, "PAYEEHISTORYSEQNO", strSequenceNo
                xmlSetAttributeValue xmlDisbursementNode, "PAYEETYPE", objComboDO.GetFirstComboValueId("PayeeType", "S")
                xmlSetAttributeValue xmlDisbursementNode, "PAYMENTSTATUS", objComboDO.GetFirstComboValueId("PaymentStatus", "P")  ' MAR1137
                xmlSetAttributeValue xmlDisbursementNode, "NETPAYMENTAMOUNT", strNetAmount
                
                xmlTempNode.appendChild xmlDisbursementNode
                
                'MAR1137
                'Add each Fee returned from GetPaymentSummary to the PaymentRecord node
                If (xmlFeeList.length > 0) Then
                    For Each xmlFee In xmlFeeList
                        xmlChangeNodeName xmlFee, "APPLICATIONFEETYPE", "FEEPAYMENT"
                        
                        strAmount = xmlGetAttributeText(xmlFee, "AMOUNTOUTSTANDING")
                        xmlFee.Attributes.removeNamedItem ("AMOUNT")
                        xmlFee.Attributes.removeNamedItem ("AMOUNTOUTSTANDING")

                        xmlSetAttributeValue xmlFee, "AMOUNTPAID", strAmount
                        xmlSetAttributeValue xmlFee, "PAYMENTEVENT", objComboDO.GetFirstComboValueId("PaymentEvent", "D")
                        xmlSetAttributeValue xmlFee, "COMPLETIONINDICATOR", "1"
                        
                        xmlTempNode.appendChild xmlFee
                    Next
                End If
                
                xmlTempRequest.appendChild xmlTempNode
                
                xmlThisResponseDoc.loadXML (objPayProc.omPayProcRequest(xmlTempRequest.xml))
                lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
                If (lngErrorNumber <> 0) Then
                    strTaskNote = "Unable to create disbursement"
                    CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
                End If

                'MAR1265 If we are not already at PreCompletion, move to that stage.
                '        Otherwise complete the task and remain at PreCompletion.

                strTMPreCompletionStage = GetGlobalParamString("TMPreCompletionStageID")

                If (strTMPreCompletionStage <> strCurrentStage) Then
                    'Move to Pre Completion Stage
                    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStage"
                    Dim xmlCaseStage  As IXMLDOMNode
                    Set xmlCaseStage = xmlTempResponse.createElement("CASESTAGE")
                    xmlTempRequest.appendChild xmlCaseStage

                    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "SOURCEAPPLICATION"
                    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "CASEID"
                    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYID"
                    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYINSTANCE"
                    xmlSetAttributeValue xmlCaseStage, "STAGEID", strTMPreCompletionStage

                    xmlTempRequest.appendChild xmlSaveAppNode.cloneNode(True)

                    Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
                    strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)

                    errCheckXMLResponse strResponse, True, vxmlResponseNode
                Else
                    'Update CaseTask as complete
                    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
                    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"

                    strUserId = xmlGetAttributeText(xmlTempRequest, "USERID")
                    strUnitId = xmlGetAttributeText(xmlTempRequest, "UNITID")
                    strUserAuthorityLevel = xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")

                    xmlSetAttributeValue xmlCaseTaskIn, "USERID", strUserId
                    xmlSetAttributeValue xmlCaseTaskIn, "UNITID", strUnitId
                    xmlSetAttributeValue xmlCaseTaskIn, "USERAUTHORITYLEVEL", strUserAuthorityLevel
                    xmlSetAttributeValue xmlCaseTaskIn, "TASKSTATUS", 40

                    xmlTempRequest.appendChild xmlCaseTaskIn.cloneNode(True)
    
                    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
                    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
                    errCheckXMLResponse strResponse, True, vxmlResponseNode
                
                End If
            Else
                'Generate Funds Release Failure task.
                strTaskNote = errGetMessageText(8545)
                'strTaskNote = "Validate Solicitor failed"
                CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                errThrowError cstrFunctionName, 8545, strTaskNote
            End If
        Else
            'MAR1110  Set up task note depending on error
            lngOmigaError = errGetOmigaErrorNumber(lngErrorNumber)
            
            If (lngOmigaError >= 8518) And (lngOmigaError <= 8521) Then
            
                If (lngOmigaError = 8518) Then
                    strTaskNote = errGetMessageText(8518)
                    'strTaskNote = "The Completion Date supplied is a day in the past"
                ElseIf (lngOmigaError = 8519) Then
                    strTaskNote = errGetMessageText(8519)
                    'strTaskNote = "The Completion Date supplied is a non-business day"
                ElseIf ((lngOmigaError = 8520) Or (lngOmigaError = 8521)) Then
                    strTaskNote = errGetMessageText(8520)
                    'strTaskNote = "Completion cannot occur on this date as the mortgage offer will have expired"
                End If
            
                CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                    
                errThrowError cstrFunctionName, lngOmigaError, strTaskNote
            Else
                strTaskNote = errGetMessageText(8546)
                'strTaskNote = "FND Completion Date is invalid"
                CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
                errThrowError cstrFunctionName, 8546, strTaskNote
            
            End If
        End If
    Else
        'Generate Funds Release Failure task.
        strTaskNote = errGetMessageText(8547)
        'strTaskNote = "MOF date is earlier than the latest ACA"
        CreateFundsReleaseFailedTask vxmlRequestNode, strTaskNote
        errThrowError cstrFunctionName, lngErrorNumber, strTaskNote

    End If
      
RunFirstTitleFNDExit:

    Set xmlFirstTitleRoot = Nothing
    Set AppFirstTitleNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskIn = Nothing
    Set xmlRequest = Nothing
    Set xmlAppNode = Nothing
    Set xmlAdvanceDate = Nothing
    Set xmlPaymentMethod = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempResponse = Nothing
    Set xmlROTNode = Nothing
    Set objNoTxBO = Nothing
    Set objPayProc = Nothing
    Set objROTBO = Nothing
    Set objMsgTmBO = Nothing
    Set xmlFeeList = Nothing
    Set xmlFee = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RunFirstTitleFNDErr:

    gobjContext.SetAbort
    
    GoTo RunFirstTitleFNDExit

End Sub

Private Sub RunFirstTitleTTL(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo RunFirstTitleTTLExit
    
    Const cstrFunctionName As String = "RunFirstTitleTTL"
    
    Dim strTitleNo1 As String
    Dim strTitleNo2 As String
    Dim strTitleNo3 As String
    Dim strUserId As String
    Dim strUnitId As String
    Dim strUserAuthorityLevel As String
    Dim strResponse As String
    Dim strTaskId As String
    Dim strTMFTAwaitingTTLTaskID   As String
    
    Dim lngErrorNumber As Long
    
    Dim objROTBO As Object
    Dim objMsgTmBO As Object
    
    Dim AppFirstTitleNode As IXMLDOMNode
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlFirstTitleRoot As IXMLDOMNode
    Dim xmlReportOnTitleNode As IXMLDOMNode
    Dim xmlRequest As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskNodeList As IXMLDOMNodeList
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlUpdateCastTaskReqNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    Set xmlRequest = vxmlRequestNode.selectSingleNode("//REQUEST")
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    
    'GetCurrentCaseStage
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GETCURRENTSTAGE"
    
    GetCurrentStage vxmlRequestNode, xmlDoc
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
    strTMFTAwaitingTTLTaskID = GetGlobalParamString("TMFTAwaitingTTLTaskID")
    
    Set xmlCaseTaskNodeList = xmlDoc.selectNodes("//CASETASK")
    
    If xmlCaseTaskNodeList.length > 0 Then
        
        For Each xmlNode In xmlCaseTaskNodeList
            
            strTaskId = xmlGetAttributeText(xmlNode, "TASKID")
            
            If strTaskId = strTMFTAwaitingTTLTaskID Then
                'Update CaseTask as complete
                Set xmlUpdateCastTaskReqNode = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlUpdateCastTaskReqNode, "OPERATION", "UPDATECASETASK"
        
                strUserId = xmlGetAttributeText(xmlUpdateCastTaskReqNode, "USERID")
                strUnitId = xmlGetAttributeText(xmlUpdateCastTaskReqNode, "UNITID")
                
                xmlSetAttributeValue xmlNode, "OWNINGUSERID", strUserId
                xmlSetAttributeValue xmlNode, "OWNINGUNITID", strUnitId
                xmlSetAttributeValue xmlNode, "TASKSTATUS", 40
    
                xmlUpdateCastTaskReqNode.appendChild xmlNode.cloneNode(True)
            
                Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
                strResponse = objMsgTmBO.TmRequest(xmlUpdateCastTaskReqNode.xml)
                errCheckXMLResponse strResponse, True, vxmlResponseNode
            End If
        
        Next
        
    End If
    
    'Get AppFirstTitle Record
    Set xmlAppNode = vxmlRequestNode.selectSingleNode("//APPLICATION")
    
    xmlSetAttributeValue xmlAppNode, "MESSAGETYPE", "TTL"
    
    GetApplicationFirstTitle vxmlRequestNode, xmlFirstTitleRoot
    errCheckXMLResponseNode xmlFirstTitleRoot, vxmlResponseNode, True
    
    Set AppFirstTitleNode = xmlFirstTitleRoot.selectSingleNode("//APPLICATIONFIRSTTITLE")
    
    strTitleNo1 = xmlGetAttributeText(AppFirstTitleNode, "TITLENUMBER1")
    strTitleNo2 = xmlGetAttributeText(AppFirstTitleNode, "TITLENUMBER2")
    strTitleNo3 = xmlGetAttributeText(AppFirstTitleNode, "TITLENUMBER3")
    
    xmlChangeNodeName vxmlRequestNode, "APPLICATION", "REPORTONTITLE"
    Set xmlReportOnTitleNode = vxmlRequestNode.selectSingleNode("//REPORTONTITLE")
            
    'MAR881 GHun
    'Check if there is an existing ReportOnTitle record
    xmlSetAttributeValue xmlRequest, "OPERATION", "GETREPORTONTITLEDATA"
    
    Set objROTBO = gobjContext.CreateInstance("omROT.omRotBO")
    xmlThisResponseDoc.loadXML objROTBO.OmRotRequest(vxmlRequestNode.xml)
    
    On Error Resume Next

    lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)

    On Error GoTo RunFirstTitleTTLExit
    
    If lngErrorNumber <> 0 Then

        'If the record not found then
        If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
            
            'Create a ROT Record
            xmlSetAttributeValue xmlRequest, "OPERATION", "CREATEREPORTONTITLE"
            
        ElseIf errGetOmigaErrorNumber(lngErrorNumber) <> oeRecordNotFound Then
            'Other than record not found raise the err
            errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
        End If
        
    Else
        'Update ROT Record
        Set xmlTempNode = xmlThisResponseDoc.selectSingleNode("//REPORTONTITLE")
        
        xmlCopyAttributeValue xmlTempNode, xmlReportOnTitleNode, "ROTGUID", "ROTGUID"
        xmlSetAttributeValue xmlRequest, "OPERATION", "UPDATEREPORTONTITLE"
    End If
        
    If (strTitleNo1 = "UNREG" Or strTitleNo2 = "UNREG" Or strTitleNo3 = "UNREG") Then
        xmlSetAttributeValue xmlReportOnTitleNode, "UNREGISTEREDRECORDED", "1"
    Else
        xmlSetAttributeValue xmlReportOnTitleNode, "TITLENUMBER", strTitleNo1
        xmlSetAttributeValue xmlReportOnTitleNode, "TITLENUMBER2", strTitleNo2
        xmlSetAttributeValue xmlReportOnTitleNode, "TITLENUMBER3", strTitleNo3
    End If
    
    'Perform the update/create
    xmlThisResponseDoc.loadXML objROTBO.OmRotRequest(vxmlRequestNode.xml)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True  'MAR881 GHun
    'MAR881 End
    
    'Update CaseTask as complete
    xmlSetAttributeValue xmlRequest, "OPERATION", "UPDATECASETASK"
        
    strUserId = xmlGetAttributeText(xmlRequest, "USERID")
    strUnitId = xmlGetAttributeText(xmlRequest, "UNITID")
    strUserAuthorityLevel = xmlGetAttributeText(xmlRequest, "USERAUTHORITYLEVEL")

    xmlSetAttributeValue xmlCaseTask, "OWNINGUSERID", strUserId
    xmlSetAttributeValue xmlCaseTask, "OWNINGUNITID", strUnitId
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlRequest.appendChild xmlCaseTask.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
RunFirstTitleTTLExit:
    
    Set objROTBO = Nothing
    Set objMsgTmBO = Nothing
    
    Set AppFirstTitleNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlFirstTitleRoot = Nothing
    Set xmlReportOnTitleNode = Nothing
    Set xmlRequest = Nothing
    Set xmlTempNode = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
        
End Sub

Private Sub RunFirstTitleMSFTOE(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
       
    Const cstrFunctionName As String = "RunFirstTitleMSFTOE"
    
    Dim objApp              As Object
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMNode
    Dim xmlAppReq           As IXMLDOMNode
    Dim xmlAppFFNode        As IXMLDOMNode
    Dim xmlElement          As IXMLDOMElement
    
On Error GoTo RunFirstTitleMSFTOEErr

    Set xmlDoc = New FreeThreadedDOMDocument40
    
    'MAR609 GHun Set FTTRANSFEROFEQUITYREQUIRED flag on ApplicationFactFind
    Set xmlRequest = vxmlRequestNode.cloneNode(False)
    Set xmlAppFFNode = xmlDoc.createElement("APPLICATIONFACTFIND")
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    Set xmlAppReq = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlElement.Text = xmlGetMandatoryAttributeText(xmlAppReq, "APPLICATIONNUMBER")
    xmlAppFFNode.appendChild xmlElement.cloneNode(True)
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = xmlGetMandatoryAttributeText(xmlAppReq, "APPLICATIONFACTFINDNUMBER")
    xmlAppFFNode.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("FTTRANSFEROFEQUITYREQUIRED")
    xmlElement.Text = "1"
    xmlAppFFNode.appendChild xmlElement
    xmlRequest.appendChild xmlAppFFNode
            
    Set objApp = gobjContext.CreateInstance("omApp.ApplicationBO")
    xmlDoc.loadXML objApp.UpdateApplicationFactFind(xmlRequest.xml)
        
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
RunFirstTitleMSFTOEErr:
    
    Set objApp = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlAppReq = Nothing
    Set xmlAppFFNode = Nothing
    Set xmlElement = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

'MAR1527  Add new interface RunFirstTitleACA
Private Sub RunFirstTitleACA(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo RunFirstTitleACAExit
    
    Const cstrFunctionName As String = "RunFirstTitleACA"
    
    Dim strUserId As String
    Dim strUnitId As String
    Dim strResponse As String
    Dim strTMFTAwaitingACATaskID As String
    Dim strValidationType As String
    Dim strPattern As String
    Dim strTaskStatus As String
    Dim bIncomplete As Boolean
    
    Dim objMsgTmBO As Object
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskNodeList As IXMLDOMNodeList
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlUpdateRequestNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim colValueIDs As Collection
    Dim varValueId As Variant
    
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    
    Set xmlUpdateRequestNode = xmlGetRequestNode(vxmlRequestNode)
        
    strUserId = xmlGetAttributeText(xmlUpdateRequestNode, "USERID")
    strUnitId = xmlGetAttributeText(xmlUpdateRequestNode, "UNITID")
    
    'GetCurrentCaseStage
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GETCURRENTSTAGE"
    
    GetCurrentStage vxmlRequestNode, xmlDoc
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
    strTMFTAwaitingACATaskID = GetGlobalParamString("TMFTAwaitingACATaskID")
    
    Set colValueIDs = New Collection
    GetValueIdsForValidationType "TaskStatus", "I", colValueIDs
    
    'Get all the incomplete tasks matching global parameter TMFTAwaitingACATaskID.
    strPattern = "//CASETASK[@TASKID='" & strTMFTAwaitingACATaskID & "']"
    Set xmlCaseTaskNodeList = xmlDoc.selectNodes(strPattern)
    
    If xmlCaseTaskNodeList.length > 0 Then
        
        'Update the tasks as complete
        For Each xmlNode In xmlCaseTaskNodeList
            
            'If the task has an Incomplete status, update the status to Complete.
            
            strTaskStatus = xmlGetAttributeText(xmlNode, "TASKSTATUS")
            bIncomplete = False
            For Each varValueId In colValueIDs
                If strTaskStatus = varValueId Then
                    bIncomplete = True
                End If
            Next
                        
            If (bIncomplete = True) Then
                Set xmlUpdateRequestNode = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlUpdateRequestNode, "OPERATION", "UPDATECASETASK"
        
                xmlSetAttributeValue xmlNode, "OWNINGUSERID", strUserId
                xmlSetAttributeValue xmlNode, "OWNINGUNITID", strUnitId
                xmlSetAttributeValue xmlNode, "TASKSTATUS", 40
    
                xmlUpdateRequestNode.appendChild xmlNode.cloneNode(True)
            
                Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
                strResponse = objMsgTmBO.TmRequest(xmlUpdateRequestNode.xml)
                errCheckXMLResponse strResponse, True, vxmlResponseNode
            
            End If
        
        Next
        
    End If
        
     'Update CaseTask as complete
    Set xmlUpdateRequestNode = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlUpdateRequestNode, "OPERATION", "UPDATECASETASK"
        
    xmlSetAttributeValue xmlCaseTask, "OWNINGUSERID", strUserId
    xmlSetAttributeValue xmlCaseTask, "OWNINGUNITID", strUnitId
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlUpdateRequestNode.appendChild xmlCaseTask.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlUpdateRequestNode.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
        
RunFirstTitleACAExit:
  
    Set objMsgTmBO = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseTaskNodeList = Nothing
    Set xmlCaseTask = Nothing
    Set xmlUpdateRequestNode = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
End Sub

Private Function MessageQueueComponentVC2_OnMessage(ByVal vstrXMLConfig As String, _
                                                    ByVal vstrXMLData As String) As Long
                                                    
On Error GoTo MessageQueueComponentVC2_OnMessageVbErr

    Const cstrFunctionName As String = "MessageQueueComponentVC2_OnMessage"
    
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMNode
   
    Dim strErrMsg As String
    Dim strResponse As String
    Dim strTASStatus As String
    
    If gobjContext Is Nothing Then
        Set gobjContext = GetObjectContext()
    End If

    strResponse = OmTmRequest(vstrXMLData)
    
    'Check the response
    Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
    Set xmlResponse = xmlTempResponse.selectSingleNode("RESPONSE")
    errCheckXMLResponseNode xmlResponse, , True
    strTASStatus = xmlGetAttributeText(xmlResponse, "TASSTATUS")
    
    ' Task Automation Processing
    If strTASStatus <> "FAILURE" Then
        'Complete so remove message from queue
        MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
        gobjContext.SetComplete
    Else
        strErrMsg = "Error processing message." & vbCrLf & _
        "Number: " & xmlGetNodeText(xmlResponse, "ERROR/NUMBER") & vbCrLf & _
        "Source: " & xmlGetNodeText(xmlResponse, "ERROR/SOURCE") & vbCrLf & _
        "Description: " & xmlGetNodeText(xmlResponse, "ERROR/DESCRIPTION")
        App.LogEvent strErrMsg, vbLogEventTypeError

        'Failure so move to dead queue
        MessageQueueComponentVC2_OnMessage = MESSQ_RESP_RETRY_MOVE_MESSAGE
        gobjContext.SetAbort
    End If
    
    
MessageQueueComponentVC2_OnMessage_Exit:
    
    Set xmlTempResponse = Nothing
    Set xmlResponse = Nothing
   
    Exit Function
    
MessageQueueComponentVC2_OnMessageVbErr:


    strErrMsg = "Error processing message." & vbCrLf & _
        "Number: " & CStr(Err.Number) & vbCrLf & _
        "Source: " & Err.Source & vbCrLf & _
        "Description: " & Err.Description
    App.LogEvent strErrMsg, vbLogEventTypeError
    
    MessageQueueComponentVC2_OnMessage = MESSQ_RESP_SUCCESS
    gobjContext.SetComplete
    
    GoTo MessageQueueComponentVC2_OnMessage_Exit

End Function

Private Function CreateLegalRepTask(ByVal vxmlRequestNode As IXMLDOMNode)
        
    On Error GoTo CreateLegalRepTaskErr

    Const cstrFunctionName As String = "omTMBO.CreateLegalRepTask"
        
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim objNoTxBO  As Object
    Dim strResponse  As String
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
    xmlTempRequest.setAttribute "OPERATION", "CREATELEGALREPTASK"
    xmlTempRequest.appendChild vxmlRequestNode.cloneNode(True)

    Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
    
    strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
    
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, , True
    
    gobjContext.SetComplete
     
CreateLegalRepTaskErrExit:
    
    Set xmlTempRequest = Nothing
    Set xmlTempDoc = Nothing
    Set objNoTxBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
    Exit Function
    
CreateLegalRepTaskErr:

    gobjContext.SetAbort
    
    GoTo CreateLegalRepTaskErrExit

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

Public Function OmTmRequest(ByVal vstrXmlIn As String) As String
On Error GoTo OmTmRequestVbErr

    Const cstrFunctionName As String = "OmTmRequest"

    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    Dim strOperation As String, _
        strAttribValue As String
    
    gobjContext.SetComplete 'AS 22/03/2006 MAR1501

    Set gobjTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    
    If gobjTmBO Is Nothing Then
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

    Debug.Print vstrXmlIn
    xmlInDoc.loadXML vstrXmlIn
    
    ' parse request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    xmlCheckMandatoryAttribute xmlRequestNode, "OPERATION"
    
    strOperation = xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    strOperation = UCase(strOperation)
    
     ' ik_debug
     ' xmlInDoc.Save "c:\omiga4Trace\omTm" & strOperation & "Request.xml"
    
    Select Case strOperation
            
        Case "CHASEUPTASK"
            ChaseUpTask xmlRequestNode
    
        Case "CREATEAPPLICATION"
            CreateApplication xmlRequestNode, xmlResponseNode
            
        Case "CREATECASESTAGETRIGGER"
            CreateCaseStageTrigger xmlRequestNode
            
        Case "CREATEADHOCCASETASK"
            CreateAdhocCaseTask xmlRequestNode
        
        Case "CRITICALDATACHECK"
            CriticalDataCheck xmlRequestNode, xmlResponseNode
    
        Case "FINDUNITTASKSUMMARY"
            FindUnitTaskSummary xmlRequestNode, xmlResponseNode
    
        'BMIDS697 GHun 22/01/2004 No longer used. The equivalent omTmNoTxBO methods
        'should be used instead
        'Case "MOVETONEXTSTAGE"
        '    MoveToNextStage xmlRequestNode
        '
        'Case "MOVETOSTAGE"
        '    MoveToStage xmlRequestNode
        'BMIDS697 End
            
        Case "REINSTATEPREVIOUSSTAGE"
            ReinstatePreviousStage xmlRequestNode
            
        Case "UPDATECASETASKPRIORITY"
            UpdateCaseTaskPriority xmlRequestNode
            
        Case "CREATECURREMPLOYERSREF"
            CreateCurrEmployersRef xmlRequestNode
            
        Case "CREATEPREVEMPLOYERSREF"
            CreatePrevEmployersRef xmlRequestNode
            
        Case "CREATEACCOUNTANTSREF"
            CreateAccountantsRef xmlRequestNode
            
        Case "CREATECURRLENDERSREF"
            CreateCurrLendersRef xmlRequestNode
            
        Case "CREATEPREVLENDERSREF"
            CreatePrevLendersRef xmlRequestNode
            
        Case "CREATECURRLANDLORDSREF"
            CreateCurrLandlordsRef xmlRequestNode
            
        Case "CREATEPREVLANDLORDSREF"
            CreatePrevLandlordsRef xmlRequestNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATECURREMPLOYERSREF"
            ValidateCurrEmployersRef xmlRequestNode, xmlResponseNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATEPREVEMPLOYERSREF"
            ValidatePrevEmployersRef xmlRequestNode, xmlResponseNode
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATEACCOUNTANTSREF"
            ValidateAccountantsRef xmlRequestNode, xmlResponseNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATECURRLENDERSREF"
            ValidateCurrLendersRef xmlRequestNode, xmlResponseNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATEPREVLENDERSREF"
            ValidatePrevLendersRef xmlRequestNode, xmlResponseNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATECURRLANDLORDSREF"
            ValidateCurrLandlordsRef xmlRequestNode, xmlResponseNode
            
        'SA SYS2256 Add in Response Node Argument
        Case "VALIDATEPREVLANDLORDSREF"
            ValidatePrevLandlordsRef xmlRequestNode, xmlResponseNode
            
        Case "PRINTDOCUMENTFORTASK"
            PrintDocumentForTask xmlRequestNode, xmlResponseNode    'MAR7 GHun
            
        ' DJP SYS1839
        Case "VALIDATEVALUATIONREPORT"
            ValidateValuationReport xmlRequestNode, xmlResponseNode
        
        ' DJP SYS1839
        Case "CREATEVALUATIONREPORT"
            CreateValuationReport xmlRequestNode
            
        'APWP3 - DPF 21/08/2002
        Case "CREATEVALUATIONREPORTNOINST"
            CreateValuationReportNoInst xmlRequestNode
        
        ' PSC SYS1971 - Start
        Case "IMPORTACCOUNTSINTOAPPLICATION"
            ImportAccountsIntoApplication xmlRequestNode, xmlResponseNode
            
        ' APS SYS1920
        Case "ISAPPLICATIONATOFFER"
            IsApplicationAtOfferStage xmlRequestNode, xmlResponseNode
            
        ' DRC BMIDS670
        Case "FREEZEUNFREEZEAPPLICATION"
            FreezeUnFreezeApplication xmlRequestNode, xmlResponseNode
                        
        'BMIDS697 GHun 04/03/2004 Moved to OmTmNoTxBO
        'Case "DECLINEAPPLICATION"
        '    DeclineApplication xmlRequestNode
        'BMIDS697 End
            
        Case "CREATEDEFAULTAPPLICATIONCONDITIONS"
            CreateDefaultApplicationConditions xmlRequestNode
        
        Case "UPDATEAPPLICATIONCONDITIONS"
            UpdateApplicationConditions xmlRequestNode
                        
        ' APS SYS1923
        Case "ISSUEOFFER"
            IssueOffer xmlRequestNode, xmlResponseNode
            
        ' JD BMIDS749
        Case "REMODELKFI"
            RemodelKFI xmlRequestNode, xmlResponseNode
            
        ' PSC SYS1879
        Case "APPROVERECOMMENDAPPLICATION"
            ApproveRecommendApplication xmlRequestNode, xmlResponseNode
        
        ' BMIDS682 a response is required
        Case "RUNCREDITCHECK"
            RunCreditCheck xmlRequestNode, xmlResponseNode
            
        Case "RUNRISKASSESSMENT"
            RunRiskAssessment xmlRequestNode
            
        Case "COMPLETEVALUERINSTRUCTIONS"
            CompleteValuerInstructions xmlRequestNode, xmlResponseNode

        ' JR SYS1878
        Case "CREATEREPORTONTITLE"
            CreateReportOnTitle xmlRequestNode
            
        Case "GETTASKCONTACTDETAILS"
            GetTaskContactDetails xmlRequestNode, xmlResponseNode
        
        ' JR SYS2048
        Case "VALIDATEREPORTONTITLE"
            ValidateReportOnTitle xmlRequestNode, xmlResponseNode
        
        ' SR SYS2325
        Case "GETNEWNUMBERS"
            GetNewNumbers xmlRequestNode, xmlResponseNode
            
        ' DC SYS2266
        Case "PROCESSTASK"
             ProcessTask xmlRequestNode
                     
        Case "RECEIVEDOFFER"
             ReceivedOffer xmlRequestNode, xmlResponseNode
             
        'DM     17/01/02    SYS3185 Made GetCurrentStage accessible through omTMRequest
        Case "GETCURRENTSTAGE"
            GetCurrentStage xmlRequestNode, xmlOutDoc
                     
        ' BMIDS00005 GHun
        Case "UPDATECRSCONTACTLOG"
            UpdateCRSContactLog xmlRequestNode, xmlResponseNode
            
        ' BMIDS00026 GHun
        Case "UPDATECRSCUSTOMER"
            UpdateCRSCustomer xmlRequestNode, xmlResponseNode
                     
        'BMIDS00025 MDC 10/06/2002
        Case "RUNHUNTERINTERFACE"
            RunHunterInterface xmlRequestNode
        'BMIDS00025 MDC 10/06/2002 - End
        
        'BMIDS00006 MDC 17/07/2002
        Case "IMPORTCUSTOMERSINTOAPPLICATION"
            ImportCustomersIntoApplication xmlRequestNode, xmlResponseNode
        'BMIDS00006 MDC 17/07/2002 - eND
        
        'MO 01/08/2002 BMIDS00218 IDWP1 - Start
        Case "CREATEACTIVITY"
            CreateActivity xmlRequestNode, xmlResponseNode
        'MO 01/08/2002 BMIDS00218 IDWP1 - End
                     
        'BMIDS00336 MDC 23/08/2002
        Case "IMPORTBUREAUDATA"
            ImportBureauData xmlRequestNode
        'BMIDS00336 MDC 23/08/2002 - End
                     
        ' PSC 21/10/2002 BMIDS00676 - Start
        Case "RUNGENERICINTERFACE"
            RunGenericInterface xmlRequestNode, xmlResponseNode
        ' PSC 21/10/2002 BMIDS00676 - End
        
        'MO - 15/11/2002 - BMIDS00814 - START
        Case "TRANSFERAPPLICATIONOWNERSHIP"
            TransferApplicationOwnership xmlRequestNode, xmlResponseNode
        'MO - 15/11/2002 - BMIDS00814 - END
        
        'PSC 05/12/2002 BM0167 - Start
        Case "UNAPPROVEAPPLICATION"
            UnapproveApplication xmlRequestNode, xmlResponseNode
        'PSC 05/12/2002 BM0167 - End
        
        'BMIDS748
        Case "RUNENQUIRYCREDITCHECK"
            RunEnquiryCreditCheck xmlRequestNode, xmlResponseNode
            
        Case "RUNUPGRADETOFULLAPPCREDITCHECK"
            RunUpgradeToFullAppCreditCheck xmlRequestNode, xmlResponseNode
               
        ' PSC 24/08/2005 MAR32 - Start
        Case "FINDCASETASKLISTLITE"
            FindCaseTaskListLite xmlRequestNode, xmlResponseNode
        Case "PROCESSTASTASK"
            ProcessTASTask xmlRequestNode, xmlResponseNode
        Case "TASMOVETONEXTSTAGE"
            TASMoveToNextStage xmlRequestNode, xmlResponseNode
        Case "RUNAUTOAPPLICATIONEXPIRY"
            RunAutoApplicationExpiry xmlRequestNode, xmlResponseNode
        ' PSC 24/08/2005 MAR32 - End
        
        'MAR23 Add RunKnowYourCustomer
        Case "RUNKNOWYOURCUSTOMER"
            KnowYourCustomer xmlRequestNode, xmlResponseNode
            
        'MAR582  Add RUNFTREISSUEOFFERINTERFACE
        'MAR609 GHun Made RUNFTCOMPDATEUPDATEINTERFACE less than 30 characters
        Case "RUNFTINITIALREQUESTINTERFACE", "RUNFTOFFERREQUESTINTERFACE", "RUNFTCANCELAPPINTERFACE", _
             "RUNFTFUNDSRELEASEFAILINTERFACE", "RUNFTMISCUPDATEINTERFACE", "RUNFTCOMPDATEUPDATEINTERFACE", _
             "RUNFTOFFEREXPIRYINTERFACE", "RUNFTREINSTATEAPPINTERFACE", "RUNFTREISSUEOFFERINTERFACE"
            RunFTInterfaceCall xmlRequestNode, xmlResponseNode
        Case "RUNFIRSTTITLETTL"
            RunFirstTitleTTL xmlRequestNode, xmlResponseNode
        Case "RUNFIRSTTITLEFND"
            RunFirstTitleFND xmlRequestNode, xmlResponseNode
        Case "RUNFIRSTTITLEMSFTOE"
            RunFirstTitleMSFTOE xmlRequestNode, xmlResponseNode
        Case "RUNFIRSTTITLEACA"                                   ' MAR1527
            RunFirstTitleACA xmlRequestNode, xmlResponseNode      ' MAR1527
        Case "HANDLEINTERFACERESPONSE"  'SR 08/10/2005 :TEMP**
            HandleInterfaceResponse xmlRequestNode, xmlResponseNode
        Case "RUNHOMETRACKVALUATION"
            RunHomeTrackValuation xmlRequestNode, xmlResponseNode
        Case "RUNESURVVALUATION"  'JD MAR40
            RunEsurvValuation xmlRequestNode, xmlResponseNode
        'MAR184 GHun
        Case "RUNXMLCREDITCHECK"
            RunXMLCreditCheck xmlRequestNode, xmlResponseNode
        Case "RUNXMLREPROCESSCREDITCHECK"
            RunXMLCreditCheck xmlRequestNode, xmlResponseNode, False, True, False
        Case "RUNXMLRESCORECREDITCHECK"
            RunXMLCreditCheck xmlRequestNode, xmlResponseNode, True, False, False
        'MAR184 End
        'MAR1617 add LIM Update flag
        Case "RUNXMLEXPUPDATELIM"
            RunXMLCreditCheck xmlRequestNode, xmlResponseNode, True, False, True
        'MAR268 SR
        Case "RUNHUNTEREXPERIANINTERFACE"
            RunHunterExperianInterface xmlRequestNode, xmlResponseNode
        'MAR316
        Case "RUNDELAYEDCOMPLETIONINTERFACE"
            RunDelayedCompletionInterface xmlRequestNode, xmlResponseNode
            
        ' TW 03/11/2005 MAR211
        'MAR563 Interface name too long - renamed
        Case "SMSFULFILMENTEXTRACT"
            SMSFulfilmentExtract xmlRequestNode, xmlResponseNode
        
        'MAR563 Interface name too long - renamed
        Case "EMAILFULFILMENTEXTRACT"
            EMailFulfilmentExtract xmlRequestNode, xmlResponseNode
        ' TW 03/11/2005 MAR211 End
        'INR   MAR731
        Case "RUNPINNUMBERTRIGGERINTERFACE"
            RunPinNumberTriggerInterface xmlRequestNode, xmlResponseNode
        
        'INR   MAR731
        Case "RUNLEGALADDRESSTRIGGERINTERFAC"
            RunLegalAddressTriggerInterface xmlRequestNode, xmlResponseNode
            
        'MAR1088 GHun
        Case "RUNAUTOAPPLICATIONDECLINE"
            RunAutoApplicationDecline xmlRequestNode, xmlResponseNode
        'MAR1088 End
        
        'MAR1434 JD
        Case "GETMAXCASETASKINSTANCE"
            RunGetMaxCaseTaskInstance xmlRequestNode, xmlResponseNode
        
        'MAR1300 GHun
        Case "SETCHANGEOFPROPERTY"
            SetChangeOfProperty xmlRequestNode, xmlResponseNode
        ' PSC 05/05/2006 MAR1593 - Start
        Case "RUNEMAILSMSINTERFACE"
            RunEmailSMSInterface xmlRequestNode, xmlResponseNode
        ' PSC 05/05/2006 MAR1593 - End
                
#If EPSOM Then

        'EP2 IK
        Case "RUNEPSOMCREDITCHECK"
            RunEpsomCreditCheck xmlRequestNode, xmlResponseNode
        
        Case "RUNEPSOMCASEASSESS"
            RunEpsomCaseAssess xmlRequestNode, xmlResponseNode
        
        Case "RUNEPSOMKYC"
            RunEpsomKYC xmlRequestNode, xmlResponseNode
        
        'LDM 12/03/2006 EP6
        Case "RUNXMLEXPQUOTATIONCONVERSIONCC"
            RunXMLExpQuotationConversionCC xmlRequestNode, xmlResponseNode
        'EP6 End
            
        'EP17 SAB
        Case "RUNREALHUNTEREXPERIANINTERFACE"
            RunRealHunterExperianInterface xmlRequestNode, xmlResponseNode
        'EP17 SAB - End
        
        'LDM 11/04/06 EP377
        Case "RUNEPSOMREPROCESSCREDITCHECK"
            RunEpsomCreditCheck xmlRequestNode, xmlResponseNode, False, True
        
        Case "RUNEPSOMRESCORECREDITCHECK"
            RunEpsomCreditCheck xmlRequestNode, xmlResponseNode, True, False
        'EP377 End
        
        ' IK_18/05/2006_EP574
        Case "BUILDHANDOVERDOC"
            BuildHandoverDoc xmlRequestNode, xmlResponseNode
        ' IK_18/05/2006_EP574_ends
        
        'EP2 IK ends

        ' PE 09/11/2006 - EP2_67
        Case "RUNVEXVALUATIONINSTRUCTION"
            RunVexValuationInstruction xmlInDoc, xmlOutDoc

        ' PE 09/11/2006 - EP2_67
        Case "RUNVEXSTATUSUPDATE"
            RunVexStatusUpdate xmlInDoc, xmlOutDoc
            
        'EP2_1903 GHun
        Case "HANDLEINBOUNDDOCUMENT"
            HandleInboundDocument xmlRequestNode
        'EP2_1903 End
#End If

        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, "OPERATION = " & strOperation
        ' PSC SYS1971 - End
    End Select
        
    If xmlGetAttributeText(xmlResponseElem, "TYPE") <> "WARNING" Then ' SR 09/09/2004 : BMIDS815
        xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    End If ' SR 09/09/2004 : BMIDS815
        
    OmTmRequest = xmlOutDoc.xml
    Debug.Print OmTmRequest
    
OmTmRequestExit:
    
    Set gobjTmBO = Nothing
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    
    Exit Function

OmTmRequestVbErr:
    
    OmTmRequest = errCreateErrorResponse()
    Debug.Print OmTmRequest
    
'IK_31/05/2006_EP626
#If EPSOM Then
    GetObjectContext.SetAbort
#End If
'IK_31/05/2006_EP626_ends
    
    GoTo OmTmRequestExit
    
End Function

'SR 08/10/2005  New procedure
Private Sub HandleInterfaceResponse(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)
                                    
On Error GoTo HandleInterfaceResponseExit

    Const cstrFunctionName As String = "HandleInterfaceResponse"
        
    'MAR630  Changes throughout procedure for task creation
        
    Dim xmlInterface As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlCaseActivity As IXMLDOMNode
    Dim xmlCaseStage As IXMLDOMNode
    Dim xmlGetCurrentStage As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlInterfaceTaskResults As IXMLDOMNode
    Dim xmlTemp As IXMLDOMNode
    Dim xmlTemp2 As IXMLDOMNode
        
    Dim strInterfaceType As String
    Dim strCreateTaskFlag As String
    Dim strExperianID As String
    Dim strCaseAssessmentID As String
    Dim strFilter As String
    Dim strResponse As String
    Dim strSQL As String
    Dim strAppNo As String
    Dim strAFFNo As String
    Dim strTaskId As String
    Dim strTaskStatus As String
    Dim strCaseActivityGuid As String
    Dim iTaskInstance As Integer
    Dim strCondition As String
    Dim blnGenerateProcessTask As Boolean
    Dim blnUWTask As Boolean
    Dim bComplete As Boolean
    Dim dblSeniorUWThreshold As Double
    Dim dblAmountRequested As Double
    Dim strFirstTitleID  As String
    Dim strMessageType As String
    Dim objAQ As Object
    Dim colValueIDs As Collection
    Dim varValueId As Variant

    Dim strReferProcTask As String
    'Dim strReferRiskTask As String
    Dim strReferUWTask As String
    Dim strReferUW2Task As String
    
    Dim strContext As String    'MAR1715 GHun
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlInterface = vxmlRequestNode.selectSingleNode(".//INTERFACE")
   
    strInterfaceType = xmlGetAttributeText(xmlInterface, "INTERFACETYPE")
    strMessageType = xmlGetAttributeText(xmlInterface, "MESSAGETYPE")
    strCreateTaskFlag = xmlGetAttributeText(xmlInterface, "CREATETASKFLAG")
    strContext = xmlGetAttributeText(xmlInterface, "CONTEXT")   'MAR1715 GHun
    
    'JD MAR342 set strAppNo here
    Set xmlTemp = vxmlRequestNode.selectSingleNode(".//APPLICATION")
    strAppNo = xmlGetAttributeText(xmlTemp, "APPLICATIONNUMBER")
    strAFFNo = xmlGetAttributeText(xmlTemp, "APPLICATIONFACTFINDNUMBER")
    
    'If any MessageSubTypes have been passed in, get all the corresponding tasks
    Set xmlInterfaceTaskResults = xmlDoc.createElement("INTERFACETASKLIST")     ' MAR712
    Set xmlNodeList = xmlInterface.selectNodes("MESSAGESUBTYPELIST/MESSAGESUBTYPE")
             
    If xmlNodeList.length > 0 Then
        'MAR609 GHun
        'EP2_682 PEdney
        strSQL = "SELECT DISTINCT(TASKID) FROM TASKINTERFACESUBMESSAGE WHERE" & _
                " INTERFACETYPE = " & strInterfaceType & _
                " AND MESSAGETYPE = " & strMessageType & _
                " AND TASKID NOT IN (SELECT TASKID FROM TASK WHERE TASK.DELETEFLAG=1)"
    
        strFilter = vbNullString
        For Each xmlNode In xmlNodeList
            If (Len(strFilter) > 0) Then
                strFilter = strFilter & ", "
            End If
            strFilter = strFilter & xmlGetAttributeText(xmlNode, "MESSAGESUBTYPE")
        Next
        'MAR609 End
        strFilter = "MESSAGESUBTYPE IN ( " & strFilter & " ) "
        strSQL = strSQL & " AND " & strFilter
        GetRecordsAsXMLForSQL strSQL, "TASKINTERFACESUBMESSAGE_TASK", xmlInterfaceTaskResults
        
        xmlChangeNodeName xmlInterfaceTaskResults, "TASKINTERFACESUBMESSAGE_TASK", "TASK_ID"
        
    Else 'get all the tasks with the MessageTypes passed in
        'MAR609 GHun
        'EP2_682 PEdney
        strSQL = "SELECT DISTINCT(TASKID) FROM TASKINTERFACEMESSAGE WHERE" & _
                " INTERFACETYPE = " & strInterfaceType & _
                " AND MESSAGETYPE = " & strMessageType & _
                " AND TASKID NOT IN (SELECT TASKID FROM TASK WHERE TASK.DELETEFLAG=1)"
        'MAR609 End
    
        GetRecordsAsXMLForSQL strSQL, "TASKINTERFACEMESSAGE_TASK", xmlInterfaceTaskResults
        
        xmlChangeNodeName xmlInterfaceTaskResults, "TASKINTERFACEMESSAGE_TASK", "TASK_ID"

    End If
    vxmlResponseNode.appendChild xmlInterfaceTaskResults
    
    'MAR712 Start
    'If no tasks have been found, no more processing is necessary
    If (xmlInterfaceTaskResults.selectNodes("//TASK_ID").length > 0) Then
        
    ' For Experian & Case Assessment, determine the tasks to be created
    strExperianID = GetFirstComboValueId("InterfaceType", "EXP")
    strCaseAssessmentID = GetFirstComboValueId("InterfaceType", "CA")   ' MAR630
    strFirstTitleID = GetFirstComboValueId("InterfaceType", "FT")
    
    blnUWTask = False 'MAR1567 initialise as false
    blnGenerateProcessTask = False
        
    If strInterfaceType = strExperianID Or strInterfaceType = strCaseAssessmentID Then
        
        'Get the Task IDs from the Global Parameters
        strReferProcTask = GetMandatoryGlobalParamString("TMReferProcTaskID")
        'strReferRiskTask = GetMandatoryGlobalParamString("TMReferRiskTaskID")
        strReferUWTask = GetMandatoryGlobalParamString("TMReferUWTaskID")
        strReferUW2Task = GetMandatoryGlobalParamString("TMReferUW2TaskID")
       
        Set objAQ = gobjContext.CreateInstance("omAQ.ApplicationQuoteBO")
        dblSeniorUWThreshold = GetMandatoryGlobalParamAmount("SeniorUnderWriterThreshold")
        
        'Loop over each task found.
        'Keep the correct Underwriter Referral task depending on Amount Requested.
        'Do not keep Processor Referral tasks if an Underwriter Referral task is present.
        For Each xmlNode In xmlInterfaceTaskResults.selectNodes("//TASK_ID")
        
            strTaskId = xmlGetAttributeText(xmlNode, "TASKID")
        
            'Check for Underwriter referral tasks
            If ((strTaskId = strReferUWTask) Or (strTaskId = strReferUW2Task)) Then
                
                'Get Amount Requested
                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                Set xmlTemp = xmlDoc.createElement("APPLICATION")
                xmlNewRequest.appendChild xmlTemp
                    
                Set xmlTemp2 = xmlDoc.createElement("APPLICATIONNUMBER")
                xmlTemp2.Text = strAppNo
                xmlTemp.appendChild xmlTemp2
                    
                Set xmlTemp2 = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
                xmlTemp2.Text = strAFFNo
                xmlTemp.appendChild xmlTemp2
                    
                strResponse = objAQ.GetAcceptedOrActiveQuoteData(xmlNewRequest.xml)
                Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
                Set xmlTemp = xmlDoc.selectSingleNode(".//MORTGAGESUBQUOTE")
                
#If EPSOM Then
                ' ik_26/04/2006_EP465
                If Not xmlTemp Is Nothing Then
                    dblAmountRequested = xmlGetNodeAsDouble(xmlTemp, "//AMOUNTREQUESTED")
                Else
                    Set xmlDoc = New FreeThreadedDOMDocument40
                    xmlDoc.async = False
                    xmlDoc.loadXML GetNewLoan(vxmlRequestNode)
                    If Not xmlDoc.selectSingleNode("RESPONSE/NEWLOAN[@AMOUNTREQUESTED]") Is Nothing Then
                        If IsNumeric(xmlDoc.selectSingleNode("RESPONSE/NEWLOAN/@AMOUNTREQUESTED").Text) Then
                            dblAmountRequested = CDbl(xmlDoc.selectSingleNode("RESPONSE/NEWLOAN/@AMOUNTREQUESTED").Text)
                        End If
                    End If
                End If
                ' ik_26/04/2006_EP465_ends
#Else
                dblAmountRequested = xmlGetNodeAsDouble(xmlTemp, "//AMOUNTREQUESTED")
#End If

                'Check Amount Requested against threshold
                If dblAmountRequested > dblSeniorUWThreshold Then
                    If (strTaskId <> strReferUW2Task) Then
                        'Remove this task from the list
                        xmlInterfaceTaskResults.removeChild xmlNode
                    Else
                        blnUWTask = True
                        blnGenerateProcessTask = False
                    End If
                Else
                    If (strTaskId <> strReferUWTask) Then
                        'Remove this task from the list
                        xmlInterfaceTaskResults.removeChild xmlNode
                    Else
                        blnUWTask = True
                        blnGenerateProcessTask = False
                    End If
                End If
            'Check for Processor referral tasks
            ElseIf (strTaskId = strReferProcTask) Then
                If (blnUWTask = False) Then
                    blnGenerateProcessTask = True
                End If
            End If
                 
        Next xmlNode
        
        'If we do not want Processor Referral tasks, remove them from the list.
        If blnGenerateProcessTask = False Then
             
            For Each xmlNode In xmlInterfaceTaskResults.selectNodes("//TASK_ID")
                If (xmlGetAttributeText(xmlNode, "TASKID") = strReferProcTask) Then
                    xmlInterfaceTaskResults.removeChild xmlNode
                End If
             Next
        End If
    End If
        
    If strCreateTaskFlag = "1" Then
        'Get all the tasks at the current stage
        If Not vxmlRequestNode.selectSingleNode(".//CASEACTIVITY") Is Nothing Then
            Set xmlCaseActivity = vxmlRequestNode.selectSingleNode(".//CASEACTIVITY").cloneNode(True)
        End If
               
        Set xmlGetCurrentStage = xmlDoc.createElement("CASEACTIVITY")
    
        xmlSetAttributeValue xmlGetCurrentStage, "SOURCEAPPLICATION", "Omiga"
        xmlSetAttributeValue xmlGetCurrentStage, "CASEID", strAppNo
        If Not xmlCaseActivity Is Nothing Then
            xmlSetAttributeValue xmlGetCurrentStage, "ACTIVITYID", xmlGetAttributeText(xmlCaseActivity, "ACTIVITYID")
        Else
            xmlSetAttributeValue xmlGetCurrentStage, "ACTIVITYID", GetGlobalParamString("TMOmigaActivity")
        End If
            
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlNewRequest.appendChild xmlGetCurrentStage
        GetCurrentStage xmlNewRequest, xmlDoc
        errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
        
        'Create tasks
        
        Dim strCaseStageSeqNo As String
        Dim strStageId As String
        Dim strTaskNotes As String
        Dim strMemoNoteID As String
        Dim blnCreateTask As Boolean
        
        Set xmlCaseStage = xmlDoc.selectSingleNode(".//CASESTAGE")
        strCaseStageSeqNo = xmlGetAttributeText(xmlCaseStage, "CASESTAGESEQUENCENO")
        strStageId = xmlGetAttributeText(xmlCaseStage, "STAGEID")
        strCaseActivityGuid = xmlGetAttributeText(xmlCaseStage, "CASEACTIVITYGUID")
        strTaskNotes = xmlGetAttributeText(xmlInterface, "TASKNOTES")
        strMemoNoteID = GetFirstComboValueId("TaskNoteType", "M")
        
        For Each xmlNode In xmlInterfaceTaskResults.selectNodes("//TASK_ID")
            'If a task does not already exist or exists but is completed create the task now.
            strTaskId = xmlGetAttributeText(xmlNode, "TASKID")
            strCondition = ".//CASETASK[@TASKID='" & strTaskId & "']"
            
            If strInterfaceType <> strFirstTitleID Then
                
                blnCreateTask = True
                
                Set xmlNodeList = xmlDoc.selectNodes(strCondition)
                If xmlNodeList.length = 0 Then
                    blnCreateTask = True
                Else
                    'The task already exists - check for it being complete
                    
                    Set colValueIDs = New Collection
                    GetValueIdsForValidationType "TaskStatus", "C", colValueIDs
                                        
                    For Each xmlTempNode In xmlNodeList
                        strTaskStatus = xmlGetAttributeText(xmlTempNode, "TASKSTATUS")
                        bComplete = False
                        For Each varValueId In colValueIDs
                            If strTaskStatus = varValueId Then
                                bComplete = True
                            End If
                        Next
                        
                        If (bComplete = False) Then
                            blnCreateTask = False
                        End If
                    Next xmlTempNode
                                            
                End If
            Else
                'Always create the task for First Title
                blnCreateTask = True
            End If
            
            
            If blnCreateTask Then ' Create Task here
                If xmlNewRequest.selectSingleNode("CASETASK") Is Nothing Then   'MAR972 GHun
                    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                    
                    'IK 25/10/2005 MAR232
                    Set xmlTemp = xmlNewRequest.appendChild(vxmlRequestNode.selectSingleNode("APPLICATION").cloneNode(False))
                    AddApplicationPriority xmlTemp
                    'IK 25/10/2005 MAR232
                End If  'MAR972 GHun
                
                Set xmlTemp = xmlDoc.createElement("CASETASK")
                xmlSetAttributeValue xmlTemp, "SOURCEAPPLICATION", "Omiga"
                xmlSetAttributeValue xmlTemp, "CASEID", strAppNo
                'DRC MAR512 - check for CaseActivity otherwise use default omiga
                If Not xmlCaseActivity Is Nothing Then
                    xmlSetAttributeValue xmlTemp, "ACTIVITYID", xmlGetAttributeText(xmlCaseActivity, "ACTIVITYID")
                Else
                    xmlSetAttributeValue xmlTemp, "ACTIVITYID", GetGlobalParamString("TMOmigaActivity")
                End If
                xmlSetAttributeValue xmlTemp, "ACTIVITYINSTANCE", "1"
                xmlSetAttributeValue xmlTemp, "CASESTAGESEQUENCENO", strCaseStageSeqNo
                xmlSetAttributeValue xmlTemp, "STAGEID", strStageId
                xmlSetAttributeValue xmlTemp, "TASKID", strTaskId
                xmlSetAttributeValue xmlTemp, "CASETASKNAME", xmlGetAttributeText(xmlNode, "TASKNAME")
                xmlSetAttributeValue xmlTemp, "TASKSTATUS", GetTaskStatusFromInterfaceMessage(strTaskId, strInterfaceType, strMessageType)  'MAR847 GHun
                               
                'MAR1715 GHun
                If Len(strContext) > 0 Then
                    xmlSetAttributeValue xmlTemp, "CONTEXT", strContext
                End If
                'MAR1715 End
                               
                'MAR972 GHun
                'If TaskNote was passed in, create TaskNote record
                If Len(strTaskNotes) > 0 Then
                    
                    Set xmlTemp2 = xmlDoc.createElement("TASKNOTE")
                    xmlSetAttributeValue xmlTemp2, "NOTEENTRY", strTaskNotes
                    xmlSetAttributeValue xmlTemp2, "NOTETYPE", strMemoNoteID
                    xmlTemp.appendChild xmlTemp2
                    
                End If
                
                xmlNewRequest.appendChild xmlTemp
                 
                'CreateAdhocCaseTask xmlNewRequest
                '
                ''If TaskNote was passed in, create TaskNote record
                'If Len(strTaskNotes) > 0 Then
                '
                '    iTaskInstance = GetMaxCaseTaskInstance(strStageId, strTaskId, strCaseActivityGuid, strCaseStageSeqNo)
                '
                '    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                '    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CreateTaskNote"
                '
                '    Set xmlTemp = xmlDoc.createElement("TASKNOTE")
                '    xmlSetAttributeValue xmlTemp, "SOURCEAPPLICATION", "Omiga"
                '    xmlSetAttributeValue xmlTemp, "CASEACTIVITYGUID", strCaseActivityGuid
                '    xmlSetAttributeValue xmlTemp, "STAGEID", strStageId
                '    xmlSetAttributeValue xmlTemp, "CASESTAGESEQUENCENO", strCaseStageSeqNo
                '    xmlSetAttributeValue xmlTemp, "TASKID", xmlGetAttributeText(xmlNode, "TASKID")
                '    xmlSetAttributeValue xmlTemp, "TASKINSTANCE", CStr(iTaskInstance)
                '    xmlSetAttributeValue xmlTemp, "NOTEENTRY", strTaskNotes
                '    xmlSetAttributeValue xmlTemp, "NOTETYPE", strMemoNoteID
                '    xmlNewRequest.appendChild xmlTemp
                '    xmlDoc.loadXML gobjTmBO.TmRequest(xmlNewRequest.xml)
                '
                '    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
                'End If
                'MAR972 End
            End If
        Next
        
        'MAR972 GHun
        If Not xmlNewRequest.selectSingleNode("CASETASK") Is Nothing Then
            CreateAdhocCaseTask xmlNewRequest
        End If
        'MAR972 End
    End If
    
    End If  'MAR712 End
    
HandleInterfaceResponseExit:
    Set xmlInterface = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlCaseStage = Nothing
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlInterfaceTaskResults = Nothing
    Set xmlGetCurrentStage = Nothing
    Set xmlTemp = Nothing
    Set xmlTemp2 = Nothing
    Set objAQ = Nothing
  
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'IK 25/10/2005 MAR232
Private Sub AddApplicationPriority(ByVal vxmlApplicationNode As IXMLDOMElement)
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    Dim strSQL As String
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    
    Set xmlElem = xmlDoc.createElement("RESPONSE")
    Set xmlNode = xmlDoc.appendChild(xmlElem)
        
    strSQL = "SELECT APPLICATIONNUMBER,APPPRIORITYSEQUENCENO,APPLICATIONPRIORITYVALUE FROM APPLICATIONPRIORITY WHERE APPLICATIONNUMBER='" & vxmlApplicationNode.getAttribute("APPLICATIONNUMBER") & "' order by APPPRIORITYSEQUENCENO desc"
    GetRecordsAsXMLForSQL strSQL, "APPLICATIONPRIORITY", xmlNode
    
    Set xmlElem = xmlNode.selectSingleNode("APPLICATIONPRIORITY[@APPLICATIONPRIORITYVALUE]")
    
    If xmlElem Is Nothing Then
        errThrowError "AddApplicationPriority", oeMissingElementValue, "no APPLICATIONPRIORITY for APPLICATIONNUMBER " & vxmlApplicationNode.getAttribute("APPLICATIONNUMBER")
    End If
    
    vxmlApplicationNode.setAttribute "APPLICATIONPRIORITY", xmlElem.getAttribute("APPLICATIONPRIORITYVALUE")

End Sub
'IK 25/10/2005 MAR232 ends

Private Sub GetRecordsAsXMLForSQL(ByVal strSQL As String, _
                                ByVal strSchemaName As String, _
                                ByVal xmlResponse As IXMLDOMNode)
    
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim xmlSchema As IXMLDOMNode
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    Set xmlSchema = adoGetSchema(strSchemaName)
     
    'MAR712  Check for records found
    If Not rst Is Nothing Then
        adoConvertRecordSetToXML xmlSchema, xmlResponse, rst, False
    End If
    
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlSchema = Nothing
End Sub

Private Function GetTaskStatusFromInterfaceMessage(ByVal vstrTaskId As String, ByVal vstrInterfaceType As String, ByVal vstrMessageType As String) As String
    
    Const cstrFunction  As String = "GetTaskStatusFromInterfaceMessage"
    Dim cmd             As ADODB.Command
    Dim rst             As ADODB.Recordset
    Dim strSQL          As String
    Dim strStatus       As String

On Error GoTo GetTaskStatusFromInterfaceMessageVbErr

    strStatus = "10"    'MAR847 GHun Default task status to incomplete
    
    'MAR847 GHun Include InterfaceType and MessageType in query
    strSQL = "SELECT CREATEASCOMPLETE FROM TASKINTERFACEMESSAGE WHERE TASKID = '" & vstrTaskId & _
        "' AND INTERFACETYPE='" & vstrInterfaceType & "' AND MESSAGETYPE='" & vstrMessageType & "'"
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    
    'MAR847 GHun Check of the task should be created as complete
    If Not rst Is Nothing Then
        'IK 25/10/2005 MAR232
        If Not (IsNull(rst.Fields(0).Value)) Then
            If CSafeBool(rst.Fields(0).Value) Then
                strStatus = "40"
            End If
        End If
    End If

    GetTaskStatusFromInterfaceMessage = strStatus
    'MAR847 End

GetTaskStatusFromInterfaceMessageVbErr:
    Set cmd = Nothing
    Set rst = Nothing
    errCheckError cstrFunction, TypeName(Me)
End Function

Private Sub GetNewNumbers( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo GetNewNumbersErr

    Const cstrFunctionName As String = "GetNewNumbers"
    
    Dim strApplNo As String, strResponse As String, strAccountNumber As String
    Dim strOtherSysAccountNumber As String, strOtherSysCustomerNumber As String, strAFFNo As String
    Dim strCustomerNo As String
    Dim xmlApplication As IXMLDOMElement, xmlNewRequest As IXMLDOMNode, xmlElem As IXMLDOMElement
    Dim xmlTable As IXMLDOMElement, xmlCustomer As IXMLDOMNode, xmlCaseTask As IXMLDOMElement
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim objApplManagerBO As Object, objApplicationBO As Object, objAdminInterfaceBO As Object
    Dim objCustomerBO As Object, objMsgTmBO As Object
    
    'Build Request and Get Application data
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "./APPLICATION")
    strApplNo = xmlGetAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAFFNo = xmlGetAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApplication = xmlDoc.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlApplication
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strApplNo
    xmlApplication.appendChild xmlElem
    
    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Set objAdminInterfaceBO = gobjContext.CreateInstance("omAdmin.AdminInterfaceBO")
    
    strResponse = objApplicationBO.GetApplicationData(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    strOtherSysAccountNumber = xmlDoc.selectSingleNode("//APPLICATION/OTHERSYSTEMACCOUNTNUMBER").Text
    
    '----------------------------------------------------------------------------------
    ' Get 'OtherSystemAccountNumber'(if required) and update the Application table
    '----------------------------------------------------------------------------------
    If strOtherSysAccountNumber = "" Then
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETNEWNUMBERS"
        
        Set xmlTable = xmlDoc.createElement("NUMBERREQUEST")
        xmlNewRequest.appendChild xmlTable
        
        xmlSetAttributeValue xmlTable, "NUMBERREQUIRED", strApplNo
        xmlSetAttributeValue xmlTable, "NUMBERTYPEREQUIRED", "A"
        xmlSetAttributeValue xmlTable, "SOURCE", "Omiga"
        
        strResponse = objAdminInterfaceBO.REQUEST(xmlNewRequest.xml)
        errCheckXMLResponse strResponse, True, vxmlResponseNode
        
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        Set xmlTable = xmlDoc.selectSingleNode(".//OTHERSYSTEMNUMBER")
        strAccountNumber = xmlGetAttributeText(xmlTable, "OTHERSYSTEMNUMBER")
        
        ' Update Application with new OtherSystemAccountNumber
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        Set xmlTable = xmlDoc.createElement("APPLICATION")
        xmlNewRequest.appendChild xmlTable
        
        Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlElem.Text = strApplNo
        xmlTable.appendChild xmlElem
        
        Set xmlElem = xmlDoc.createElement("OTHERSYSTEMACCOUNTNUMBER")
        xmlElem.Text = strAccountNumber
        xmlTable.appendChild xmlElem
        
        strResponse = objApplicationBO.Update(xmlNewRequest.xml)
        errCheckXMLResponse strResponse, True, vxmlResponseNode
    End If
    
    '------------------------------------------------------------------------------
    ' Get 'OtherSystemCustomerNumber'(if required) and update the Customer table
    '------------------------------------------------------------------------------
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlTable = xmlDoc.createElement("CUSTOMERROLE")
    xmlNewRequest.appendChild xmlTable
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strApplNo
    xmlTable.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = strAFFNo
    xmlTable.appendChild xmlElem
    
    Set objApplManagerBO = gobjContext.CreateInstance("omApp.ApplicationManagerBO")
    strResponse = objApplManagerBO.FindCustomersForApplication(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    For Each xmlCustomer In xmlDoc.selectNodes("RESPONSE/CUSTOMER")
        strCustomerNo = xmlCustomer.selectSingleNode(".//CUSTOMERNUMBER").Text
        strOtherSysCustomerNumber = xmlCustomer.selectSingleNode(".//OTHERSYSTEMCUSTOMERNUMBER").Text
        
        'if not stored in Customer Table, get it from AdminInterface and update Customer table
        If strOtherSysCustomerNumber = "" Then
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            Set xmlTable = xmlDoc.createElement("NUMBERREQUEST")
            xmlNewRequest.appendChild xmlTable
        
            xmlSetAttributeValue xmlTable, "NUMBERREQUIRED", strCustomerNo
            xmlSetAttributeValue xmlTable, "NUMBERTYPEREQUIRED", "C"
            xmlSetAttributeValue xmlTable, "SOURCE", "Omiga"
        
            strResponse = objAdminInterfaceBO.REQUEST(xmlNewRequest.xml)
            Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        
            Set xmlTable = xmlDoc.selectSingleNode(".//OTHERSYSTEMNUMBER")
            strOtherSysCustomerNumber = xmlGetAttributeText(xmlTable, "OTHERSYSTEMNUMBER")
            
            'Update Customer table
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            Set xmlTable = xmlDoc.createElement("CUSTOMER")
            xmlNewRequest.appendChild xmlTable
            
            Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlElem.Text = strCustomerNo
            xmlTable.appendChild xmlElem
            
            Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlElem.Text = "1"
            xmlTable.appendChild xmlElem
            
            Set xmlElem = xmlDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
            xmlElem.Text = strOtherSysCustomerNumber
            xmlTable.appendChild xmlElem
            
            Set objCustomerBO = gobjContext.CreateInstance("omCust.CustomerBO")
            strResponse = objCustomerBO.UpdateCustomer(xmlNewRequest.xml)
            errCheckXMLResponse strResponse, True, vxmlResponseNode
        End If
    Next xmlCustomer
        
    '-----------------------------------------------
    'Update CaseTask as complete
    '-----------------------------------------------
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")

    Dim strUserId As String, strUnitId As String, strUserAuthorityLevel As String
    strUserId = xmlGetAttributeText(xmlNewRequest, "USERID")
    strUnitId = xmlGetAttributeText(xmlNewRequest, "UNITID")
    strUserAuthorityLevel = xmlGetAttributeText(xmlNewRequest, "USERAUTHORITYLEVEL")

    xmlSetAttributeValue xmlCaseTask, "USERID", strUserId
    xmlSetAttributeValue xmlCaseTask, "UNITID", strUnitId
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", strUserAuthorityLevel
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
GetNewNumbersExit:
    
    Set objApplicationBO = Nothing
    Set objApplManagerBO = Nothing
    Set objAdminInterfaceBO = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    Set xmlApplication = Nothing
    Set xmlTable = Nothing
    Set xmlElem = Nothing
    Set xmlCustomer = Nothing
    Set xmlCaseTask = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
GetNewNumbersErr:
    gobjContext.SetComplete
    
    GoTo GetNewNumbersExit
End Sub


Private Sub CreateDefaultApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode)
    
    On Error GoTo CreateDefaultApplicationConditionsErr

    Const cstrFunctionName As String = "CreateDefaultApplicationConditions"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMElement
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim objBO As Object
    Dim colValueIDs As Collection
    Dim strValueId As String

    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    
    ' validate REQUEST
    Set xmlTableNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    strApplicationNumber = xmlGetAttributeText(xmlTableNode, "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(xmlTableNode, "APPLICATIONFACTFINDNUMBER")
    
    If strApplicationNumber = "" Then
        errThrowError cstrFunctionName, oeMissingElementValue, "Application Number"
    End If
    
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    
    ' Build the request and call 'Conditions.DeleteAllApplicationConditions'
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "DELETEALLAPPLICATIONCONDITIONS"
    
    Set xmlTableNode = xmlDoc.createElement("APPLICATIONCONDITIONS")
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber

    xmlSetAttributeValue xmlTableNode, "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
    xmlSetAttributeValue xmlTableNode, "USERMODIFIED", "0"

    Set colValueIDs = New Collection
    GetValueIdsForValidationType "ConditionType", "STD", colValueIDs
    'Extract first value id
    strValueId = colValueIDs.Item(1)
    xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", strValueId
    xmlNewRequest.appendChild xmlTableNode
    
    'Call Delete for the first time
    xmlThisResponseDoc.loadXML (objBO.OmAppProcRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    'Set strProductCondType = Get ValueID for ValidationType 'P' (Product Condition)
    Set colValueIDs = New Collection
    GetValueIdsForValidationType "ConditionType", "P", colValueIDs
    'Extract first value id
    strValueId = colValueIDs.Item(1)
    xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", strValueId
    
    'Call Delete for the second time
    xmlThisResponseDoc.loadXML (objBO.OmAppProcRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    'Build the request and call 'ConditionsBO.CreateDefaultApplicationConditions'
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEDEFAULTAPPLICATIONCONDITIONS"
    xmlThisResponseDoc.loadXML (objBO.OmAppProcRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    'Update conditions task as pending
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    xmlNewRequest.appendChild vxmlRequestNode.selectSingleNode(".//CASETASK")
    Set xmlTableNode = xmlNewRequest.selectSingleNode(".//CASETASK")
    xmlSetAttributeValue xmlTableNode, "TASKSTATUS", 20
    
    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    xmlThisResponseDoc.loadXML (objBO.TmRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    gobjContext.SetComplete
    
CreateDefaultApplicationConditionsExit:
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateDefaultApplicationConditionsErr:
    gobjContext.SetAbort
    
    GoTo CreateDefaultApplicationConditionsExit

End Sub
Private Sub CreateApplication( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CreateApplicationExit
    
    Const cstrFunctionName As String = "CreateApplication"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCustomerRequestNode As IXMLDOMNode
    Dim xmlApplicationFactFindNode As IXMLDOMNode ' PJO 14/12/2005 MAR850
    Dim xmlCaseActivityRequestNode As IXMLDOMNode
    Dim xmlCreateApplicationAndCustomerResponseNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlSrceNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim objBO As Object
    
    ' validate request ================================================================
    xmlCheckMandatoryAttribute vxmlRequestNode, "UNITID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "CHANNELID"
    
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    'PJO 14/12/2005 MAR850
    Set xmlApplicationFactFindNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATIONFACTFIND")
    
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "TYPEOFAPPLICATION"
'---- SYS1950 TJ
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "TYPEOFBUYER"
    
    Set xmlCustomerRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CUSTOMER")
    xmlCheckMandatoryAttribute xmlCustomerRequestNode, "CUSTOMERROLETYPE"
    xmlCheckMandatoryAttribute xmlCustomerRequestNode, "CUSTOMERORDER"
    xmlCheckMandatoryAttribute xmlCustomerRequestNode, "CUSTOMERNUMBER"
    xmlCheckMandatoryAttribute xmlCustomerRequestNode, "CUSTOMERVERSIONNUMBER"
    
    Set xmlCaseActivityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "ACTIVITYID"

    ' create CreateApplicationAndCustomer request ======================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    Set xmlElem = xmlThisRequestDoc.createElement("REQUEST")
    xmlElem.setAttribute "OPERATION", "CREATE"
    xmlCopyAttribute vxmlRequestNode, xmlElem, "UNITID"
    xmlCopyAttribute vxmlRequestNode, xmlElem, "USERID"
    xmlCopyAttribute vxmlRequestNode, xmlElem, "CHANNELID"
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    Set xmlThisRequestNode = xmlThisRequestNode.appendChild(xmlElem)
    
    xmlElemFromAttrib xmlThisRequestNode, xmlCustomerRequestNode, "CUSTOMERNUMBER"
    xmlElemFromAttrib xmlThisRequestNode, xmlCustomerRequestNode, "CUSTOMERVERSIONNUMBER"
    xmlElemFromAttrib xmlThisRequestNode, xmlCustomerRequestNode, "CUSTOMERROLETYPE"
    xmlElemFromAttrib xmlThisRequestNode, xmlCustomerRequestNode, "CUSTOMERORDER"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "TYPEOFAPPLICATION"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "TYPEOFBUYER"
    ' PJO 14/12/2005 MAR850
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationFactFindNode, "DIRECTINDIRECTBUSINESS"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "CHANNELID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "CHANNELID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "USERID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "UNITID"
    
    ' APS SYS1993
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "APPLICATIONPRIORITYVALUE"

    ' JLD SYS2839 required for internet system
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOCKING"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "RESIDENCYOKINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "AGEOKINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOANAMOUNTINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOANTERMINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "PRIMARYRESIDENCEINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "MARKETINGSOURCE"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "INTENDEDAPPLICANTCOUNT"
     
    'AM 30/09/2005 MAR57
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "REGULATIONINDICATOR"

    ' run CreateApplicationAndCustomer ==============================================
    
    Debug.Print "CreateApplicationAndCustomer request:"
    Debug.Print xmlThisRequestDoc.xml
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    
    Set objBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    xmlThisResponseDoc.loadXML objBO.CreateApplicationAndCustomer(xmlThisRequestDoc.xml)
    Set objBO = Nothing
    
    Debug.Print "CreateApplicationAndCustomer response:"
    Debug.Print xmlThisResponseDoc.xml
    
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONNUMBER"
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONFACTFINDNUMBER"
    
    Set xmlCreateApplicationAndCustomerResponseNode = _
        xmlThisResponseDoc.selectSingleNode("RESPONSE")

    ' create CreateActivity request =================================================
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST")
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "CreateActivity"
    
    ' clone CASEACTIVITY node from original request
    Set xmlNode = _
        xmlThisRequestNode.appendChild(xmlCaseActivityRequestNode.cloneNode(False))
    
    ' get APPLICATIONNUMBER node from CreateApplicationAndCustomer response
    Set xmlSrceNode = xmlCreateApplicationAndCustomerResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    
    ' create CASEID attribute from APPLICATIONNUMBER
    xmlSetAttributeValue xmlNode, "CASEID", xmlSrceNode.Text
    
    ' create APPLICATION node from CreateApplicationAndCustomer response
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    
    Set xmlSrceNode = xmlCreateApplicationAndCustomerResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    Set xmlSrceNode = xmlCreateApplicationAndCustomerResponseNode.selectSingleNode("APPLICATION/APPLICATIONFACTFINDNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    'xmlThisRequestNode.appendChild xmlCaseActivityRequestNode.cloneNode(False)
    xmlThisRequestNode.appendChild xmlElem.cloneNode(False)
    
    ' run CreateActivity ============================================================
    
    Debug.Print "CreateActivity request:"
    Debug.Print xmlThisRequestNode.xml
    
    CreateActivity xmlThisRequestNode, vxmlResponseNode
    
CreateApplicationExit:
    
    Set xmlApplicationRequestNode = Nothing
    Set xmlCustomerRequestNode = Nothing
    Set xmlCaseActivityRequestNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub CreateActivity(ByVal vxmlRequestNode As IXMLDOMNode, _
                            Optional ByVal vxmlResponseNode As IXMLDOMNode = Nothing)

    On Error GoTo CreateActivityExit
    
    Const cstrFunctionName As String = "CreateActivity"
    
    Dim xmlCaseActivityRequestNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    
    Dim xmlNextStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
    Dim xmlNextCaseStageDoc As FreeThreadedDOMDocument40
    Dim xmlNextCaseStageNode As IXMLDOMNode
    Dim xmlRbResponseNode As IXMLDOMNode
    
    Dim objCommonBO As IomTmCommonBO
      
    ' validate request
    Set xmlCaseActivityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
    
    ' get next CASESTAGE detail
    Set xmlNextStageDoc = New FreeThreadedDOMDocument40
    xmlNextStageDoc.validateOnParse = False
    xmlNextStageDoc.setProperty "NewParser", True
    xmlNextStageDoc.async = False
    GetNextStage vxmlRequestNode, xmlCaseActivityRequestNode, xmlNextStageDoc
    Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
        
    If AnyRulesForStage(xmlNextStageNode) = True Then

        'Get APPLICATION data
        Set xmlApplicationDataDoc = New FreeThreadedDOMDocument40
        xmlApplicationDataDoc.validateOnParse = False
        xmlApplicationDataDoc.setProperty "NewParser", True
        xmlApplicationDataDoc.async = False
        Set objCommonBO = gobjContext.CreateInstance(App.Title & ".OmTmCommonBO")
        objCommonBO.GetApplicationData xmlApplicationRequestNode, xmlApplicationDataDoc, "TMRBTemplate"
        Set objCommonBO = Nothing
        Set xmlRbResponseNode = xmlApplicationDataDoc.selectSingleNode("RESPONSE")
        
        'Run rules
        If xmlNextCaseStageDoc Is Nothing Then
            Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
            xmlNextCaseStageDoc.validateOnParse = False
            xmlNextCaseStageDoc.setProperty "NewParser", True
            xmlNextCaseStageDoc.async = False
        Else
            xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
        End If
        RunUserRules _
            vxmlRequestNode, _
            xmlNextStageNode, _
            xmlRbResponseNode, _
            xmlNextCaseStageDoc
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    Else
        If xmlNextCaseStageDoc Is Nothing Then
            Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
            xmlNextCaseStageDoc.validateOnParse = False
            xmlNextCaseStageDoc.setProperty "NewParser", True
            xmlNextCaseStageDoc.async = False
        Else
            xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
        End If
        CreateCaseStageFromStage _
            xmlNextStageNode, _
            xmlNextCaseStageDoc
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
    End If
    
    'Run CreateActivity request
    'EP2_902 pass gobjContext as parameter
    AddDefaultValuesToCaseStage gobjContext, vxmlRequestNode, xmlNextCaseStageNode, xmlNextStageNode
    
    CreateActivityRequest _
        vxmlRequestNode, _
        xmlCaseActivityRequestNode, _
        xmlNextCaseStageNode
    
    'Call omApp.ApplicationBO to create application stage on omiga
    CreateApplicationStage _
        vxmlRequestNode, _
        xmlApplicationRequestNode, _
        xmlNextStageNode, _
        xmlNextCaseStageNode, _
        vxmlResponseNode
        
CreateActivityExit:

    ' clean up
    Set xmlCaseActivityRequestNode = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlNextStageNode = Nothing
    Set xmlRbResponseNode = Nothing
    Set xmlNextCaseStageNode = Nothing
    Set xmlNextStageDoc = Nothing
    Set xmlApplicationDataDoc = Nothing
    Set xmlNextCaseStageDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ReinstatePreviousStage(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo ReinstatePreviousStageExit

    Const cstrFunctionName As String = "ReinstatePreviousStage"
    
    Dim xmlRequestCaseActivityNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlPreviousStageDoc As FreeThreadedDOMDocument40
    Dim xmlPreviousStageNode As IXMLDOMNode
    Dim xmlPreviousStageDetailDoc As FreeThreadedDOMDocument40
    Dim xmlPreviousStageDetailNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisOperationNode As IXMLDOMNode
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlThisTaskNode As IXMLDOMNode
    Dim xmlNextStageNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    Dim strPattern As String
    
    ' validate request =================================================================
    Set xmlRequestCaseActivityNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlRequestCaseActivityNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseActivityNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseActivityNode, "ACTIVITYID"
    
    ' do GetCurrentStage request =======================================================
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlGetMandatoryNode(xmlCurrentStageDoc, "RESPONSE/CASESTAGE")
    
    ' error if current stage is not 'exception' (e.g. Cancel / Decline) stage
    If xmlGetAttributeAsBoolean(xmlCurrentStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
        errThrowError cstrFunctionName, oeTmNotExceptionStage
    End If
    
    ' get previous CaseStage details =======================================================
    ' create FindArchiveStageList request
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    Set xmlPreviousStageDoc = New FreeThreadedDOMDocument40
    xmlPreviousStageDoc.validateOnParse = False
    xmlPreviousStageDoc.setProperty "NewParser", True
    xmlPreviousStageDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(True))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "FindArchiveStageList"
    
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST/CASEACTIVITY")
    
    ' add ACTIVITYINSTANCE to CASEACTIVITY request
    xmlCopyAttribute xmlCurrentStageNode, xmlThisRequestNode, "ACTIVITYINSTANCE"
    
    ' add CASESTAGESEQUENCENO (current CASESTAGESEQUENCENO - 1) to CASEACTIVITY request
    xmlSetAttributeValue _
        xmlThisRequestNode, _
        "CASESTAGESEQUENCENO", _
        CStr(xmlGetAttributeAsInteger(xmlCurrentStageNode, "CASESTAGESEQUENCENO") - 1)
        
    RunTmRequest xmlThisRequestDoc, xmlPreviousStageDoc
    Set xmlPreviousStageNode = xmlPreviousStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    If xmlPreviousStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoStageDetail
    End If
    
    ' get Stage details for previous CaseStage =========================================
    Set xmlPreviousStageDetailDoc = New FreeThreadedDOMDocument40
    xmlPreviousStageDetailDoc.validateOnParse = False
    xmlPreviousStageDetailDoc.setProperty "NewParser", True
    xmlPreviousStageDetailDoc.async = False
    ' create GetStageDetail request
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST")
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageDetail"
    Set xmlElem = xmlThisRequestDoc.createElement("STAGE")
    xmlCopyAttribute xmlRequestCaseActivityNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlPreviousStageNode, xmlElem, "STAGEID"
    xmlThisRequestNode.appendChild xmlElem
    ' run request & check response
    RunTmRequest xmlThisRequestDoc, xmlPreviousStageDetailDoc
    Set xmlPreviousStageDetailNode = _
        xmlPreviousStageDetailDoc.selectSingleNode("RESPONSE/STAGE")
    If xmlPreviousStageDetailNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoStageDetail, "Previous Stage Details Not Found"
    End If
    
    ' create multiple operation request block ==========================================
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST")
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    xmlThisRequestNode.Attributes.removeNamedItem "OPERATION"
    
    'DPF 19/11/02 -  have amended this block to only be added if we have incomplete tasks
    'n the current stage, otherwise an error is created in another method
    Set xmlElem = xmlThisRequestDoc.createElement("OPERATION")
    xmlElem.setAttribute "NAME", "UpdateCaseTask"
    ProcessNotApplicableTasks xmlCurrentStageNode, xmlElem
    
    If xmlElem.hasChildNodes = True Then
        Set xmlThisOperationNode = xmlThisRequestNode.appendChild(xmlElem)
    End If
    
    ' create SetCurrentCaseStageComplete request
    Set xmlElem = xmlThisRequestDoc.createElement("OPERATION")
    xmlElem.setAttribute "NAME", "SetCurrentCaseStageComplete"
    Set xmlThisOperationNode = xmlThisRequestNode.appendChild(xmlElem)
    ' create CASESTAGE elem
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlThisOperationNode.appendChild xmlElem
    
    ' create CreateCaseStage request
    Set xmlElem = xmlThisRequestDoc.createElement("OPERATION")
    xmlElem.setAttribute "NAME", "CreateCaseStage"
    Set xmlThisOperationNode = xmlThisRequestNode.appendChild(xmlElem)
    
    ' create CASESTAGE request element
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
    
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlPreviousStageNode, xmlElem, "STAGEID"
    
    ' add CASESTAGESEQUENCENO (current CASESTAGESEQUENCENO + 1) to CASEACTIVITY request
    xmlSetAttributeValue _
        xmlElem, _
        "CASESTAGESEQUENCENO", _
        CStr(xmlGetAttributeAsInteger(xmlCurrentStageNode, "CASESTAGESEQUENCENO") + 1)
    
    Set xmlNextStageNode = xmlThisOperationNode.appendChild(xmlElem)
    
    ' examine previous tasks
    For Each xmlNode In xmlPreviousStageNode.childNodes
        
        If xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSCancelled Then
            
            Set xmlThisCaseTaskNode = _
                xmlNextStageNode.appendChild(xmlNode.cloneNode(False))
            
            xmlThisCaseTaskNode.Attributes.getNamedItem("TASKSTATUS").Text = _
                CStr(omiga4TASKSTATUSIncomplete)
                
            strPattern = _
                "STAGETASK[@TASKID='" & _
                xmlGetAttributeText(xmlThisCaseTaskNode, "TASKID") & _
                "']"
                
            Set xmlThisTaskNode = _
                xmlPreviousStageDetailNode.selectSingleNode(strPattern)
                
            If Not xmlThisTaskNode Is Nothing Then ' error ???
                SetCaseTaskDueDateAndTime xmlThisCaseTaskNode, xmlThisTaskNode
            End If
        
        End If
    
    Next
    
    RunTmRequest xmlThisRequestDoc
        
    ' call omApp.ApplicationBO to create application stage on omiga ====================
    CreateApplicationStage _
        vxmlRequestNode, _
        xmlApplicationRequestNode, _
        xmlPreviousStageDetailNode, _
        xmlNextStageNode
    
ReinstatePreviousStageExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestCaseActivityNode = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlPreviousStageNode = Nothing
    Set xmlPreviousStageDetailNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisOperationNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlThisTaskNode = Nothing
    Set xmlNextStageNode = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlPreviousStageDoc = Nothing
    Set xmlPreviousStageDetailDoc = Nothing
    Set xmlThisRequestDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub UpdateApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode)
    
On Error GoTo UpdateCaseTaskPriorityExit

    Const cstrFunctionName As String = "UpdateApplicationConditions"
    
    Dim xmlNodeList As IXMLDOMNodeList, xmlNode As IXMLDOMNode, xmlNewRequest As IXMLDOMNode
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    
    Dim strUserId As String, strUnitId As String, strResponse As String
    Dim strUserAuthorityLevel As String
    
    Dim objBO As Object
    
    Set xmlNewRequest = vxmlRequestNode.cloneNode(True)
    
    Set xmlNodeList = vxmlRequestNode.selectNodes(".//APPLICATIONCONDITIONS")
    If xmlNodeList.length <> 0 Then
        strUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
        strUnitId = xmlGetAttributeText(vxmlRequestNode, "UNITID")
        strUserAuthorityLevel = xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
        
        ' Assign UserId, UnitId values from Request to each ApplicationConditions node
        ' and update application conditions
        For Each xmlNode In xmlNewRequest.selectNodes(".//APPLICATIONCONDITIONS")
            xmlSetAttributeValue xmlNode, "USERID", strUserId
            xmlSetAttributeValue xmlNode, "UNITID", strUnitId
        Next xmlNode
   
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATEAPPLICATIONCONDITIONS"
    
        Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
        xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(xmlNewRequest.xml)
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    End If
    
    ' Update case task as complete
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
    Set xmlNode = xmlNewRequest.selectSingleNode(".//CASETASK")
    xmlSetAttributeValue xmlNode, "USERID", strUserId
    xmlSetAttributeValue xmlNode, "UNITID", strUnitId
    xmlSetAttributeValue xmlNode, "USERAUTHORITYLEVEL", strUserAuthorityLevel
    xmlSetAttributeValue xmlNode, "TASKSTATUS", 40
    
    Set objBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    xmlThisResponseDoc.loadXML objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
UpdateCaseTaskPriorityExit:
    
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlThisResponseDoc = Nothing
    Set objBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub UpdateCaseTaskPriority(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo UpdateCaseTaskPriorityExit

    Const cstrFunctionName As String = "UpdateCaseTaskPriority"
    
    Dim xmlCaseStageRequestNode As IXMLDOMNode
    Dim xmlPriorityRequestNode As IXMLDOMNode
    
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlOutstandingCaseTaskList As IXMLDOMNodeList
    Dim xmlOutstandingCaseTaskNode As IXMLDOMNode
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    
    Dim gobjEnquiry As Object
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlSysDateNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim strTmpResponse As String
    Dim strNonWorkingDay As String
    
    Dim xmlStageTaskDoc As FreeThreadedDOMDocument40
    Dim xmlStageNode As IXMLDOMNode
    Dim xmlStageTaskNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    
    Dim strXSLPattern As String, _
        strPriority As String
        
    Dim lngOffset As Long
    
    Dim dtDueDateTime As Date
    
    ' validate request =================================================================
    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    Set xmlPriorityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "PRIORITY")
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlPriorityRequestNode, "APPLICATIONPRIORITY"
    
    ' get current CASESTAGE detail =====================================================
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    
    ' create GetStageDetail request ====================================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    ' clone request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set OPERATION attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageDetail"
    ' create STAGETASK node
    Set xmlElem = xmlThisRequestDoc.createElement("STAGE")
    xmlCopyAttribute xmlCurrentStageNode, xmlElem, "STAGEID"
    xmlCopyAttributeValue xmlPriorityRequestNode, xmlElem, "APPLICATIONPRIORITY", "CASEPRIORITY"
    xmlThisRequestNode.appendChild xmlElem
    ' run TmRequest
    Set xmlStageTaskDoc = New FreeThreadedDOMDocument40
    xmlStageTaskDoc.validateOnParse = False
    xmlStageTaskDoc.setProperty "NewParser", True
    xmlStageTaskDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlStageTaskDoc
    Set xmlStageNode = xmlStageTaskDoc.selectSingleNode("RESPONSE/STAGE")
    If xmlStageNode Is Nothing Then
        'JD MAR1188 errThrowError cstrFunctionName, oeTmNoTaskDetail
        errThrowError cstrFunctionName, oeTmNoStageDetail, "StageID = " + xmlGetAttributeText(xmlCurrentStageNode, "STAGEID")
    End If
    
    Set xmlOutstandingCaseTaskList = _
        xmlCurrentStageNode.selectNodes("CASETASK[@TASKSTATUS='10' or @TASKSTATUS='20']")
        
    If xmlOutstandingCaseTaskList.length <> 0 Then
    
        strPriority = xmlGetAttributeText(xmlPriorityRequestNode, "APPLICATIONPRIORITY")
        strXSLPattern = "[@CASEPRIORITY='" & strPriority & "']"
        
        ' create UpdateCaseTask request(s) =============================================
        ' drop existing (GetStageTaskList) request
        xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
        ' set OPERATION attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
        
        For Each xmlOutstandingCaseTaskNode In xmlOutstandingCaseTaskList
        
            ' PSC 02/03/2006 MAR1341
            strXSLPattern = _
                "STAGETASK[@TASKID='" & _
                xmlGetAttributeText(xmlOutstandingCaseTaskNode, "TASKID") & _
                "']" & _
                "[@CASEPRIORITY='" & strPriority & "']" & _
                "[@ADJUSTMENTDAYS != '0' or @ADJUSTMENTHOURS != '0' or @ADJUSTMENTMINUTES != '0']"

            Set xmlStageTaskNode = xmlStageNode.selectSingleNode(strXSLPattern)
            
            If Not xmlStageTaskNode Is Nothing Then
            
                dtDueDateTime = _
                    xmlGetAttributeAsDate(xmlOutstandingCaseTaskNode, "TASKDUEDATEANDTIME")
            
                ' PSC 02/03/2006 MAR1341
                lngOffset = _
                    (xmlGetAttributeAsInteger(xmlStageTaskNode, "ADJUSTMENTDAYS") * 24 * 60) + _
                    (xmlGetAttributeAsInteger(xmlStageTaskNode, "ADJUSTMENTHOURS") * 60) + _
                    xmlGetAttributeAsInteger(xmlStageTaskNode, "ADJUSTMENTMINUTES")
    
                If lngOffset <> 0 Then
                    ' PSC 02/03/2006 MAR1341
                    dtDueDateTime = DateAdd("n", lngOffset, dtDueDateTime)
                End If
                '
                ' Check for non working day
                '
                Set xmlSysDateNode = xmlThisRequestDoc.createElement("SYSTEMDATE")
                xmlThisRequestNode.appendChild xmlSysDateNode
                Set xmlTempNode = xmlThisRequestDoc.createElement("DATE")
                xmlTempNode.Text = CStr(dtDueDateTime)
                xmlSysDateNode.appendChild xmlTempNode
                Set xmlTempNode = xmlThisRequestDoc.createElement("CHANNELID")
                xmlTempNode.Text = xmlGetAttributeText(vxmlRequestNode, "CHANNELID")
                xmlSysDateNode.appendChild xmlTempNode
                '
                ' Calling omBase.SystemDateBO.CheckNonWorkingOccurence
                '
                Set gobjEnquiry = gobjContext.CreateInstance("omBase.SystemDatesBO")
                strTmpResponse = gobjEnquiry.CheckNonWorkingOccurence(xmlSysDateNode.xml)
                Set xmlOutDoc = xmlLoad(strTmpResponse, cstrFunctionName)
                '
                ' Error Check
                '
                Set xmlTempNode = xmlOutDoc.firstChild
                errCheckXMLResponseNode xmlTempNode, , True
            
                '
                ' Check that it's a working day and if not, move forward till it is
                '
                Set xmlTempNode = xmlOutDoc.selectSingleNode("RESPONSE/SYSTEMDATE")
                strNonWorkingDay = xmlGetNodeText(xmlTempNode, "NONWORKINGIND")
                If strNonWorkingDay = "1" Then
                    Set xmlTempNode = xmlThisRequestDoc.createElement("DIRECTION")
                    xmlTempNode.Text = "+"
                    xmlSysDateNode.appendChild xmlTempNode
                    '
                    ' Calling omBase.SystemDateBO.FindWorkingDay
                    '
                    strTmpResponse = gobjEnquiry.FindWorkingDay(xmlSysDateNode.xml)
                    Set xmlOutDoc = xmlLoad(strTmpResponse, cstrFunctionName)
                    '
                    ' Error Check
                    '
                    Set xmlTempNode = xmlOutDoc.firstChild
                    errCheckXMLResponseNode xmlTempNode, , True
                    Set xmlTempNode = xmlOutDoc.selectSingleNode("RESPONSE/SYSTEMDATE")
                    dtDueDateTime = xmlGetNodeAsDate(xmlTempNode, "DATE")
                End If
                xmlThisRequestNode.removeChild xmlSysDateNode
                
                If dtDueDateTime < Now() Then
                    dtDueDateTime = Now()
                End If
                
                
                Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
                
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "CASEID"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "ACTIVITYID"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "STAGEID"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "TASKID"
                xmlCopyAttribute xmlOutstandingCaseTaskNode, xmlElem, "TASKINSTANCE"
                
                xmlSetAttributeValue xmlElem, "TASKDUEDATEANDTIME", dtDueDateTime
                
                xmlThisRequestNode.appendChild xmlElem
            
            End If
        
        Next
        
        If xmlThisRequestNode.hasChildNodes Then
            ' run UpdateCaseTask request ===============================================
            RunTmRequest xmlThisRequestDoc
        
        End If
        
    End If
    
UpdateCaseTaskPriorityExit:
    
    Set xmlElem = Nothing
    Set xmlCaseStageRequestNode = Nothing
    Set xmlPriorityRequestNode = Nothing
    Set xmlOutstandingCaseTaskList = Nothing
    Set xmlOutstandingCaseTaskNode = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlStageNode = Nothing
    Set xmlStageTaskNode = Nothing
    Set xmlStageTaskDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ChaseUpTask(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo ChaseUpTaskExit

    Const cstrFunctionName As String = "ChaseUpTask"
    
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlUpdateCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlCreateCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskDetailDoc As FreeThreadedDOMDocument40
    Dim xmlChaseTaskTemplateDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisCaseTaskDetailNode As IXMLDOMNode
    Dim xmlChaseTaskTemplateNode As IXMLDOMNode
    Dim xmlChaseTaskDetailNode As IXMLDOMNode
    Dim xmlUpdateCaseTaskRequestNode As IXMLDOMNode
    Dim xmlCreateCaseTaskRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    
    ' validate request ================================================================
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    
    ' get this CaskTask detail =========================================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetCaseTasks"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlThisRequestNode.appendChild xmlElem
    
    ' run Tm GetCaseTasks request
    Set xmlThisCaseTaskDetailDoc = New FreeThreadedDOMDocument40
    xmlThisCaseTaskDetailDoc.validateOnParse = False
    xmlThisCaseTaskDetailDoc.setProperty "NewParser", True
    xmlThisCaseTaskDetailDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisCaseTaskDetailDoc
    Set xmlThisCaseTaskDetailNode = _
        xmlThisCaseTaskDetailDoc.selectSingleNode("RESPONSE/CASETASK")
    If xmlThisCaseTaskDetailNode Is Nothing Then
        'JD MAR1188 add more detail to error
        errThrowError cstrFunctionName, oeTmNoTaskDetail, "TaskID = " + xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKID")
    End If
    
    ' get Task template detail for chasing task =======================================
    ' drop existing request child node(s)
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskList"
    ' create Task element
    Set xmlElem = xmlThisRequestDoc.createElement("TASK")
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttributeValue xmlThisCaseTaskDetailNode, xmlElem, "CHASINGTASK", "TASKID"
    xmlThisRequestNode.appendChild xmlElem
    ' create response doc.
    Set xmlChaseTaskTemplateDoc = New FreeThreadedDOMDocument40
    xmlChaseTaskTemplateDoc.validateOnParse = False
    xmlChaseTaskTemplateDoc.setProperty "NewParser", True
    xmlChaseTaskTemplateDoc.async = False
    ' run Tm GetCaseTasks request
    RunTmRequest xmlThisRequestDoc, xmlChaseTaskTemplateDoc
    Set xmlChaseTaskTemplateNode = _
        xmlChaseTaskTemplateDoc.selectSingleNode("RESPONSE/TASK")
    If xmlChaseTaskTemplateNode Is Nothing Then
        'JD MAR1188 add more detail to message
        errThrowError cstrFunctionName, oeTmNoTaskDetail, "TaskID = " + xmlGetAttributeText(xmlThisCaseTaskDetailNode, "CHASINGTASK")
    End If
    
    ' create UpdateCaseTask request ====================================================
    ' create chase-up task from new CaseTask details plus
    ' attributes from existing casetask
    Set xmlUpdateCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlUpdateCaseTaskDoc.validateOnParse = False
    xmlUpdateCaseTaskDoc.setProperty "NewParser", True
    Set xmlUpdateCaseTaskRequestNode = _
        xmlUpdateCaseTaskDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlUpdateCaseTaskRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CaseTask element
    Set xmlElem = xmlUpdateCaseTaskDoc.createElement("CASETASK")
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    
    'BG 07/11/01 SYS3333 wrong global used
    'xmlSetAttributeValue xmlElem, "TASKSTATUS", omiga4TASKSTATUSCarriedForward
     xmlSetAttributeValue xmlElem, "TASKSTATUS", omiga4TASKSTATUSChasedUp
     
    xmlUpdateCaseTaskRequestNode.appendChild xmlElem
 '   xmlSetAttributeValue xmlElem, "TASKNAME", xmlChaseTaskTemplateDoc.selectSingleNode("RESPONSE/TASK").Attributes.getNamedItem("TASKNAME").Text
    
    
    ' create CreateCaseTask request ====================================================
    ' create chase-up task from new CaseTask details plus
    ' attributes from existing casetask
    Set xmlCreateCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlCreateCaseTaskDoc.validateOnParse = False
    xmlCreateCaseTaskDoc.setProperty "NewParser", True
    Set xmlCreateCaseTaskRequestNode = _
        xmlCreateCaseTaskDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlCreateCaseTaskRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"
    ' create CaseTask element
    Set xmlElem = xmlCreateCaseTaskDoc.createElement("CASETASK")
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttributeValue xmlThisCaseTaskDetailNode, xmlElem, "CHASINGTASK", "TASKID"
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "TASKINSTANCE"
    ' BMIDS522 KRW 11/10/04
    xmlSetAttributeValue xmlElem, "CASETASKNAME", xmlChaseTaskTemplateDoc.selectSingleNode("RESPONSE/TASK").Attributes.getNamedItem("TASKNAME").Text
 
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "MANDATORYINDICATOR"
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "CONTEXT"
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "CUSTOMERIDENTIFIER"
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "ALLOCATEDBYUSERID"
    xmlCopyAttribute xmlThisCaseTaskDetailNode, xmlElem, "ORIGINATINGSTAGEID"
    SetCaseTaskDueDateAndTime xmlElem, xmlChaseTaskTemplateNode
    xmlCreateCaseTaskRequestNode.appendChild xmlElem
    
    ' run composite request
    Dim xmlUpdateNodes As Collection
    Set xmlUpdateNodes = New Collection
    xmlUpdateNodes.Add xmlUpdateCaseTaskRequestNode
    xmlUpdateNodes.Add xmlCreateCaseTaskRequestNode
    RunCompositeRequest xmlUpdateNodes
    Set xmlUpdateNodes = Nothing
    
ChaseUpTaskExit:

    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisCaseTaskDetailNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisCaseTaskDetailDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub CreateCaseStageTrigger(ByVal vxmlRequestNode As IXMLDOMNode)
On Error GoTo CreateCaseStageTriggerExit

    Const cstrFunctionName As String = "CreateCaseStageTrigger"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlTriggerTaskDoc As FreeThreadedDOMDocument40
    
    Dim xmlCaseTaskRequestNode As IXMLDOMNode
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlTriggerTaskNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    ' validate request ================================================================
    Set xmlCaseTaskRequestNode = vxmlRequestNode.selectSingleNode("CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "ACTIVITYINSTANCE"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "STAGEID"
    
    ' create GetStageTaskList request ==================================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    ' clone request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set OPERATION attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageTaskDetail"
    ' create STAGETASK node
    Set xmlElem = xmlThisRequestDoc.createElement("STAGETASK")
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "STAGEID"
    xmlElem.setAttribute "TRIGGERTASKIND", "1"
    xmlThisRequestNode.appendChild xmlElem
    ' run TmRequest
    Set xmlTriggerTaskDoc = New FreeThreadedDOMDocument40
    xmlTriggerTaskDoc.validateOnParse = False
    xmlTriggerTaskDoc.setProperty "NewParser", True
    xmlTriggerTaskDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlTriggerTaskDoc
    'BG 04/11/01 SYS3287 added [@TRIGGERTASKIND=""1""] condition to select the correct stagetask element
    Set xmlTriggerTaskNode = xmlTriggerTaskDoc.selectSingleNode("RESPONSE/STAGETASK[@TRIGGERTASKIND=""1""]")
    If xmlTriggerTaskNode Is Nothing Then
        'JD MAR1188 add more details to message.
        errThrowError cstrFunctionName, oeTmNoTaskDetail, "No trigger task on stage stageID = " + xmlGetAttributeText(xmlCaseTaskRequestNode, "STAGEID")
    End If
    
    ' create CreateCaseTask request ==================================================
    ' drop request child nodes
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    ' set OPERATION attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"
    ' create STAGETASK node
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlTriggerTaskNode, xmlElem, "TASKID"
    xmlCopyAttributeValue xmlTriggerTaskNode, xmlElem, "TASKNAME", "CASETASKNAME"
    'AW 28/09/06    EP1135
    SetCaseTaskDueDateAndTime xmlTriggerTaskNode, xmlTriggerTaskNode
    xmlElem.setAttribute "TASKDUEDATEANDTIME", xmlGetAttributeText(xmlTriggerTaskNode, "TASKDUEDATEANDTIME", "")
    'AW 28/09/06    EP1135  - End
    xmlThisRequestNode.appendChild xmlElem
    ' run TmRequest
    RunTmRequest xmlThisRequestDoc, xmlTriggerTaskDoc
    
CreateCaseStageTriggerExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlCaseTaskRequestNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlTriggerTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlTriggerTaskDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub CreateAdhocCaseTask(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo CreateAdhocCaseTaskExit

    Const cstrFunctionName As String = "CreateAdhocCaseTask"
    
    Dim xmlProcessTaskDoc As FreeThreadedDOMDocument40
    Dim xmlProcessTaskRequest As IXMLDOMNode
    Dim xmlProcessTasks As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlCaseTask As IXMLDOMNode
    Dim objNTxBO As IOmTmNTxBO
    Dim objNoTxBO As omTmNoTxBO
    Dim xmlRoot As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlPrinter As IXMLDOMNode        ' PSC 03/10/2005 MAR32

    Dim xmlTempResponse As FreeThreadedDOMDocument40
    
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseTaskDetails As IXMLDOMNode
    
    Dim strResponse As String
    Dim strAppNo As String
    
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNTxBO")
    Set xmlProcessTaskDoc = New FreeThreadedDOMDocument40
    xmlProcessTaskDoc.validateOnParse = False
    xmlProcessTaskDoc.setProperty "NewParser", True
    xmlProcessTaskDoc.async = False
    
    ' Pass this to a new transaction so that it is committed and avialable for the
    ' processing of tasks which are run automatically on creation
    Set xmlProcessTasks = objNTxBO.CreateAdhocCaseTask(vxmlRequestNode)
    Set xmlCaseTaskList = xmlProcessTasks.selectNodes("CASETASK")
    
    ' Need to get the application data
    If xmlCaseTaskList.length > 0 Then
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
        Set xmlCaseTask = xmlCaseTaskList.Item(0)
        strAppNo = xmlGetAttributeText(xmlCaseTask, "CASEID")

        Set xmlApplication = vxmlRequestNode.selectSingleNode("APPLICATION").cloneNode(True)
        xmlSetAttributeValue xmlApplication, "APPLICATIONNUMBER", strAppNo
        xmlSetAttributeValue xmlApplication, "APPLICATIONFACTFINDNUMBER", "1"   'MAR885 GHun
        Set xmlCustomer = vxmlRequestNode.selectSingleNode("CUSTOMER")
        ' PSC 03/10/2005 MAR32
        Set xmlPrinter = vxmlRequestNode.selectSingleNode("PRINTER")
    End If
    
    Set xmlProcessTaskRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlProcessTaskRequest, "OPERATION", "ProcessTask"
    
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetAdditionalCaseTaskData"
    Set xmlNode = xmlProcessTaskDoc.createElement("CASETASK")
    xmlTempRequest.appendChild xmlNode
    
    For Each xmlCaseTask In xmlCaseTaskList
            
        xmlCopyAttribute xmlCaseTask, xmlNode, "CASEACTIVITYGUID"
        xmlCopyAttribute xmlCaseTask, xmlNode, "STAGEID"
        xmlCopyAttribute xmlCaseTask, xmlNode, "CASESTAGESEQUENCENO"
        xmlCopyAttribute xmlCaseTask, xmlNode, "TASKID"
        xmlCopyAttribute xmlCaseTask, xmlNode, "TASKINSTANCE"
        
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        Set xmlTempResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlTempResponse.documentElement, , True
        
        Set xmlCaseTaskDetails = xmlTempResponse.selectSingleNode("RESPONSE/CASETASK")
        
        Set xmlProcessTaskRequest = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlProcessTaskRequest, "OPERATION", "ProcessTask"
        xmlProcessTaskRequest.appendChild xmlCaseTaskDetails
        xmlSetAttributeValue xmlCaseTaskDetails, "COMPLETETASKONPRINT", "1"
        xmlProcessTaskRequest.appendChild xmlApplication.cloneNode(True)
        
        If Not xmlCustomer Is Nothing Then
            xmlProcessTaskRequest.appendChild xmlCustomer.cloneNode(True)
        End If
        
        ' PSC 03/10/2005 MAR32 - Start
        If Not xmlPrinter Is Nothing Then
            xmlProcessTaskRequest.appendChild xmlPrinter.cloneNode(True)
        End If
        ' PSC 03/10/2005 MAR32 - End
    
        On Error GoTo ProcessTaskError
        objNTxBO.ProcessTask xmlProcessTaskRequest
        On Error GoTo CreateAdhocCaseTaskExit
    Next
    
    gobjContext.SetComplete

CreateAdhocCaseTaskExit:
    'Debug.Print Err.Description
    Set xmlProcessTaskDoc = Nothing
    Set xmlProcessTaskRequest = Nothing
    Set xmlProcessTasks = Nothing
    Set xmlCaseTaskList = Nothing
    Set xmlCaseTask = Nothing
    Set objNTxBO = Nothing
    Set xmlRoot = Nothing
    Set xmlApplication = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTaskDetails = Nothing
    Set xmlCustomer = Nothing
    Set xmlPrinter = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateAdhocCaseTaskError:
    gobjContext.SetAbort
    GoTo CreateAdhocCaseTaskExit

ProcessTaskError:
    Err.Clear
    Resume Next
End Sub

Private Sub CriticalDataCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                              ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo CriticalDataCheckExit

    Const cstrFunctionName As String = "CriticalDataCheck"
    
    'Dim xmlApplicationRequestNode As IXMLDOMNode ' BC MAR842
    Dim xmlRequestNode As IXMLDOMNode ' BC MAR842
    
    Dim xmlContextNode As IXMLDOMNode
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlCurrentStageNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
    Dim xmlCriticalDataRequestDoc As FreeThreadedDOMDocument40
    Dim xmlCriticalDataRequestNode As IXMLDOMNode
    Dim xmlCriticalDataRulesResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCriticalDataRulesResponseNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlGlobalParameterList As IXMLDOMNodeList
    Dim xmlGPNode As IXMLDOMNode
    
    'Dim blnAfterCreditCheckStage As Boolean
    Dim blnCompletedCreditCheck As Boolean, _
        blnOutstandingCreditCheck As Boolean, _
        blnOutstandingReScoreCreditCheckTask As Boolean, _
        blnOutstandingCostModelTask As Boolean, _
        blnCreditCheckTestRequired As Boolean, _
        blnCostModelTestRequired As Boolean, _
        blnDataUpdateRequired As Boolean
        
    Dim blnAfterHunterStage As Boolean              ' MAR1126
    Dim blnOutstandingHunterTask As Boolean         ' MAR1126
    Dim blnCompletedHunter As Boolean               ' MAR1126
    Dim blnHunterTestRequired As Boolean            ' MAR1126
    
    'Dim objNTxBO As IOmTmNTxBO  'SYS3270
    Dim objNoTxBO As omTmNoTxBO  'BC MAR842
    
    'BC MAR842 13/12/2005 Begin
    Dim xmlTempRequest As IXMLDOMNode
    'Dim xmlApplicationNode As IXMLDOMNode
    Dim strResponse As String
    'BC MAR842 13/12/2005 End
    
    'MAR1717 GHun
    Dim xmlAppCreditCheck   As IXMLDOMElement
    Dim xmlTemp             As IXMLDOMElement
    'MAR1717 End
    
    gobjContext.SetComplete

    ' validate request ================================================================
    Set xmlContextNode = vxmlRequestNode.selectSingleNode("CRITICALDATACONTEXT")
    xmlCheckMandatoryAttribute xmlContextNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlContextNode, "APPLICATIONFACTFINDNUMBER"
    xmlCheckMandatoryAttribute xmlContextNode, "APPLICATIONPRIORITY"
    xmlCheckMandatoryAttribute xmlContextNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlContextNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlContextNode, "ACTIVITYINSTANCE"
    xmlCheckMandatoryAttribute xmlContextNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlContextNode, "COMPONENT"
    xmlCheckMandatoryAttribute xmlContextNode, "METHOD"
    
    'MAR727 - check the context node to see if an update is required
    If UCase$(xmlGetAttributeText(xmlContextNode, "METHOD")) = "NONE" Then
      blnDataUpdateRequired = False
    Else
      blnDataUpdateRequired = True
    End If
    'MAR 727

    ' get current CASESTAGE detail =====================================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
    xmlCopyAttribute xmlContextNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlContextNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlContextNode, xmlElem, "STAGEID"
    xmlCopyAttributeValue xmlContextNode, xmlElem, "APPLICATIONNUMBER", "CASEID"
    xmlThisRequestNode.appendChild xmlElem

    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    GetCurrentStage xmlThisRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")

    ' credit check test required
    CheckTasksOnCurrentStage _
        xmlCurrentStageNode, _
        blnOutstandingCreditCheck, _
        blnOutstandingReScoreCreditCheckTask, _
        blnCompletedCreditCheck, _
        blnOutstandingCostModelTask, _
        blnOutstandingHunterTask, _
        blnCompletedHunter                            ' MAR1126
        
    'MAR1172 only check that an actual credit check has been rather than the stage and task
    '
'    If Not blnOutstandingCreditCheck Then
'        blnAfterCreditCheckStage = _
'            IsAfterCreditCheckStage(xmlGetAttributeText(xmlContextNode, "STAGEID"))
'    End If
'   If (Not blnOutstandingCreditCheck) And _
'  (blnAfterCreditCheckStage Or blnCompletedCreditCheck) _

    If (Not blnOutstandingCreditCheck) Then
        'MAR1717 GHun
        Set xmlAppCreditCheck = InitialCreditCheckDone(xmlContextNode)
        If Not xmlAppCreditCheck Is Nothing Then
            If xmlGetAttributeAsLong(xmlAppCreditCheck, "SEQUENCENUMBER") > 0 Then
                blnCreditCheckTestRequired = True
            End If
        End If
        'MAR1717 End
    End If
    
  'MAR1172 Ends
    ' cost modelling test required
    'MAR1126  Remove check for Cost Model required
    If Not blnOutstandingCostModelTask Then
        blnCostModelTestRequired = True
    End If

    'MAR1126 Check Hunter
    If Not blnOutstandingHunterTask Then
        blnAfterHunterStage = IsAfterHunterStage(xmlGetAttributeText(xmlContextNode, "STAGEID"))
    End If

    If (Not blnOutstandingHunterTask) And (blnAfterHunterStage Or blnCompletedHunter) Then
        blnHunterTestRequired = True
    End If

    If blnDataUpdateRequired Then  'MAR727 GHun
    
        If blnCreditCheckTestRequired Or blnCostModelTestRequired Or blnHunterTestRequired Then   ' MAR1126
          
        
           ' get Application before image
            Set xmlApplicationDataDoc = New FreeThreadedDOMDocument40
            xmlApplicationDataDoc.validateOnParse = False
            xmlApplicationDataDoc.setProperty "NewParser", True
            xmlApplicationDataDoc.async = False
            
'BC MAR842 13/12/2005 Begin
'            Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
'            objNTxBO.GetApplicationData xmlContextNode, xmlApplicationDataDoc, "CDRBTemplate"
'            Set objNTxBO = Nothing
           
            Set xmlRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CRITICALDATACONTEXT")
            Set xmlTempRequest = xmlThisRequestDoc.createElement("REQUEST")
            xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONNUMBER"
            xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONFACTFINDNUMBER"

            xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetApplicationData"
            ' run and then load and check response
            
            Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
            strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            Set xmlApplicationDataDoc = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlApplicationDataDoc.selectSingleNode("RESPONSE"), , True
            Set objNoTxBO = Nothing
'BC MAR842 13/12/2005 End
            
            Set xmlCriticalDataRequestDoc = New FreeThreadedDOMDocument40
            xmlCriticalDataRequestDoc.validateOnParse = False
            xmlCriticalDataRequestDoc.setProperty "NewParser", True
            xmlCriticalDataRequestDoc.async = False
    
            Set xmlCriticalDataRequestNode = _
                xmlCriticalDataRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
            Set xmlElem = xmlCriticalDataRequestDoc.createElement("BEFORE")
            Set xmlNode = xmlCriticalDataRequestNode.appendChild(xmlElem)
    
            'MAR1717 GHun Append ApplicationCreditCheck details if they exist
            Set xmlTemp = xmlApplicationDataDoc.selectSingleNode("RESPONSE/RESPONSE/APPLICATION").cloneNode(True) 'BC MAR842
            If Not xmlAppCreditCheck Is Nothing Then
                xmlTemp.appendChild xmlAppCreditCheck.cloneNode(True)   'MAR1759 GHun
            End If
            
            xmlNode.appendChild xmlTemp
            'MAR1717 End
    
        End If  'MAR727 GHun
        
        
    ' call Data Update Function
        ForwardBORequest vxmlRequestNode, vxmlResponseNode
    
        If blnCreditCheckTestRequired Or blnCostModelTestRequired Or blnHunterTestRequired Then  'MAR727 MAR1126
            
            ' get Application after image
            xmlApplicationDataDoc.removeChild xmlApplicationDataDoc.firstChild
            
'BC MAR842 13/12/2005 Begin
'            Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
'            objNTxBO.GetApplicationData xmlContextNode, xmlApplicationDataDoc, "CDRBTemplate"
'            Set objNTxBO = Nothing
           
            Set xmlRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CRITICALDATACONTEXT")
            Set xmlTempRequest = xmlThisRequestDoc.createElement("REQUEST")
            xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONNUMBER"
            xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONFACTFINDNUMBER"

            xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetApplicationData"
            ' run and then load and check response
            
            Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
            strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            Set xmlApplicationDataDoc = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlApplicationDataDoc.selectSingleNode("RESPONSE"), , True
            Set objNoTxBO = Nothing
'BC MAR842 13/12/2005 End
            
            Set xmlElem = xmlCriticalDataRequestDoc.createElement("AFTER")
            Set xmlNode = xmlCriticalDataRequestNode.appendChild(xmlElem)
            
            xmlNode.appendChild _
                xmlApplicationDataDoc.selectSingleNode("RESPONSE/RESPONSE/APPLICATION").cloneNode(True) 'BC MAR842
        End If  'MAR727 GHun
        'MARS 727 - This is a straight through i.e. BEFORE & AFTER data have already been defined
    Else ' No update required
        Set xmlCriticalDataRequestDoc = New FreeThreadedDOMDocument40
        xmlCriticalDataRequestDoc.validateOnParse = False
        xmlCriticalDataRequestDoc.setProperty "NewParser", True
        xmlCriticalDataRequestDoc.async = False
        Set xmlCriticalDataRequestNode = _
        xmlCriticalDataRequestDoc.appendChild(vxmlRequestNode.cloneNode(True))
        'MAR 727 - temporary work around
        ' do the CREDITCHECK_TEST or COSTMODELLING_TEST attribs exist - if so remove them
        Set xmlElem = xmlCriticalDataRequestDoc.firstChild
        xmlElem.removeAttribute ("CREDITCHECK_TEST")
        xmlElem.removeAttribute ("COSTMODELLING_TEST")
        
        'MAR1759 GHun
        If Not xmlAppCreditCheck Is Nothing Then
            Set xmlTemp = xmlCriticalDataRequestDoc.selectSingleNode("/REQUEST/BEFORE/APPLICATION")
            If Not xmlTemp Is Nothing Then
                xmlTemp.appendChild xmlAppCreditCheck.cloneNode(True)
            End If
        End If
        'MAR1759 End
    End If
    
    'MARS 727 - end
    If Not (blnDataUpdateRequired) Or blnCreditCheckTestRequired Or blnCostModelTestRequired _
                                                                 Or blnHunterTestRequired Then      'MAR727 MAR1126
        'MAR 856 DRC - get the globalparameterlist for the tasks
        Set xmlRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CRITICALDATACONTEXT")
        Set xmlTempRequest = xmlThisRequestDoc.createElement("REQUEST")
        xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONNUMBER"
        xmlCopyAttribute xmlRequestNode, xmlTempRequest, "APPLICATIONFACTFINDNUMBER"
    
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GetApplicationData"
        ' run and then load and check response
        
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        Set xmlApplicationDataDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlApplicationDataDoc.selectSingleNode("RESPONSE"), , True
        Set objNoTxBO = Nothing
        Set xmlGlobalParameterList = xmlApplicationDataDoc.selectNodes(".//GLOBALPARAMETER")
        For Each xmlGPNode In xmlGlobalParameterList
            xmlCriticalDataRequestNode.appendChild xmlGPNode
        Next
        'MAR 856 DRC
               
        If blnCreditCheckTestRequired Then
            Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("CREDITCHECK_TEST")
            xmlAttrib.Text = "YES"
            xmlCriticalDataRequestNode.Attributes.setNamedItem xmlAttrib
        End If
        
        If blnCostModelTestRequired Then
            Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("COSTMODELLING_TEST")
            xmlAttrib.Text = "YES"
            xmlCriticalDataRequestNode.Attributes.setNamedItem xmlAttrib
        End If
                
        'MAR1126
        If blnHunterTestRequired Then
 ' MAR1323      Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("RESENDHUNTERTEST") ' change the name
            Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("HUNTER_TEST")
            xmlAttrib.Text = "YES"
            xmlCriticalDataRequestNode.Attributes.setNamedItem xmlAttrib
        End If
' MAR1323 - do these two tests only when the Omiga data has been updated from CRM
        If Not (blnDataUpdateRequired) Then
            Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("ADDRESSCHANGED_TEST")
            xmlAttrib.Text = "YES"
            xmlCriticalDataRequestNode.Attributes.setNamedItem xmlAttrib
            Set xmlAttrib = xmlCriticalDataRequestDoc.createAttribute("CUSTCATCHANGED_TEST")
            xmlAttrib.Text = "YES"
            xmlCriticalDataRequestNode.Attributes.setNamedItem xmlAttrib
        End If
' MAR1323 - End
        
                
        ' Call the Critical Data Rules
    
        Set xmlCriticalDataRulesResponseDoc = New FreeThreadedDOMDocument40
        xmlCriticalDataRulesResponseDoc.validateOnParse = False
        xmlCriticalDataRulesResponseDoc.setProperty "NewParser", True
        xmlCriticalDataRulesResponseDoc.async = False
            
        RunCriticalDataRules xmlCriticalDataRequestNode, xmlCriticalDataRulesResponseDoc
        
        Set xmlCriticalDataRulesResponseNode = _
            xmlCriticalDataRulesResponseDoc.selectSingleNode("RESPONSE")
    
        errCheckXMLResponseNode xmlCriticalDataRulesResponseNode, , True
        xmlCopyAttribute xmlCriticalDataRulesResponseNode, vxmlResponseNode, "DATACHANGEDFLAG"
        'DRC MARS856 - attach customers changed data to response
        vxmlResponseNode.appendChild xmlCriticalDataRulesResponseNode.selectSingleNode(".//CUSTOMERS")
     
        ' any tasks to be created?
        If Not xmlCriticalDataRulesResponseNode.selectSingleNode("CASETASK") Is Nothing Then
        
            
            CreateCriticalDataTasks _
                vxmlRequestNode, _
                xmlCurrentStageNode, _
                xmlContextNode, _
                xmlCriticalDataRulesResponseNode, _
                blnOutstandingReScoreCreditCheckTask
    
         End If
    
    End If
    
CriticalDataCheckExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'MAR1126 Add check for Hunter Send tasks

Private Sub CheckTasksOnCurrentStage( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByRef vblnHasOutstandingCreditCheckTask As Boolean, _
    ByRef vblnHasOutstandingRescoreCreditCheckTask As Boolean, _
    ByRef vblnHasCompletedCreditCheckTask As Boolean, _
    ByRef vblnOutstandingCostModelTask As Boolean, _
    ByRef vblnOutstandingHunterTask As Boolean, _
    ByRef vblnHasCompletedHunterTask As Boolean)                    ' MAR1126
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    
    Dim strXSL As String, _
        strCreditCheckTask As String, _
        strReCreditCheckTask As String, _
        strRescoreCreditCheckTask As String, _
        strReCostModelTask As String
    Dim strHunterTask As String                                     ' MAR1126
    Dim strResendHunterTask As String                               ' MAR1126
    
    Dim blnOutstandingCreditCheckTask As Boolean, _
        blnCompletedCreditCheckTask As Boolean, _
        blnOutstandingReDoCreditCheckTask As Boolean, _
        blnOutstandingReScoreCreditCheckTask As Boolean, _
        blnCompletedReDoCreditCheckTask As Boolean
        
    Dim blnOutstandingHunterTask As Boolean                         ' MAR1126
    Dim blnCompletedHunterTask As Boolean                           ' MAR1126
    Dim blnOutstandingResendHunterTask As Boolean                   ' MAR1126
    Dim blnCompletedResendHunterTask As Boolean                     ' MAR1126

    Dim intTaskInstance As Integer, _
        intThisInstance As Integer
        
    ' needs global param?
    ' PSC 18/01/2006 MAR1074
    strCreditCheckTask = GetGlobalParamString("TMInitialCreditCheck")
    strReCreditCheckTask = GetGlobalParamString("TMReprocessCreditCheck")
    strRescoreCreditCheckTask = GetGlobalParamString("TMRescoreCreditCheck")
    strReCostModelTask = GetGlobalParamString("TMRemodelMortgage")
    strHunterTask = GetGlobalParamString("TMHunter")                    ' MAR1126
    strResendHunterTask = GetGlobalParamString("TMResendHunterTaskId")  ' MAR1126
    
    strXSL = "CASETASK[@TASKID='" & strCreditCheckTask & "']"
    
    Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
    
    intTaskInstance = -1
    
    For Each xmlNode In xmlNodeList
        intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
        If intThisInstance > intTaskInstance Then
            intTaskInstance = intThisInstance
            If ((xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete) Or _
                (xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSNotApplicable)) Then
                blnOutstandingCreditCheckTask = False
                blnCompletedCreditCheckTask = True
            Else
                blnOutstandingCreditCheckTask = True
                blnCompletedCreditCheckTask = False
            End If
        End If
    Next
   
    If Len(strReCreditCheckTask) > 0 Then
    
        strXSL = "CASETASK[@TASKID='" & strReCreditCheckTask & "']"
        
        Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
        
        intTaskInstance = -1
        
        For Each xmlNode In xmlNodeList
            intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
            If intThisInstance > intTaskInstance Then
                intTaskInstance = intThisInstance
                If ((xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete) Or _
                    (xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSNotApplicable)) Then
                    
                    blnOutstandingReDoCreditCheckTask = False
                    blnCompletedReDoCreditCheckTask = True
                Else
                    blnOutstandingReDoCreditCheckTask = True
                    blnCompletedReDoCreditCheckTask = False
                End If
            End If
        Next
        
    End If
    
    If blnOutstandingCreditCheckTask Or _
       blnOutstandingReDoCreditCheckTask Then
        vblnHasOutstandingCreditCheckTask = True
    Else
        If blnCompletedCreditCheckTask Or _
           blnCompletedReDoCreditCheckTask Then
            vblnHasCompletedCreditCheckTask = True
        End If
    End If
    'MAR65 - DRC 05/10/2005 check for outstanding rescore credit task
    If Len(strRescoreCreditCheckTask) > 0 Then
    
        strXSL = "CASETASK[@TASKID='" & strRescoreCreditCheckTask & "']"
        
        Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
        
        intTaskInstance = -1
        
        For Each xmlNode In xmlNodeList
            intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
            If intThisInstance > intTaskInstance Then
                intTaskInstance = intThisInstance
                If ((xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete) Or _
                    (xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSNotApplicable)) Then
                    vblnHasOutstandingRescoreCreditCheckTask = False
                Else
                    vblnHasOutstandingRescoreCreditCheckTask = True
                End If
            End If
        Next
        
    End If
    'MAR65 - End
    
    'MAR1126 Check for outstanding Hunter tasks
  
    If Len(strHunterTask) > 0 Then

        strXSL = "CASETASK[@TASKID='" & strHunterTask & "']"
    
        Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
    
        intTaskInstance = -1
    
        For Each xmlNode In xmlNodeList
            intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
            If intThisInstance > intTaskInstance Then
                intTaskInstance = intThisInstance
                If ((xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete) Or _
                    (xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSNotApplicable)) Then
                    
                    blnOutstandingHunterTask = False
                    blnCompletedHunterTask = True
                Else
                    blnOutstandingHunterTask = True
                    blnCompletedHunterTask = False
                End If
            End If
        Next
    End If
   
    If Len(strResendHunterTask) > 0 Then
    
        strXSL = "CASETASK[@TASKID='" & strResendHunterTask & "']"
            
        Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
        
        intTaskInstance = -1
        
        For Each xmlNode In xmlNodeList
            intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
            If intThisInstance > intTaskInstance Then
                intTaskInstance = intThisInstance
                If ((xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete) Or _
                    (xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSNotApplicable)) Then
                    blnOutstandingResendHunterTask = False
                    blnCompletedResendHunterTask = True
                Else
                    blnOutstandingResendHunterTask = True
                    blnCompletedResendHunterTask = False
                End If
            End If
        Next
        
    End If
    
    If blnOutstandingHunterTask Or _
       blnOutstandingResendHunterTask Then
        vblnOutstandingHunterTask = True
    Else
        If blnCompletedHunterTask Or _
           blnCompletedResendHunterTask Then
            vblnHasCompletedHunterTask = True
        End If
    End If
    
    'MAR1126 - End
   
    strXSL = "CASETASK[@TASKID='" & strReCostModelTask & "']"
    
    Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
    
    intTaskInstance = -1
    
    For Each xmlNode In xmlNodeList
        intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
        If intThisInstance > intTaskInstance Then
            intTaskInstance = intThisInstance
            If xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete Then
                vblnOutstandingCostModelTask = False
            Else
                vblnOutstandingCostModelTask = True
            End If
        End If
    Next

End Sub

'MAR1717 GHun Not used
'Private Function IsAfterCreditCheckStage(ByVal vstrThisStageId As String) As Boolean
'
'    Const cstrFunctionName As String = "IsAfterCreditCheckStage"
'
'    Dim strCreditCheckStage As String
'
'    strCreditCheckStage = GetMandatoryGlobalParamString("TMCreditCheckStage")
'
'    If Not IsNumeric(vstrThisStageId) Then
'        errThrowError _
'            cstrFunctionName, _
'            oeInvalidParameter, _
'            "this Stage Id not numeric: " & vstrThisStageId
'    End If
'
'    If Not IsNumeric(strCreditCheckStage) Then
'        errThrowError _
'            cstrFunctionName, _
'            oeInvalidParameter, _
'            "CreditCheck Stage Id not numeric: " & strCreditCheckStage
'    End If
'
'    IsAfterCreditCheckStage = CInt(vstrThisStageId) > CInt(strCreditCheckStage)
'
'End Function
'MAR1717 End

'MAR1126  Add function
Private Function IsAfterHunterStage(ByVal vstrThisStageId As String) As Boolean

    Const cstrFunctionName As String = "IsAfterHunterStage"

    Dim strHunterStage As String
    
    strHunterStage = GetMandatoryGlobalParamString("TMHunterStage")
    
    If Not IsNumeric(vstrThisStageId) Then
        errThrowError _
            cstrFunctionName, _
            oeInvalidParameter, _
            "this Stage Id not numeric: " & vstrThisStageId
    End If
    
    If Not IsNumeric(strHunterStage) Then
        errThrowError _
            cstrFunctionName, _
            oeInvalidParameter, _
            "Hunter Stage Id not numeric: " & strHunterStage
    End If
    
    IsAfterHunterStage = CInt(vstrThisStageId) > CInt(strHunterStage)

End Function

'MAR1717 GHun Never used
'Private Function CostModelCheckRequired( _
'    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
'    ByVal vxmlApplicationNode As IXMLDOMNode) _
'    As Boolean
'
'    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
'    Dim xmlAttrib As IXMLDOMAttribute
'    Dim objNTxBO As IOmTmNTxBO  'SYS3270
'    ' add _EXTRACTTYPE_ 'NODE' to application node
'    ' this gets copied to omRB request by GetApplicationData,
'    ' this causes omRB to return APPLICATION node only (no child nodes)
'    Set xmlAttrib = vxmlApplicationNode.ownerDocument.createAttribute("_EXTRACTTYPE_")
'    xmlAttrib.Text = "NODE"
'    vxmlApplicationNode.Attributes.setNamedItem xmlAttrib
'
'    Set xmlApplicationDataDoc = New FreeThreadedDOMDocument40
'    xmlApplicationDataDoc.validateOnParse = False
'    xmlApplicationDataDoc.setProperty "NewParser", True
'    xmlApplicationDataDoc.async = False
'    'SYS3270 {
'    'GetApplicationData vxmlApplicationNode, xmlApplicationDataDoc
'    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
'    objNTxBO.GetApplicationData vxmlApplicationNode, xmlApplicationDataDoc
'    Set objNTxBO = Nothing
'    'SYS3270 }
'
'
'    vxmlApplicationNode.Attributes.removeNamedItem "_EXTRACTTYPE_"
'
'    CostModelCheckRequired = _
'        Len( _
'            xmlGetNodeText(xmlApplicationDataDoc, "RESPONSE/APPLICATION/@ACCEPTEDQUOTENUMBER")) _
'            > 0
'
'End Function
'MAR1717 End

Private Sub ForwardBORequest(ByVal vxmlRequestNode As IXMLDOMNode, _
                             ByVal vxmlResponseNode As IXMLDOMNode)

    Const cstrFunctionName As String = "ForwardBORequest"

    On Error GoTo ForwardBORequestExit:
    
    Dim objNTxBO As IOmTmNTxBO
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    
    objNTxBO.ForwardBORequest vxmlRequestNode, vxmlResponseNode
    
ForwardBORequestExit:

    Set objNTxBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub RunCriticalDataRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCriticalDataRulesResponseDoc As FreeThreadedDOMDocument40)
    
    Dim objCDRulesBO As Object
    
    Set objCDRulesBO = gobjContext.CreateInstance("OmCDRules.OmCDRulesBO")
    
    vxmlCriticalDataRulesResponseDoc.loadXML _
        objCDRulesBO.RunRules(vxmlRequestNode.xml)
        
    Set objCDRulesBO = Nothing
        
    errCheckXMLResponseNode _
        vxmlCriticalDataRulesResponseDoc.selectSingleNode("RESPONSE"), , True

End Sub

Private Sub RunFTInterfaceCall(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName          As String = "RunFTInterfaceCall"
    Dim objFirstTitleBO             As Object
    Dim vxmlFirstTitleResponseDoc   As FreeThreadedDOMDocument40
    
On Error GoTo RunFTInterfaceCallVbErr

    Set vxmlFirstTitleResponseDoc = New FreeThreadedDOMDocument40
    
    Set objFirstTitleBO = gobjContext.CreateInstance("OmFirstTitle.FirstTitleOutboundBO")
    
    vxmlFirstTitleResponseDoc.loadXML objFirstTitleBO.Execute(vxmlRequestNode.xml)
        
    errCheckXMLResponseNode vxmlFirstTitleResponseDoc.documentElement, vxmlResponseNode, True
      
RunFTInterfaceCallVbErr:
    Set objFirstTitleBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
End Sub

Private Sub RunEsurvValuation(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    Dim objESurvBO As Object
    Dim vxmlESurvResponseDoc As FreeThreadedDOMDocument40
    Set vxmlESurvResponseDoc = New FreeThreadedDOMDocument40
    
    Set objESurvBO = gobjContext.CreateInstance("OmESurv.OutBoundBO")
    
    vxmlESurvResponseDoc.loadXML objESurvBO.RunEsurvValuation(vxmlRequestNode.xml)
        
    errCheckXMLResponseNode vxmlESurvResponseDoc.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    Set objESurvBO = Nothing
        
End Sub

'MAR750  Updated
Private Sub RunDelayedCompletionInterface(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Dim xmlDoc               As FreeThreadedDOMDocument40
    Dim xmlApplication       As IXMLDOMNode
    Dim xmlApp               As IXMLDOMNode
    Dim xmlCaseTask          As IXMLDOMNode
    Dim xmlTaskNode          As IXMLDOMNode
    Dim xmlTaskElem          As IXMLDOMElement
    Dim xmlTempRequest       As IXMLDOMElement
    Dim xmlNewRequest        As IXMLDOMElement
    Dim xmlElem              As IXMLDOMElement
    Dim objDelayCompletionBO As Object
    Dim objAppBO             As Object

    Dim strResponse          As String
    Dim strTaskDueDate       As String
    Dim strApplNo            As String
    Dim strAFFNo             As String
    Dim strOtherSysAccountNumber As String

On Error GoTo RunDelayedCompletionInterfaceExit

    Const cstrFunctionName As String = "RunDelayedCompletionInterface"
    
    'Initialise
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    
    'Build Request and Get Application data
    strTaskDueDate = xmlGetAttributeAsDate(xmlCaseTask, "TASKDUEDATEANDTIME")
    
    strApplNo = xmlGetAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAFFNo = xmlGetAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApp = xmlDoc.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlApp
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strApplNo
    xmlApp.appendChild xmlElem
    
    strResponse = objAppBO.GetApplicationData(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    strOtherSysAccountNumber = xmlDoc.selectSingleNode("//APPLICATION/OTHERSYSTEMACCOUNTNUMBER").Text
 
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlElem = xmlApplication.cloneNode(False)
    xmlNewRequest.appendChild xmlElem
    xmlSetAttributeValue xmlElem, "OTHERSYSTEMACCOUNTNUMBER", strOtherSysAccountNumber
 
    Set xmlElem = xmlDoc.createElement("CASETASK")
    xmlNewRequest.appendChild xmlElem
    xmlSetAttributeValue xmlElem, "TASKDUEDATE", strTaskDueDate
    
    'Call the Interface
    Set objDelayCompletionBO = CreateObject("omDC.DelayCompletionBO")
    
    strResponse = objDelayCompletionBO.RunDelayedCompletionInterface(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    'Update Case Task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)
    
    strResponse = gobjTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
            
RunDelayedCompletionInterfaceExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
       
    Set xmlDoc = Nothing
    Set xmlApplication = Nothing
    Set xmlApp = Nothing
    Set xmlTaskNode = Nothing
    Set xmlCaseTask = Nothing
    Set xmlTaskElem = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNewRequest = Nothing
    Set objDelayCompletionBO = Nothing
    
    gobjContext.SetComplete
    
    Exit Sub
    
End Sub

Private Sub RunHomeTrackValuation(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlRespDoc          As FreeThreadedDOMDocument40
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCustomerNodeList As IXMLDOMNodeList 'JD MAR503
    Dim xmlNewCustListNode  As IXMLDOMNode  'JD MAR503
    Dim xmlCustNode         As IXMLDOMNode  'JD MAR503
    Dim xmlCaseTask         As IXMLDOMNode
    Dim xmlTaskNode         As IXMLDOMNode
    Dim xmlNodeList         As IXMLDOMNodeList
    Dim xmlTaskElem         As IXMLDOMElement
    Dim xmlTaskNote         As IXMLDOMElement
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlCaseActivity     As IXMLDOMElement
    Dim xmlApplNodeAttr     As IXMLDOMNode
    
    Dim objNoTxBO           As omTmNoTxBO
    Dim objHomeTrackBO      As Object
    Dim objMsgTmBO          As Object
    Dim objAppBO            As Object

    Dim strApplPriority     As String
    Dim strAppNum           As String ' MAR1250
    Dim strAppFFNum         As String ' MAR1250
    Dim strResponse         As String
    Dim strValStatus        As String
    Dim strValType          As String
    Dim strTaskStatus       As String
    Dim strTaskId           As String
    Dim strTaskNotes        As String
    Dim strMemoNoteID       As String
    Dim i                   As Integer
    
    'MAR630
    Dim strStageId          As String
    Dim strCaseActivityGuid As String
    Dim strCaseStageSeqNo   As String
    Dim iTaskInstance       As Integer
    
    Dim xmlAppCopy As IXMLDOMNode       ' PSC 13/02/2006 MAR1250
    
    
On Error GoTo RunHomeTrackValuationExit

    Const cstrFunctionName As String = "RunHomeTrackValuation"
    'Step 1:
    'Initialise
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'Validate
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    'The customer list may already exist otherwise get it from the database
    Set xmlCustomerListNode = xmlGetNode(vxmlRequestNode, "CUSTOMERLIST")
    If xmlCustomerListNode Is Nothing Then
        'check for some customer nodes. If we have some create a CUSTOMERLIST node   JD MAR503
        Set xmlCustomerNodeList = vxmlRequestNode.selectNodes("APPLICATION/CUSTOMER")
        If xmlCustomerNodeList.length = 0 Then
            Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
            xmlTempRequest.setAttribute "OPERATION", "FINDCUSTOMERSFORCREDITCHECK"
            
            ' PSC 13/02/2006 MAR1250 - Start
            Set xmlAppCopy = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
            xmlTempRequest.appendChild xmlAppCopy
            
            ' run and then load and check response
            Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
            strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
            xmlTempDoc.loadXML strResponse
            errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
            Set xmlCustomerListNode = xmlGetMandatoryNode(xmlTempDoc.documentElement, "CUSTOMERLIST")
            
            'RF 11/02/2006  MAR1250 Start - need list to be attribute based
            Set xmlCustomerListNode = xmlCreateAttributeBasedResponse(xmlCustomerListNode, True)
            'RF 11/02/2006  MAR1250 End
            
        Else
            Set xmlNewCustListNode = xmlTempDoc.createElement("CUSTOMERLIST")
            vxmlRequestNode.appendChild xmlNewCustListNode
            For Each xmlCustNode In xmlCustomerNodeList
                xmlNewCustListNode.appendChild xmlCustNode
            Next
            Set xmlCustomerListNode = xmlGetNode(vxmlRequestNode, "CUSTOMERLIST")
        End If
    End If
    
    'read CASETASK dtl
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    ' PSC 13/02/2006 MAR1250
    strApplPriority = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONPRIORITY")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    'MAR263 <APPLICATION APPLICATIONNUMBER=... APPLICATIONFACTFINDNUMBER=.../> required
    Set xmlApplNodeAttr = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' RF 11/02/2006 MAR1250 Start - get values here as it seems later on xml can be either
    ' attribute based or element based depending on calling routes
    ' PSC 13/02/2006 MAR1250 - Start
    strAppNum = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strAppFFNum = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    ' PSC 13/02/2006 MAR1250 - End
    ' RF 11/02/2006 MAR1250 End

    xmlTempRequest.appendChild xmlApplNodeAttr
    xmlTempRequest.appendChild xmlCustomerListNode
    
    'Step 2: a.call the dot net component
    Set objHomeTrackBO = gobjContext.CreateInstance("omHomeTrack.HomeTrackBO")
    ' run and then load and check response
    strResponse = objHomeTrackBO.RunHomeTrackValuation(xmlTempRequest.xml)
    Set objHomeTrackBO = Nothing
 
    'debug purposes
    'strResponse = "<RESPONSE TYPE=" & Chr$(34) & "SUCCESS" & Chr$(34) & "><VALUATION STATUS=" & Chr$(34) & "NOTSATISFIED" & Chr$(34) & " VALUATIONTYPE=" & Chr$(34) & "4" & Chr$(34) & "/><TASKS><TASK TASKID=" & Chr$(34) & "HomeT_Response" & Chr$(34) & " TASKNOTES=" & Chr$(34) & "Is ERRORID blank" & Chr$(34) & "/><TASK TASKID=" & Chr$(34) & "HomeT_Response" & Chr$(34) & " TASKNOTES=" & Chr$(34) & "Is ERRORID blank xxx" & Chr$(34) & "/></TASKS></RESPONSE>"
    Set xmlRespDoc = New FreeThreadedDOMDocument40
    xmlRespDoc.loadXML strResponse
    errCheckXMLResponseNode xmlRespDoc.selectSingleNode("RESPONSE"), , True
    
    ' check status returned
    strValStatus = xmlGetAttributeText(xmlRespDoc.selectSingleNode("//VALUATION"), "STATUS")
    strValType = xmlGetAttributeText(xmlRespDoc.selectSingleNode("//VALUATION"), "VALUATIONTYPE")
    
    'update CaseTask Status depeding on status returned
    Select Case UCase$(strValStatus)
    Case "SATISFIED"
        'set task to completed "40"
        strTaskStatus = "40"
    Case "NOTSATISFIED"
        'set task to not aplicable
        strTaskStatus = "30"
    End Select
    
    'b. Update Case Task
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
    Set xmlTaskElem = xmlCaseTask.cloneNode(True)
    xmlSetAttributeValue xmlTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlTaskElem, "TASKSTATUS", strTaskStatus
    xmlTempRequest.appendChild xmlTaskElem
    
    ' run and then load and check response
    Set objMsgTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    
    'c. if not applicable, update NewProperty Valuation Type from .net
    If (strTaskStatus = "30") Then
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        Set xmlTaskNode = xmlTempDoc.createElement("NEWPROPERTY")
        
        Set xmlTaskElem = xmlTempDoc.createElement("APPLICATIONNUMBER")
        xmlTaskElem.Text = strAppNum ' RF 11/02/2006 MAR1250
        xmlTaskNode.appendChild xmlTaskElem
        
        Set xmlTaskElem = xmlTempDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlTaskElem.Text = strAppFFNum ' RF 11/02/2006 MAR1250
        xmlTaskNode.appendChild xmlTaskElem
        
        Set xmlTaskElem = xmlTempDoc.createElement("VALUATIONTYPE")
        xmlTaskElem.Text = strValType
        xmlTaskNode.appendChild xmlTaskElem
        
        xmlTempRequest.appendChild xmlTaskNode

        Set objAppBO = gobjContext.CreateInstance("omApp.NewPropertyBO")
        strResponse = objAppBO.UpdateNewProperty(xmlTempRequest.xml)
        errCheckXMLResponse strResponse, True
    End If
    
    'Step 3: create list of task returned from .net
    Set xmlNodeList = xmlRespDoc.selectNodes("//TASK")
    strMemoNoteID = GetFirstComboValueId("TaskNoteType", "M")
    
    If xmlNodeList.length > 0 Then
        
        For i = 0 To xmlNodeList.length - 1 Step 1
            Set xmlTaskNode = xmlNodeList.Item(i)
            strTaskId = xmlGetAttributeText(xmlTaskNode, "TASKID")
            strTaskNotes = xmlGetAttributeText(xmlTaskNode, "TASKNOTES")
            
            Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlTempRequest.setAttribute "OPERATION", "CREATEADHOCCASETASK"
            
            Set xmlTaskElem = xmlTempDoc.createElement("APPLICATION")
            xmlSetAttributeValue xmlTaskElem, "APPLICATIONPRIORITY", strApplPriority
            xmlTempRequest.appendChild xmlTaskElem
            
            Set xmlTaskElem = xmlTempDoc.createElement("CASETASK")
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "SOURCEAPPLICATION"
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "CASEID"
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "ACTIVITYID"
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "ACTIVITYINSTANCE"
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "STAGEID"
            xmlCopyAttribute xmlCaseTask, xmlTaskElem, "CASESTAGESEQUENCENO"
            
            xmlSetAttributeValue xmlTaskElem, "TASKID", strTaskId
            xmlSetAttributeValue xmlTaskElem, "TASKSTATUS", "10"
            'xmlSetAttributeValue xmlTaskElem, "MANDATORYINDICATOR", "1"
            xmlSetSysDateToNodeAttrib xmlTaskElem, "TASKDUEDATEANDTIME"
    
            xmlTempRequest.appendChild xmlTaskElem
            
            'create adhoc task
            CreateAdhocCaseTask xmlTempRequest
            
            'If TaskNote was passed in, create TaskNote record
            If Len(strTaskNotes) > 0 Then
                'get UserID, UnitID
                Dim strUserId As String, strUnitId As String
                Call GetApplicationOwners(gobjContext, vxmlRequestNode, strUserId, strUnitId)
                
                'MAR630 Get Task Instance
                strStageId = xmlGetAttributeText(xmlCaseTask, "STAGEID")
                strCaseActivityGuid = xmlGetAttributeText(xmlCaseTask, "CASEACTIVITYGUID")
                strCaseStageSeqNo = xmlGetAttributeText(xmlCaseTask, "CASESTAGESEQUENCENO")
                
                iTaskInstance = GetMaxCaseTaskInstance(strStageId, strTaskId, strCaseActivityGuid, strCaseStageSeqNo)
                
                Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlTempRequest, "OPERATION", "CREATETASKNOTE"

                Set xmlTaskNote = xmlTempDoc.createElement("TASKNOTE")
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "CASEID"
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "ACTIVITYID"
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "ACTIVITYINSTANCE"
                xmlSetAttributeValue xmlTaskNote, "TASKINSTANCE", CStr(iTaskInstance) ' MAR630
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "STAGEID"
                xmlCopyAttribute xmlCaseTask, xmlTaskNote, "CASESTAGESEQUENCENO"
                
                xmlSetAttributeValue xmlTaskNote, "TASKID", strTaskId
                xmlSetAttributeValue xmlTaskNote, "NOTEENTRY", strTaskNotes
                xmlSetAttributeValue xmlTaskNote, "NOTEORIGINATINGUSERID", strUnitId
                xmlSetAttributeValue xmlTaskNote, "NOTETYPE", strMemoNoteID
                
                xmlTempRequest.appendChild xmlTaskNote
                ' run and then load and check response
                strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
                xmlTempDoc.loadXML (strResponse)
                errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
            End If
        Next i
    End If
    gobjContext.SetComplete
    
RunHomeTrackValuationExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Set xmlTempDoc = Nothing
    Set xmlRespDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlApplNodeAttr = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlTaskNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlCaseTask = Nothing
    Set xmlTaskElem = Nothing
    Set xmlTaskNote = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseActivity = Nothing
    
    Set objNoTxBO = Nothing
    Set objHomeTrackBO = Nothing
    Set objMsgTmBO = Nothing
    Set objAppBO = Nothing
    Set xmlAppCopy = Nothing ' PSC 13/02/2006 MAR1250
    
    Exit Sub
    
End Sub
Private Sub CreateCriticalDataTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCriticalDataContextNode As IXMLDOMNode, _
    ByVal vxmlCriticalDataRulesResponseNode As IXMLDOMNode, _
    ByVal vblnOutstandingRescoreCreditCheckTask As Boolean)

    On Error GoTo CreateCriticalDataTasksExit

    Const cstrFunctionName As String = "CreateCriticalDataTasks"
    
    ' create CreateAdhocCaseTask request
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTestNode As IXMLDOMNode
    Dim strReCreditCheckTask As String
    Dim strRescoreCreditCheckTask As String
    Dim objBO As Object
    Dim strResponse As String
    
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateAdhocCaseTask"
    
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    xmlCopyAttribute vxmlCriticalDataContextNode, xmlElem, "APPLICATIONPRIORITY"
    xmlThisRequestNode.appendChild xmlElem
    strReCreditCheckTask = GetGlobalParamString("TMReprocessCreditCheck")
    strRescoreCreditCheckTask = GetGlobalParamString("TMRescoreCreditCheck")
  
    For Each xmlNode In vxmlCriticalDataRulesResponseNode.childNodes
    

        'Don't create a new Rescore Task if there already is one on the current stage
        If (Not (vblnOutstandingRescoreCreditCheckTask) Or _
            (xmlGetAttributeText(xmlNode, "TASKID") <> strRescoreCreditCheckTask)) Then
            Set xmlThisCaseTaskNode = xmlThisRequestNode.appendChild(xmlNode.cloneNode(False))
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "SOURCEAPPLICATION"
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "CASEID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "CASESTAGESEQUENCENO"
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "ACTIVITYID"
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "ACTIVITYINSTANCE"
            xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "STAGEID"
        
        End If
        '
        'A ReProcess Credit check has precedence over a Rescore so must check for this
        ' and set to 'Non Applicable' any existing rescore task  when we want a new reprocess task
        '
        If vblnOutstandingRescoreCreditCheckTask And _
           (xmlGetAttributeText(xmlNode, "TASKID") = strReCreditCheckTask) Then
           For Each xmlTestNode In vxmlCurrentStageNode.childNodes
              If ((xmlGetAttributeText(xmlTestNode, "TASKID") = strRescoreCreditCheckTask) And _
                (xmlGetAttributeAsInteger(xmlTestNode, "TASKSTATUS") < 30)) Then
                
                ' Now update Case Task
                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
                Set xmlCaseTaskElem = xmlTestNode.cloneNode(True)
                xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
                xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
                xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
                xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", 30
                xmlNewRequest.appendChild xmlCaseTaskElem
                
                Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
                strResponse = objBO.TmRequest(xmlNewRequest.xml)
                
                errCheckXMLResponse strResponse, True
                            
              End If
           Next
            
        End If
         
     Next
     'DRC - need to check that we've got at least one new task
     If Not (xmlThisRequestNode.selectSingleNode(".//CASETASK") Is Nothing) Then
        CreateAdhocCaseTask xmlThisRequestNode
     End If
     
    
    
CreateCriticalDataTasksExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub





Private Sub CreateCurrEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreateCurrEmployersRefErr

    Const cstrFunctionName As String = "CreateCurrEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTEMPLOYERSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreateCurrEmployersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreateCurrEmployersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateCurrEmployersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateCurrEmployersRefErr:

    gobjContext.SetAbort
    
    GoTo CreateCurrEmployersRefExit
    
End Sub

Private Sub ValidateCurrEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidateCurrEmployersRefErr

    Const cstrFunctionName As String = "ValidateCurrEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTEMPLOYERSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    
    ' Call referencesBO.ValidateCurrEmployersRef
    ' NOT WRITTEN YET
    ' SA SYS2256 Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidateCurrEmployersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256
    
    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidateCurrEmployersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidateCurrEmployersRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidateCurrEmployersRefExit
    
End Sub

Private Sub CreatePrevEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreatePrevEmployersRefErr

    Const cstrFunctionName As String = "CreatePrevEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSEMPLOYERSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreatePrevEmployersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreatePrevEmployersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreatePrevEmployersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
    Exit Sub
    
CreatePrevEmployersRefErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevEmployersRefExit
    
End Sub

Private Sub ValidatePrevEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidatePrevEmployersRefErr

    Const cstrFunctionName As String = "ValidatePrevEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "VALPREVEMP"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256

    ' Call referencesBO.ValidatePrevEmployersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidatePrevEmployersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256
    
    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidatePrevEmployersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
    Exit Sub
    
ValidatePrevEmployersRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidatePrevEmployersRefExit
    
End Sub

Private Sub CreateAccountantsRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreateAccountantsRefErr

    Const cstrFunctionName As String = "CreateAccountantsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "ACCOUNTANTREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreateAccountantsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreateAccountantsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateAccountantsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateAccountantsRefErr:

    gobjContext.SetAbort
    
    GoTo CreateAccountantsRefExit

    
End Sub

Private Sub ValidateAccountantsRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidateAccountantsRefErr

    Const cstrFunctionName As String = "ValidateAccountantsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "VALACCTREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    
    ' Call referencesBO.ValidateAccountantsRef
    ' NOT YET WRITTEN
    ' SA SYS2256 Code now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidateAccountantsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256
    
    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidateAccountantsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidateAccountantsRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidateAccountantsRefExit

    
End Sub

Private Sub CreateCurrLendersRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreateCurrLendersRefErr

    Const cstrFunctionName As String = "CreateCurrLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTLENDERSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreateCurrLendersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreateCurrLendersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateCurrLendersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateCurrLendersRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo CreateCurrLendersRefExit

    
End Sub

Private Sub ValidateCurrLendersRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidateCurrLendersRefErr

    Const cstrFunctionName As String = "ValidateCurrLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "VALLENDREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256

    ' Call referencesBO.ValidateCurrLendersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidateCurrLendersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidateCurrLendersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidateCurrLendersRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidateCurrLendersRefExit

    
End Sub

Private Sub CreatePrevLendersRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreatePrevLendersRefErr

    Const cstrFunctionName As String = "CreatePrevLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSLENDERSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreatePrevLendersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreatePrevLendersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreatePrevLendersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreatePrevLendersRefErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevLendersRefExit

    
End Sub

Private Sub ValidatePrevLendersRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidatePrevLendersRefErr

    Const cstrFunctionName As String = "ValidatePrevLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "VALPREVLENDREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256

    ' Call referencesBO.ValidatePrevLendersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidatePrevLendersRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidatePrevLendersRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidatePrevLendersRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidatePrevLendersRefExit

    
End Sub

Private Sub CreateCurrLandlordsRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreateCurrLandlordsRefErr

    Const cstrFunctionName As String = "CreateCurrLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTLANDLORDSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreateCurrLandlordsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreateCurrLandlordsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateCurrLandlordsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreateCurrLandlordsRefErr:

    gobjContext.SetAbort
    
    GoTo CreateCurrLandlordsRefExit

    
End Sub

Private Sub ValidateCurrLandlordsRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidateCurrLandlordsRefErr

    Const cstrFunctionName As String = "ValidateCurrLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "VALIDATELANDLORDREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256

    ' Call referencesBO.ValidateCurrLandlordsRef
    ' NOT YET WRITTEN
    'SA SYS2256 Code now availble
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidateCurrLandlordsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 16/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidateCurrLandlordsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidateCurrLandlordsRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidateCurrLandlordsRefExit

    
End Sub

Private Sub CreatePrevLandlordsRef(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreatePrevLandlordsRefErr

    Const cstrFunctionName As String = "CreatePrevLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSLANDLORDSREF"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreatePrevLandlordsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreatePrevLandlordsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreatePrevLandlordsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
CreatePrevLandlordsRefErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevLandlordsRefExit

    
End Sub

Private Sub ValidatePrevLandlordsRef(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ValidatePrevLandlordsRefErr

    Const cstrFunctionName As String = "ValidatePrevLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    'xmlGetMandatoryNode vxmlRequestNode, "VALIDATELANDLORDREF"  '-- SA SYS2256
    xmlGetMandatoryNode vxmlRequestNode, "VALIDATEPREVLANDREF"  '++ SA SYS2256
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    
    ' Call referencesBO.ValidatePrevLandlordsRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.loadXML objBO.OmAppProcRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "ValidatePrevLandlordsRef response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'SA SYS2256 09/05/01 {
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'SA SYS2256

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

ValidatePrevLandlordsRefExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidatePrevLandlordsRefErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidatePrevLandlordsRefExit

    
End Sub

'MAR7 GHun return response node
Private Sub PrintDocumentForTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo PrintDocumentForTaskErr

    Const cstrFunctionName As String = "PrintDocumentForTask"
        
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlRequestPrintAttributesNode As IXMLDOMNode
    Dim xmlRequestApplicationNode As IXMLDOMNode
    Dim objPrintBO As Object
    Dim strValue As String
    Dim strResponseFromPrint As String
 
    'MAR7 GHun
    Dim xmlPrintResponseDoc  As FreeThreadedDOMDocument40
    Set xmlPrintResponseDoc = New FreeThreadedDOMDocument40
    
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    'MAR7 End
 
    Set xmlRequestPrintAttributesNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTATTRIBUTES")
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlRequestApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "COPIES"
    'RF 16/02/2006 MAR1251 Start - Allow template destination of DMS only
    Dim strDestinationType As String
    strDestinationType = xmlGetMandatoryAttributeText(xmlRequestPrintAttributesNode, "DESTINATIONTYPE")
    If Not (strDestinationType = "D" Or strDestinationType = "EF") Then ' D = DMS Only 'EP2_1903 GHun also exclude EF
        xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "PRINTER"
    End If
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "DPSDOCUMENTID"
    'xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "DESTINATIONTYPE"
    'RF 16/02/2006 MAR1251 End
    
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "OUTPUTDOCUMENT"

    'Call the print engine
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "PrintDocument"
    
    'BG 10/10/02 BMIDS00612 add new indicator to ensure archiving happens for these prints.
    xmlSetAttributeValue xmlThisRequestNode, "PRINTINDICATOR", "1"
    'BG 10/10/02 BMIDS00612 END
    
    ' create PRINTDATA element
    Set xmlElem = xmlThisRequestDoc.createElement("PRINTDATA")
    
    ' copy required search attributes
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CASEID")
    xmlElem.setAttribute "APPLICATIONNUMBER", strValue
    
    'AQR SYS2266 DRC - ADDED APPLICATIONFACTNUMBER
    strValue = xmlGetAttributeText(xmlRequestApplicationNode, "APPLICATIONFACTFINDNUMBER")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strValue
    
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CUSTOMERIDENTIFIER")
    xmlElem.setAttribute "CUSTOMERNUMBER", strValue
    
    strValue = xmlGetAttributeText(xmlRequestApplicationNode, "CUSTOMERVERSIONNUMBER")
    xmlElem.setAttribute "CUSTOMERVERSIONNUMBER", strValue
    
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CONTEXT")
    xmlElem.setAttribute "RECIPIENTKEY", strValue
    
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKID")
    xmlElem.setAttribute "REFTYPE", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "METHODNAME")
    xmlElem.setAttribute "METHODNAME", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "RECIPIENTTYPE")
    xmlElem.setAttribute "RECIPIENTTYPE", strValue
    
    xmlThisRequestNode.appendChild xmlElem
    
    ' PSC 04/07/2005 MAR5 - Start
    ' Create TEMPLATEDATA node if required
    strValue = xmlGetAttributeText(xmlRequestApplicationNode, "PRINTKFI")
    If strValue = "1" Then
        Set xmlElem = xmlThisRequestDoc.createElement("TEMPLATEDATA")
        xmlCopyAttribute xmlRequestApplicationNode, xmlElem, "QUOTATIONNUMBER"
        xmlCopyAttribute xmlRequestApplicationNode, xmlElem, "MORTGAGESUBQUOTENUMBER"
        xmlThisRequestNode.appendChild xmlElem
    End If
    ' PSC 04/07/2005 MAR5 - End
        
    ' create CONTROLDATA element
    Set xmlElem = xmlThisRequestDoc.createElement("CONTROLDATA")
    ' clone search attributes
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "OUTPUTDOCUMENT")
    xmlElem.setAttribute "DOCUMENTID", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "PRINTER")
    xmlElem.setAttribute "PRINTER", strValue
        
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "COPIES")
    xmlElem.setAttribute "COPIES", strValue
    
    ' PSC 04/07/2005 MAR5 - Start
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DELIVERYTYPE")
    xmlElem.setAttribute "DELIVERYTYPE", strValue
    ' PSC 04/07/2005 MAR5 - End
    
    'MAR7 GHun
    'RF 16/02/2006 MAR1251 Start - Allow template destination of DMS only
    'Dim strDestinationType As String
    'strDestinationType = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DESTINATIONTYPE")
    'RF 16/02/2006 MAR1251 End
    xmlElem.setAttribute "DESTINATIONTYPE", strDestinationType
    'MAR7 End
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DPSDOCUMENTID")
    xmlElem.setAttribute "DPSDOCUMENTID", strValue
    
    'AQR SYS4467 - DR Need to add more print attributes for DMS
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "HOSTTEMPLATENAME")
    xmlElem.setAttribute "HOSTTEMPLATENAME", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "HOSTTEMPLATEDESCRIPTION")
    xmlElem.setAttribute "HOSTTEMPLATEDESCRIPTION", strValue
    
    ' PSC 04/07/2005 MAR5 - Start
    xmlElem.setAttribute "POSTTOWEB", "0"
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "COMPRESSIONMETHOD")
    xmlElem.setAttribute "COMPRESSIONMETHOD", strValue
    ' PSC 04/07/2005 MAR5 - End

    'EP2_1903 GHun
    If strDestinationType = "EF" Then
        'Dim xmlRecipients As IXMLDOMNode
        'Dim xmlNode As IXMLDOMNode
        'Dim xmlEF As IXMLDOMElement
        'Dim xmlOutputType As IXMLDOMNode
        'Set xmlOutputType = xmlThisRequestDoc.createElement("OUTPUTTYPE")
        
        'Set xmlEF = xmlThisRequestDoc.createElement("EMAILFULFILMENT")
        'xmlOutputType.appendChild xmlEF
        
        xmlElem.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlRequestCaseTaskNode, "CASEID")
        xmlElem.setAttribute "EMAILRECIPIENTTYPE", xmlGetAttributeText(xmlRequestPrintAttributesNode, "EMAILRECIPIENTTYPE")
'        Set xmlRecipients = GetEmailRecipient(xmlGetAttributeText(xmlRequestCaseTaskNode, "CASEID"), xmlGetAttributeText(xmlRequestPrintAttributesNode, "EMAILRECIPIENTTYPE"))
'
'        For Each xmlNode In xmlRecipients.selectNodes("RECIPIENT")
'            xmlEF.appendChild xmlNode
'        Next
    
        strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CONTEXT")
        If Len(strValue) > 0 Then
            xmlElem.setAttribute "DOCUMENTLOCATION", strValue
        End If
        
        'xmlElem.appendChild xmlOutputType
        
    'add email recipients, app number and document location
    End If
    'EP2_1903 End

    xmlThisRequestNode.appendChild xmlElem

    'Call omPrintBO.PrintDocument
    
    On Error Resume Next
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    On Error GoTo PrintDocumentForTaskErr
    
    If Not objPrintBO Is Nothing Then
        strResponseFromPrint = objPrintBO.omRequest(xmlThisRequestDoc.xml)
        
        'CheckResponse for errors
        'MAR7 GHun
        Set xmlPrintResponseDoc = xmlLoad(strResponseFromPrint, cstrFunctionName)
        errCheckXMLResponseNode xmlPrintResponseDoc.documentElement, , True
        If strDestinationType = "W" Then
            vxmlResponseNode.appendChild xmlPrintResponseDoc.selectSingleNode(".//DOCUMENTCONTENTS")
        End If
        'MAR7 End
    End If
    
    'Call omTMBO.UpdateCaseTask
    'MAR7 GHun
    ' PSC /27/09/2005 MAR32 - Start
    Dim blnCompleteTask As Boolean
    blnCompleteTask = xmlGetAttributeAsBoolean(xmlRequestCaseTaskNode, "COMPLETETASKONPRINT")

    If blnCompleteTask Then
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        ' copy request node
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
        ' create CASETASK element
        Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
        ' clone search attributes
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
        If xmlGetAttributeText(xmlRequestCaseTaskNode, "INPUTPROCESS") = "" Then
            xmlElem.setAttribute "TASKSTATUS", "40" 'Complete
        Else
            xmlElem.setAttribute "TASKSTATUS", "20" 'Pending
        End If
        
        xmlThisRequestNode.appendChild xmlElem
        
        ' run Tm UpdateCaseTask request
        xmlThisResponseDoc.async = False
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        
        ' parse response
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    End If
    ' PSC /27/09/2005 MAR32 - End
    'MAR7 End

    gobjContext.SetComplete

PrintDocumentForTaskExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlRequestPrintAttributesNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
PrintDocumentForTaskErr:

    gobjContext.SetAbort
    
    GoTo PrintDocumentForTaskExit
    
End Sub

Private Sub FindUnitTaskSummary( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo FindUnitTaskSummaryExit
    Const cstrFunctionName As String = "FindUnitTaskSummary"
    
    'BMIDS585 No longer used
    'Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    'Dim xmlThisRequestNode As IXMLDOMNode
    'Dim xmlUnitListDoc As FreeThreadedDOMDocument40
    'Dim xmlUnitListNode As IXMLDOMNode
    'Dim xmlCaseTaskResponseNode As IXMLDOMNode
    'Dim xmlActivityRequestNode As IXMLDOMNode
    'Dim xmlOmigaUserResponseElem As IXMLDOMElement
    'Dim xmlUserCaseTaskList As IXMLDOMNodeList
    'Dim xmlElem As IXMLDOMElement
    'Dim xmlNode As IXMLDOMNode
    'Dim objOmOrg As Object
    'Dim strThisUserId As String
    'BMIDS585 End
    
    Dim xmlCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlUnitRequestNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    Dim xmlUserNode As IXMLDOMNode
    Dim xmlUserNodeList As IXMLDOMNodeList
    
    ' validate request ================================================================
    Set xmlUnitRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "UNIT")
    xmlCheckMandatoryAttribute xmlUnitRequestNode, "UNITID"
    
    'BMIDS585 No longer used
    'Set xmlActivityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "ACTIVITY")
    'xmlCheckMandatoryAttribute xmlActivityRequestNode, "SOURCEAPPLICATION"
    'xmlCheckMandatoryAttribute xmlActivityRequestNode, "ACTIVITYID"
        
    '' get user list ===================================================================
    'Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    'Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    '
    'xmlThisRequestNode.Attributes.removeNamedItem ("OPERATION")
    '
    'Set xmlElem = xmlThisRequestDoc.createElement("USERLIST")
    'Set xmlNode = xmlThisRequestNode.appendChild(xmlElem)
    '
    'Set xmlElem = xmlThisRequestDoc.createElement("UNITID")
    'xmlElem.Text = xmlGetAttributeText(xmlUnitRequestNode, "UNITID")
    'xmlNode.appendChild xmlElem
    '
    'Set xmlUnitListDoc = New FreeThreadedDOMDocument40
    'xmlUnitListDoc.async = False
    '
    'Debug.Print "FindUserList request:"
    'Debug.Print xmlThisRequestDoc.xml
    '
    'Set objOmOrg = gobjContext.CreateInstance("omOrg.OrganisationBO")
    'xmlUnitListDoc.loadXML objOmOrg.FindUserList(xmlThisRequestDoc.xml)
    'Set objOmOrg = Nothing
    '
    'Debug.Print "FindUserList response:"
    'Debug.Print xmlUnitListDoc.xml
    '
    'errCheckXMLResponseNode xmlUnitListDoc.selectSingleNode("RESPONSE"), , True
    'Set xmlUnitListNode = xmlUnitListDoc.selectSingleNode("RESPONSE/OMIGAUSERLIST")
    '
    '' create FindCaseTaskList request ==================================================
    '
    '' clear request node
    'xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    '
    '' set operation
    'xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "FindCaseTaskList"
    '
    '' create CASETASK element
    'Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    'xmlElem.setAttribute "OWNINGUNITID", xmlGetAttributeText(xmlUnitRequestNode, "UNITID")
    'xmlElem.setAttribute "TASKSTATUS", "I"
    'xmlElem.setAttribute "_ORDERBY_", "OWNINGUSERID,CASEID"
    'xmlThisRequestNode.appendChild xmlElem
    'BMIDS585 End
    
    ' run the request
    Set xmlCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlCaseTaskDoc.validateOnParse = False
    xmlCaseTaskDoc.setProperty "NewParser", True
    xmlCaseTaskDoc.async = False
       
    RunTmRequest vxmlRequestNode.ownerDocument, xmlCaseTaskDoc
    errCheckXMLResponseNode xmlCaseTaskDoc.selectSingleNode("RESPONSE"), , True
    
    Debug.Print "FindCaseTaskList response:"
    Debug.Print xmlCaseTaskDoc.xml
    
    'BMIDS585 Not longer used
    'For Each xmlUserNode In xmlUnitListNode.childNodes
    '
    '    Set xmlOmigaUserResponseElem = vxmlResponseNode.ownerDocument.createElement("OMIGAUSER")
    '
    '    strThisUserId = xmlGetNodeText(xmlUserNode, "USERID")
    '
    '    xmlOmigaUserResponseElem.setAttribute "USERID", strThisUserId
    '
    '    Set xmlUserCaseTaskList = _
    '        xmlCaseTaskDoc.selectNodes( _
    '            "RESPONSE/CASETASK[@OWNINGUSERID='" & strThisUserId & "']")
    '
    '    If xmlUserCaseTaskList.length = 0 Then
    '        xmlOmigaUserResponseElem.setAttribute "APPLICATIONS", "0"
    '        xmlOmigaUserResponseElem.setAttribute "TASKS", "0"
    '        xmlOmigaUserResponseElem.setAttribute "OUTSTANDINGTASKS", "0"
    '    Else
    '        GetUserTotals xmlOmigaUserResponseElem, xmlUserCaseTaskList
    '    End If
    '
    '    vxmlResponseNode.appendChild xmlOmigaUserResponseElem
    '
    'Next
    '
    'GetGroupTotals vxmlResponseNode
    'BMIDS585 End
    
    'BMIDS585 Change UnitTotals and UserAvergages to expected node names
    Set xmlTempElem = xmlCaseTaskDoc.selectSingleNode("RESPONSE/TASKSUMMARY/OMIGAUSER[@USERID = '_UNITTOTAL']")
    xmlChangeNodeName xmlTempElem, "OMIGAUSER", "UNITTOTALS"
    xmlTempElem.removeAttribute "USERID"
    
    Set xmlTempElem = xmlCaseTaskDoc.selectSingleNode("RESPONSE/TASKSUMMARY/OMIGAUSER[@USERID = '_USERAVG']")
    xmlChangeNodeName xmlTempElem, "OMIGAUSER", "USERAVERAGES"
    xmlTempElem.removeAttribute "USERID"
    
    Set xmlUserNodeList = xmlCaseTaskDoc.selectNodes("RESPONSE/TASKSUMMARY/*")
    For Each xmlUserNode In xmlUserNodeList
        vxmlResponseNode.appendChild xmlUserNode.cloneNode(True)
    Next
    
    Set xmlTempElem = vxmlResponseNode
    xmlTempElem.setAttribute "TYPE", "SUCCESS"
    'BM0585 End
    
FindUnitTaskSummaryExit:
    
    Set xmlUserNode = Nothing
    Set xmlUnitRequestNode = Nothing
    Set xmlCaseTaskDoc = Nothing
    Set xmlTempElem = Nothing
    Set xmlTempNode = Nothing
    Set xmlParentNode = Nothing
    Set xmlUserNodeList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub GetUserTotals( _
    ByVal vxmlOmigaUserResponseElem As IXMLDOMElement, _
    ByVal vxmlUserCaseTaskList As IXMLDOMNodeList)

'BMIDS585 No longer used, but sub cannot be deleted as OOSS will just replace it
'    On Error GoTo GetUserTotalsExit
'    Const cstrFunctionName As String = "GetUserTotals"
'
'    Dim xmlNode As IXMLDOMNode
'    Dim varCaseIds As Collection
'
'    Dim strThisAppNo As String
'
'    Dim intLoop As Integer, _
'        intOSTasks As Integer
'
'    Dim blnCaseHit As Boolean
'
'    Dim dtLimit As Date, _
'        dtTest As Date
'
'    dtLimit = Now()
'
'    Set varCaseIds = New Collection
'
'    For Each xmlNode In vxmlUserCaseTaskList
'
'        strThisAppNo = xmlGetAttributeText(xmlNode, "CASEID")
'        blnCaseHit = False
'
'        For intLoop = 1 To varCaseIds.Count
'            If varCaseIds.Item(intLoop) = strThisAppNo Then
'                blnCaseHit = True
'                Exit For
'            End If
'        Next
'
'        If blnCaseHit = False Then
'            varCaseIds.Add strThisAppNo
'        End If
'
'        If xmlAttributeValueExists(xmlNode, "TASKDUEDATEANDTIME") Then
'
'            dtTest = _
'                xmlGetAttributeText(xmlNode, "TASKDUEDATEANDTIME")
'
'            If dtTest < dtLimit Then
'                intOSTasks = intOSTasks + 1
'            End If
'
'        End If
'
'    Next
'
'    vxmlOmigaUserResponseElem.setAttribute "APPLICATIONS", varCaseIds.Count
'    vxmlOmigaUserResponseElem.setAttribute "TASKS", vxmlUserCaseTaskList.length
'    vxmlOmigaUserResponseElem.setAttribute "OUTSTANDINGTASKS", intOSTasks
'
'GetUserTotalsExit:
'
'    Set xmlNode = Nothing
'    Set varCaseIds = Nothing
'
'    errCheckError cstrFunctionName

End Sub

Private Sub GetGroupTotals(ByVal vxmlResponseNode As IXMLDOMNode)

'BMIDS585 No longer used, but sub cannot be deleted as OOSS will just replace it
'    On Error GoTo GetGroupTotalsExit
'    Const cstrFunctionName As String = "GetGroupTotals"
'
'    Dim xmlNode As IXMLDOMNode
'    Dim xmlNodeList As IXMLDOMNodeList
'    Dim xmlElem As IXMLDOMElement
'
'    Dim intApps As Integer, _
'        intTasks As Integer, _
'        intOSTasks As Integer
'
'    Set xmlNodeList = vxmlResponseNode.selectNodes("OMIGAUSER")
'    For Each xmlNode In xmlNodeList
'        intApps = intApps + xmlGetAttributeAsInteger(xmlNode, "APPLICATIONS")
'        intTasks = intTasks + xmlGetAttributeAsInteger(xmlNode, "TASKS")
'        intOSTasks = intOSTasks + xmlGetAttributeAsInteger(xmlNode, "OUTSTANDINGTASKS")
'    Next
'
'    Set xmlElem = vxmlResponseNode.ownerDocument.createElement("UNITTOTALS")
'    xmlElem.setAttribute "APPLICATIONS", intApps
'    xmlElem.setAttribute "TASKS", intTasks
'    xmlElem.setAttribute "OUTSTANDINGTASKS", intOSTasks
'
'    vxmlResponseNode.appendChild xmlElem
'
'    If xmlNodeList.length > 0 Then
'        Set xmlElem = vxmlResponseNode.ownerDocument.createElement("USERAVERAGES")
'        xmlElem.setAttribute "APPLICATIONS", Round(intApps / xmlNodeList.length)
'        xmlElem.setAttribute "TASKS", Round(intTasks / xmlNodeList.length)
'        xmlElem.setAttribute "OUTSTANDINGTASKS", Round(intOSTasks / xmlNodeList.length)
'        vxmlResponseNode.appendChild xmlElem
'    End If
'
'GetGroupTotalsExit:
'
'    Set xmlNode = Nothing
'    Set xmlNodeList = Nothing
'    Set xmlElem = Nothing
'
'    errCheckError cstrFunctionName
'BMIDS585 End

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

Private Sub GetNextStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlNextStageDoc As FreeThreadedDOMDocument40)
    
    On Error GoTo GetNextStageExit
    Const cstrFunctionName As String = "GetNextStage"
    
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
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "STAGEID"
    xmlCopyAttribute vxmlCurrentStageNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlRequestNode.appendChild xmlElem
    
    ' run base Tm GetNextStageTaskList request
    RunTmRequest xmlRequestDoc, vxmlNextStageDoc
    
    ' validate for next STAGE in response
    xmlCheckMandatoryNode vxmlNextStageDoc, "RESPONSE/STAGE"
    
GetNextStageExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

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

Private Sub CreateActivityRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCaseActivityRequestNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode)
    
    On Error GoTo CreateActivityRequestExit
    Const cstrFunctionName As String = "CreateActivityRequest"

    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
    
    Dim intCaseStageSeq As Integer
    'AQR SYS2266
    Dim objNTxBO As IOmTmNTxBO
       
    ' create CreateCaseStage request ==================================================
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    ' copy input request (no child nodes)
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    ' set operation name
    xmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseActivity"
    
    ' create CASECTIVITY request element
    Set xmlElem = xmlRequestDoc.createElement("CASEACTIVITY")
    xmlCopyAttribute vxmlCaseActivityRequestNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute vxmlCaseActivityRequestNode, xmlElem, "CASEID"
    xmlCopyAttribute vxmlCaseActivityRequestNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute vxmlCaseActivityRequestNode, xmlElem, "ACTIVITYINSTANCE"
    Set xmlCaseActivityNode = xmlRequestNode.appendChild(xmlElem)
    
    ' create CASESTAGE request element
    Set xmlElem = xmlRequestDoc.createElement("CASESTAGE")
    ' copy STAGEID from rules response
    xmlCopyAttribute vxmlNextCaseStageNode, xmlElem, "STAGEID"
    xmlElem.setAttribute "CASESTAGESEQUENCENO", "1"
    Set xmlCaseStageNode = xmlCaseActivityNode.appendChild(xmlElem)
    
    ' for each task returned by rules
    For Each xmlNode In vxmlNextCaseStageNode.childNodes
        ' create CASETASK element
        Set xmlElem = xmlRequestDoc.createElement("CASETASK")
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
    
    ' run base Tm CreateCaseStage request
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    'AQR SYS2266 DRC
    ' The New Transaction Class is used to ensure that the stage creation is committed
    ' before any automatic tasks are processed
    
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.RunTmRequest xmlRequestDoc, xmlResponseDoc
    Set objNTxBO = Nothing
    
CreateActivityRequestExit:
    
    Set xmlRequestNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
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
                        xmlCopyAttributeValue xmlStageTask, xmlNode, "TASKNAME", "CASETASKNAME"
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

Private Sub ProcessNotApplicableTasks( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageUpdateNode As IXMLDOMNode)
    
    On Error GoTo ProcessNotApplicableTasksExit
    Const cstrFunctionName As String = "ProcessNotApplicableTasks"
    
    ModifyOutstandingTasks _
        vxmlCurrentStageNode, vxmlCurrentStageUpdateNode, omiga4TASKSTATUSNotApplicable
        
ProcessNotApplicableTasksExit:
    
    errCheckError cstrFunctionName

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

Private Sub CreateApplicationStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlApplicationRequestNode As IXMLDOMNode, _
    ByVal vxmlStageDetailNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
    Optional ByVal vxmlResponseNode As IXMLDOMNode = Nothing)
    
    Const cstrFunctionName As String = "CreateApplicationStage"
    On Error GoTo CreateApplicationStageExit

    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objApplicationBO As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Set xmlRequestNode = xmlRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' drop OPERATION attrib.
    xmlRequestNode.Attributes.removeNamedItem "OPERATION"
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONSTAGE")
    Set xmlRequestNode = xmlRequestNode.appendChild(xmlElem)
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlApplicationRequestNode, "APPLICATIONNUMBER")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("STAGENUMBER")
    xmlElem.Text = xmlGetAttributeText(vxmlStageDetailNode, "STAGEID")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("STAGENAME")
    xmlElem.Text = xmlGetAttributeText(vxmlStageDetailNode, "STAGENAME")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("DATETIME")
    xmlElem.Text = Now()
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("STAGESEQUENCENO")
    xmlElem.Text = xmlGetAttributeText(vxmlStageDetailNode, "STAGESEQUENCENO")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlRequestDoc.createElement("CASESTAGESEQUENCENO")
    xmlElem.Text = xmlGetAttributeText(vxmlNextCaseStageNode, "CASESTAGESEQUENCENO")
    xmlRequestNode.appendChild xmlElem
    
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Debug.Print xmlRequestDoc.xml
    xmlResponseDoc.loadXML objApplicationBO.CreateApplicationStage(xmlRequestDoc.xml)
    Debug.Print xmlResponseDoc.xml
    ' temporary frig, ignore error from omiga
    ' errCheckXMLResponseNode xmlResponseDoc.selectSingleNode("RESPONSE")
    
    If Not vxmlResponseNode Is Nothing Then
    
        Set xmlElem = vxmlResponseNode.ownerDocument.createElement("APPLICATION")
        Set xmlNode = vxmlResponseNode.appendChild(xmlElem)
        
        xmlNode.appendChild xmlRequestNode.selectSingleNode("APPLICATIONNUMBER").cloneNode(True)
        xmlNode.appendChild xmlRequestNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").cloneNode(True)
        xmlNode.appendChild xmlRequestNode.selectSingleNode("STAGENUMBER").cloneNode(True)
        xmlNode.appendChild xmlRequestNode.selectSingleNode("STAGENAME").cloneNode(True)
        xmlNode.appendChild xmlRequestNode.selectSingleNode("STAGESEQUENCENO").cloneNode(True)
    
    End If
    
CreateApplicationStageExit:
    
    Set objApplicationBO = Nothing
    
    Set xmlRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub RunCompositeRequest(ByVal vxmlRequestNodes As Collection)
    
    Const cstrFunctionName As String = "RunCompositeRequest"
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
        
    RunTmRequest xmlThisRequestDoc
    
RunCompositeRequestExit:
    
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlInRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
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
    
    errCheckError cstrFunctionName
End Sub

'MAR1717 GHun Debugging code that is never called
'Private Sub Terminate( _
'    ByVal vxmlDoc As FreeThreadedDOMDocument40, _
'    ByVal vstrFunctionName As String)
'
'    vxmlDoc.Save ("c:\temp\OmTmBO.xml")
'
'    Shell _
'        "c:\program files\plus!\microsoft internet\iexplore.exe c:\temp\OmTmBO.xml", _
'        vbMaximizedFocus
'
'    gobjContext.SetAbort
'
'    Err.Raise _
'        oeUnspecifiedError, _
'        TypeName(Me) & "." & vstrFunctionName, _
'        "OmTmBO Terminate Request"
'
'End Sub
'MAR1717 End

Private Sub ImportAccountsIntoApplication( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ImportAccountsIntoApplicationExit
    
    Const cstrFunctionName As String = "ImportAccountsIntoApplication"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMElement
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseNode As IXMLDOMNode
    
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCustomerRequestNode As IXMLDOMNode
    Dim xmlCaseActivityRequestNode As IXMLDOMNode
    Dim xmlImportAccountsResponseNode As IXMLDOMNode
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlSrceNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    Dim objBO As Object
    Dim strResponse As String
    
    Set xmlCaseActivityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "ACTIVITYID"

    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "CUSTOMER", True)
    
    Set objBO = gobjContext.CreateInstance("omApp.ApplicationManagerBO")
    strResponse = objBO.ImportAccountsIntoApplication(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    Set objBO = Nothing
    
    Debug.Print "ImportAccountsIntoApplication response:"
    Debug.Print xmlThisResponseDoc.xml
    
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONNUMBER"
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONFACTFINDNUMBER"
    
    Set xmlImportAccountsResponseNode = _
        xmlThisResponseDoc.selectSingleNode("RESPONSE")

    ' create CreateActivity request =================================================
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST")
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "CreateActivity"
    
    ' clone CASEACTIVITY node from original request
    Set xmlNode = _
        xmlThisRequestNode.appendChild(xmlCaseActivityRequestNode.cloneNode(False))
    
    ' get APPLICATIONNUMBER node from CreateApplicationAndCustomer response
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    
    ' create CASEID attribute from APPLICATIONNUMBER
    xmlSetAttributeValue xmlNode, "CASEID", xmlSrceNode.Text
    
    ' create APPLICATION node from CreateApplicationAndCustomer response
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONFACTFINDNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    'xmlThisRequestNode.appendChild xmlCaseActivityRequestNode.cloneNode(False)
    xmlThisRequestNode.appendChild xmlElem.cloneNode(False)
    
    ' run CreateActivity ============================================================
    
    Debug.Print "CreateActivity request:"
    Debug.Print xmlThisRequestNode.xml
    
    CreateActivity xmlThisRequestNode, vxmlResponseNode
    
ImportAccountsIntoApplicationExit:
    
    Set xmlApplicationRequestNode = Nothing
    Set xmlCustomerRequestNode = Nothing
    Set xmlCaseActivityRequestNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub IsApplicationAtOfferStage( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)


    On Error GoTo IsApplicationAtOfferStageExit
'GD BMIDS00037 REWORK OF METHOD
    Const cstrFunctionName As String = "IsApplicationAtOfferStage"
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlCaseStageRequestNode As IXMLDOMNode
    Dim xmlCaseStageResponseNode As IXMLDOMNode
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
        
    Dim lngSequenceNumber As Long
    Dim lngTMOfferStage As Long
    Dim strTMIssueOfferTaskName As String
    Dim strTMReIssueOfferTaskName As String
    Dim strPattern As String
    
    Dim strCaseId As String
    Dim strSourceApp As String
    Dim strActivityId As String
    Dim lngTaskStatus As Long
            
    Dim strFreezeDataIndicator As String
    Dim xmlNodeList As IXMLDOMNodeList
    Dim iIndex As Integer
    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "ACTIVITYID"
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
        
    Set xmlCaseActivityNode = vxmlRequestNode.ownerDocument.createElement("CASEACTIVITY")
    
    vxmlRequestNode.appendChild xmlCaseActivityNode
    
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "CASEID"
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "SOURCEAPPLICATION"
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "ACTIVITYID"
    
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GetCurrentStage"
            
    ' run Tm GetCaseTasks request and get the stage information for the current stage
    RunTmRequest vxmlRequestNode.ownerDocument, xmlThisResponseDoc
    
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    Set xmlCaseStageResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/CASESTAGE")
    
    lngSequenceNumber = xmlGetAttributeAsLong(xmlCaseStageResponseNode, "STAGESEQUENCENO")
    
    ' get the global parameter
    lngTMOfferStage = GetGlobalParamAmount("TMOfferStage")
        
    If lngSequenceNumber > lngTMOfferStage Then
        ' Do we have a reissue offer task that is complete?
        strTMReIssueOfferTaskName = GetGlobalParamString("TMReIssueOffer")
        strPattern = "CASETASK[@TASKID='" & strTMReIssueOfferTaskName & "']"
        Set xmlNodeList = xmlCaseStageResponseNode.selectNodes(strPattern)
        If (xmlNodeList.length > 0) Then
            strFreezeDataIndicator = "1"
            'Loop thru all the re-issue offers (if any)
            For iIndex = 0 To (xmlNodeList.length - 1)
                lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
                If Not (IsValidationType("TaskStatus", lngTaskStatus, "C")) Then 'if any re-issue offers INcomplete then unfreeze
                    strFreezeDataIndicator = "0"
                    Exit For
                End If
            Next
        Else
            strFreezeDataIndicator = "1" ' handles either no reissue or a reissue which is not complete
        End If
        
        
        
    ElseIf lngSequenceNumber = lngTMOfferStage Then
        ' if the issue offer task is complete then is there a reissue offer task?
    
        strTMIssueOfferTaskName = GetGlobalParamString("TMIssueOffer") 'Get IssueOffer Task Name
        
        strPattern = "CASETASK[@TASKID='" & strTMIssueOfferTaskName & "']"
        Set xmlNodeList = xmlCaseStageResponseNode.selectNodes(strPattern) 'Get list of all IssueOffers
        If (xmlNodeList.length > 0) Then 'If there are some issue Offers
            strFreezeDataIndicator = "1" 'Check if any issue offers are not complete
            For iIndex = 0 To (xmlNodeList.length - 1)
                lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
                If Not (IsValidationType("TaskStatus", lngTaskStatus, "C")) Then
                    strFreezeDataIndicator = "0"
                    Exit For
                End If
            Next
            If (strFreezeDataIndicator = "1") Then 'All the Issue Offers are complete
                strTMReIssueOfferTaskName = GetGlobalParamString("TMReIssueOffer")

                strPattern = "CASETASK[@TASKID='" & strTMReIssueOfferTaskName & "']"

                Set xmlNodeList = xmlCaseStageResponseNode.selectNodes(strPattern)
                If (xmlNodeList.length > 0) Then 'If there are any re-issue offers
                    'strFreezeDataIndicator = "1" 'If any incomplete Freeze = '0"
                    For iIndex = 0 To (xmlNodeList.length - 1)
                        lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
                        If Not (IsValidationType("TaskStatus", lngTaskStatus, "C")) Then
                            strFreezeDataIndicator = "0" 'Any reissue offers Incomplete
                            Exit For
                        End If
                    Next
                End If
            Else
                'strFreezeDataIndicator = "0" 'Issue Offer Incomplete
            End If
            
        Else
            strFreezeDataIndicator = "0" 'Issue Offer not yet complete
        End If
   Else
        strFreezeDataIndicator = "0" 'Current Stage has not yet progressed to offer
   End If
    
    Set xmlResponseNode = vxmlResponseNode.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlResponseNode, "FREEZEDATAINDICATOR", strFreezeDataIndicator
    vxmlResponseNode.appendChild xmlResponseNode
        
IsApplicationAtOfferStageExit:
    
    Set xmlThisResponseDoc = Nothing
    Set xmlCaseStageRequestNode = Nothing
    Set xmlCaseStageResponseNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'DRC BMIDS670 New sub to replace IsAppliCationAtOffer
Private Sub FreezeUnFreezeApplication( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)


    On Error GoTo FreezeUnFreezeApplicationExit

    Const cstrFunctionName As String = "FreezeUnFreezeApplication"
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlCaseStageRequestNode As IXMLDOMNode
    Dim xmlCaseStageResponseNode As IXMLDOMNode
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlPreviousStagesDoc As FreeThreadedDOMDocument40
    Dim xmlPreviousStagesResponseNode As IXMLDOMNode
    
        
 
    Dim strPattern As String
    
    Dim strCaseId As String
    Dim strThisTaskId As String
    Dim strSourceApp As String
    Dim strActivityId As String
    Dim lngTaskStatus As Long
    Dim lngSequenceNumber As Long
    Dim blncomplete As Boolean
    Dim dteLatestCompletionDate As Date
    Dim strLatestCompletionDate As String
    Dim strFreezingTaskID As String
    
    
    ' Last date a freezing task was complete
    strLatestCompletionDate = "01/01/2000"
    dteLatestCompletionDate = CSafeDate(strLatestCompletionDate)
    strFreezingTaskID = ""
    
            
    Dim strFreezeDataIndicator As String
    Dim xmlNodeList As IXMLDOMNodeList
    Dim iIndex As Integer
    
    'Check Request XML validity
    
    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "ACTIVITYID"
    
        
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    
    'Retrieve tasks for current stage
        
    Set xmlCaseActivityNode = vxmlRequestNode.ownerDocument.createElement("CASEACTIVITY")
    
    vxmlRequestNode.appendChild xmlCaseActivityNode
    
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "CASEID"
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "SOURCEAPPLICATION"
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "ACTIVITYID"
    
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GetCurrentStage"
            
    ' run Tm GetCaseTasks request and get the stage information for the current stage
    RunTmRequest vxmlRequestNode.ownerDocument, xmlThisResponseDoc
    
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    Set xmlCaseStageResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/CASESTAGE")
    lngSequenceNumber = xmlGetAttributeAsLong(xmlCaseStageResponseNode, "STAGESEQUENCENO")
    
    
              
    strPattern = "CASETASK"
    Set xmlNodeList = xmlCaseStageResponseNode.selectNodes(strPattern)
    strFreezeDataIndicator = ""
     'Go Thru list looking for tasks that will freeze or unfreeze the app
     'If ANY unfreeze task is incomplete, Unfreeze the Application Data
     'If ALL freeze tasks are complete the Application Data is frozen
    
    If (xmlNodeList.length > 0) Then
        'Loop thru all the tasks on the current stage
        For iIndex = 0 To (xmlNodeList.length - 1)
            strThisTaskId = xmlGetAttributeText(xmlNodeList.Item(iIndex), "TASKID")
            lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
            blncomplete = IsValidationType("TaskStatus", lngTaskStatus, "C")
            'Is the Task an  incomplete "UnFreeze Data Task" ?
            If (IsFreezeUnfreezeDataTask("UnFreeze", strThisTaskId) And _
                Not (blncomplete)) Then
                    strFreezeDataIndicator = "0"
                    Exit For
            End If
        Next
        If (Len(strFreezeDataIndicator) = 0) Then
            For iIndex = 0 To (xmlNodeList.length - 1)
                ' Test for completed "Freeze Data" tasks
                'BMIDS692    Revisited Start
                strThisTaskId = xmlGetAttributeText(xmlNodeList.Item(iIndex), "TASKID")
                lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
                blncomplete = IsValidationType("TaskStatus", lngTaskStatus, "C")
                'BMIDS692    Revisited End
                If (IsFreezeUnfreezeDataTask("Freeze", strThisTaskId) And _
                    blncomplete) Then
                        strFreezeDataIndicator = "1"
                        'BMIDS 692 DRC 10/02/04 - START
                        If (dteLatestCompletionDate < _
                                xmlGetAttributeAsDate(xmlNodeList.Item(iIndex), "TASKSTATUSSETDATETIME")) Then
                                dteLatestCompletionDate = xmlGetAttributeAsDate(xmlNodeList.Item(iIndex), "TASKSTATUSSETDATETIME")
                                strFreezingTaskID = strThisTaskId
                        End If
                        'Exit For
                        'BMIDS 692 DRC 10/02/04 - END
                End If
                
            Next
        End If
    End If
    ' if there is no Task in the current stage that freezes/unfreezes the data, we need to check the
    ' previous stage for a freezing data task
    
        If (Len(strFreezeDataIndicator) = 0 And lngSequenceNumber > 1) Then
        
'Retrieve tasks for previous stages

' get previous CaseStage details =======================================================
    ' create FindArchiveStageList request
        Set xmlPreviousStagesDoc = New FreeThreadedDOMDocument40
        xmlPreviousStagesDoc.validateOnParse = False
        xmlPreviousStagesDoc.setProperty "NewParser", True
        xmlPreviousStagesDoc.async = False
        
        xmlSetAttributeValue vxmlRequestNode, "OPERATION", "FindArchiveStageList"
            ' add ACTIVITYINSTANCE to CASEACTIVITY node in request
        xmlCopyAttribute xmlCaseStageResponseNode, xmlCaseActivityNode, "ACTIVITYINSTANCE"
    
      ' run Tm FindArchiveStageList request and get the stage information for the previous stages
        RunTmRequest vxmlRequestNode.ownerDocument, xmlPreviousStagesDoc
        
        Set xmlPreviousStagesResponseNode = xmlPreviousStagesDoc.selectSingleNode("RESPONSE/CASESTAGE")
        'Check for an existing previous stage
        If Not (xmlPreviousStagesResponseNode Is Nothing) Then
          strPattern = "CASETASK"
          Set xmlNodeList = xmlPreviousStagesDoc.getElementsByTagName(strPattern)
          If (xmlNodeList.length > 0) Then
        'Loop thru all the tasks on the previous stage
            For iIndex = 0 To (xmlNodeList.length - 1)
                strThisTaskId = xmlGetAttributeText(xmlNodeList.Item(iIndex), "TASKID")
                lngTaskStatus = xmlGetAttributeAsLong(xmlNodeList.Item(iIndex), "TASKSTATUS")
                blncomplete = IsValidationType("TaskStatus", lngTaskStatus, "C")
                ' Test for a completed "Freeze Data" task
                If (IsFreezeUnfreezeDataTask("Freeze", strThisTaskId) And _
                    blncomplete) Then
                        strFreezeDataIndicator = "1"
                        'BMIDS 692 DRC 10/02/04 - START
                        If (dteLatestCompletionDate < _
                            xmlGetAttributeAsDate(xmlNodeList.Item(iIndex), "TASKSTATUSSETDATETIME")) Then
                            dteLatestCompletionDate = xmlGetAttributeAsDate(xmlNodeList.Item(iIndex), "TASKSTATUSSETDATETIME")
                            strFreezingTaskID = strThisTaskId
                        End If
                        'Exit For
                        'BMIDS 692 DRC 10/02/04 - END
                End If
                
            Next
           End If
            
        End If
    End If

    Set xmlResponseNode = vxmlResponseNode.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlResponseNode, "FREEZEDATAINDICATOR", strFreezeDataIndicator
    'BMIDS 692 DRC 10/02/04 - added optional Freezing task ID
    xmlSetAttributeValue xmlResponseNode, "FREEZETASKID", strFreezingTaskID
    vxmlResponseNode.appendChild xmlResponseNode
        
FreezeUnFreezeApplicationExit:
    
    Set xmlThisResponseDoc = Nothing
    Set xmlCaseStageRequestNode = Nothing
    Set xmlCaseStageResponseNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
Private Sub IssueOffer(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo IssueOfferVbErr
    
    Const cstrFunctionName As String = "IssueOffer"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlTMCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlQuoteNumberNode As IXMLDOMElement
    Dim xmlQuotationNode As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlRemodelQuotationRequest As IXMLDOMNode
    Dim xmlTMNewCaseTaskNode As IXMLDOMNode
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlTempReq As FreeThreadedDOMDocument40
    Dim xmlTempChildElem As IXMLDOMElement
    Dim xmlCreateResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlCaseStageElem As IXMLDOMElement
    Dim xmlQuotationRequestDoc As FreeThreadedDOMDocument40
    Dim xmlAppNode  As IXMLDOMNode
    Dim xmlAppApproveRequestDoc  As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskElem As IXMLDOMElement
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlCustomerNodeList As IXMLDOMNodeList
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlPrintAttributeNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    Dim xmlCaseTaskCopy As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlPanelLegalRepDetails As FreeThreadedDOMDocument40
    Dim xmlAppLegalRepNode As IXMLDOMNode
    
    Dim lngErrorNumber  As Long
    
    Dim objMsgTmBO As Object
    Dim objAppBO As Object
    Dim objAppQuoteBO As Object
    Dim objQuotation As Object
    Dim objAppProcBO  As Object
    Dim gobjEnquiry As Object
    Dim objAppThirdPartyBO As Object
    Dim objCustomerBO As Object             ' PSC 26/10/2005 MAR300
    Dim objRABO As Object                   ' MAR550
    Dim objCreditCheckBO As Object          ' MAR550
    
    Dim dteApplicationApprovalDate As Date
    
    Dim strTMReceivedOffer As String
    Dim strValidationType As String
    Dim strPattern As String
    Dim strApplicationApprovalDate As String
    Dim strAcceptedQuoteNumber As String
    Dim strMortgageSubQuoteNo As String
    Dim strTMRemodelMortgage As String
    Dim strResponse As String
    Dim strCondition As String
    Dim strTMLTVRiskAssessTask As String
    Dim strCaseActivityGuid As String
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strUserId As String
    Dim strUnitId As String
    Dim strTaskName As String
    Dim strPrinterDestination As String
    Dim strApplicant As String
    Dim strCustomerIdentifier As String
    Dim strTmpResponse As String
    Dim strStatus As String
    Dim strDirectoryGUID As String
    Dim strTypeOfApplication As String
    Dim strTMIssueOffer As String
    Dim strTMReIssueOffer  As String
    Dim strTMFTIssueOfferTaskID  As String
    Dim strTMReIssueOfferTaskID  As String
    Dim strCreateTaskID As String
    Dim strCurrentTaskId As String
        
    Dim blnTMAutoConditions As Boolean
    Dim blnTMAutoApprove   As Boolean
    Dim blnApplicationUnderReview As Boolean
    Dim blnValidMortgageSubQuote As Boolean
    Dim blnPrintingEnabled As Boolean
    Dim blnRaiseLegalRepTask As Boolean
    
    ' PSC 26/10/2005 MAR300 - Start
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlSearch As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNode
    Dim xmlCustomersIn As IXMLDOMNodeList
    Dim xmlCustomerIn As IXMLDOMNode
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlMessage As IXMLDOMNode
    Dim xmlOfferExpiryRequest As IXMLDOMNode
    Dim xmlOfferExpiryCaseTask As IXMLDOMNode
    
    Dim strTMOfferExpiryTaskId As String
    Dim lngTMDaysToOfferExpiry As Long
    Dim blnTMGenerateOfferExpiryTask As Boolean
    Dim dteDueDate As Date

    Dim blnUseAdminGetCustomerDetail As Boolean
    ' PSC 26/10/2005 MAR300 - End
    
    'MAR550
    Dim xmlRANode As IXMLDOMNode
    Dim xmlCCNode As IXMLDOMNode
    Dim xmlOverrideNode As IXMLDOMNode
    Dim xmlRANodeList As IXMLDOMNodeList
    Dim xmlCreditCheckNodeList As IXMLDOMNodeList
    Dim iRuleResult As Integer
    Dim strOverrideDate As String
    Dim strOverrideReason As String
    Dim bRARuleFailed As Boolean
    Dim bSMRuleFailed As Boolean
    Dim bKYCStatusFailed As Boolean
    Dim strTaskId As String
    Dim blnCheckRARules As Boolean
    Dim blnCheckSMReasons As Boolean
    Dim blnCheckKYCStatus As Boolean
    Dim iSeqNo As Integer
    Dim iMaxSeqNo As Integer
    Dim strKYCStatus As String
    Dim strCustomerName As String
    Dim strErrorNumber As String
    
    'BC MAR1549 31/03/2006 Begin
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    'BC MAR1549 31/03/2006 End
    
    Dim objNoTxBO As omTmNoTxBO    'EP974/MAR1848

    Dim sIsLegalRepToBeUsed As String    ' EP2_1293

    blnPrintingEnabled = False

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    ' check mandatory items
    xmlCheckMandatoryAttribute vxmlRequestNode, "UNITID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    
    'Check for Mandatory Application Node
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    'Check for Mandatory CaseTask Node
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "TASKID"
    
    Set objMsgTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Set objAppThirdPartyBO = gobjContext.CreateInstance("omApp.ApplicationThirdPartyBO")
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    Set objAppProcBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    Set objAppQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    Set objQuotation = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    
    ' PSC 26/10/2005 MAR300
    Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
  
    blnUseAdminGetCustomerDetail = GetGlobalParamBoolean("UseAdminGetCustDetailAppAccess")
        
    If blnUseAdminGetCustomerDetail Then
        
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        Set xmlSearch = xmlThisRequestDoc.createElement("SEARCH")
        xmlTempRequest.appendChild xmlSearch
        xmlSetAttributeValue xmlSearch, "CUSTOMERDATACHECK", "1"
        Set xmlCustomerList = xmlThisRequestDoc.createElement("CUSTOMERLIST")
        xmlSearch.appendChild xmlCustomerList
        
        Set xmlCustomersIn = xmlApplicationNode.selectNodes("CUSTOMER")
        
        For Each xmlCustomerIn In xmlCustomersIn
            Set xmlCustomer = xmlThisRequestDoc.createElement("CUSTOMER")
            xmlSetAttributeValue xmlCustomer, "NOLOCK", "1"
            Set xmlNode = xmlThisRequestDoc.createElement("CUSTOMERNUMBER")
            xmlCustomer.appendChild xmlNode
            xmlNode.Text = xmlGetAttributeText(xmlCustomerIn, "CUSTOMERNUMBER")
            Set xmlNode = xmlThisRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlCustomer.appendChild xmlNode
            xmlNode.Text = xmlGetAttributeText(xmlCustomerIn, "CUSTOMERVERSIONNUMBER")
            Set xmlNode = xmlThisRequestDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
            xmlCustomer.appendChild xmlNode
            xmlNode.Text = xmlGetAttributeText(xmlCustomerIn, "OTHERSYSTEMCUSTOMERNUMBER")
            xmlCustomerList.appendChild xmlCustomer
        Next
        
        Set xmlNode = xmlThisRequestDoc.createElement("CRITICALDATACONTEXT")
        xmlSearch.appendChild xmlNode
        xmlCopyAttribute xmlApplicationNode, xmlNode, "APPLICATIONNUMBER"
        xmlCopyAttribute xmlApplicationNode, xmlNode, "APPLICATIONFACTFINDNUMBER"
        xmlCopyAttribute xmlApplicationNode, xmlNode, "APPLICATIONPRIORITY"
        xmlCopyAttribute xmlCaseTaskNode, xmlNode, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTaskNode, xmlNode, "ACTIVITYID"
        xmlCopyAttribute xmlCaseTaskNode, xmlNode, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlCaseTaskNode, xmlNode, "STAGEID"
        
        'EP974/MAR1848 - Call GetAndSynchroniseCustomerDetails in a separate transaction
        'strResponse = objCustomerBO.GetAndSynchroniseCustomerDetails(xmlTempRequest.xml)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETANDSYNCHRONISECUSTOMERDETAILS"
        Set objNoTxBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".OmTmNoTxBO")
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        'EP974/MAR1848 End
        
        Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
        If xmlGetAttributeAsBoolean(xmlThisResponseDoc.documentElement, "DATACHANGED") Then
            On Error GoTo WarningExit
            errThrowError cstrFunctionName, 8538
        End If
    End If
    ' PSC 26/10/2005 MAR300 - End

'Step:
    
    'GetApplication Data
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)

    strResponse = objAppBO.GetApplicationData(xmlThisRequestDoc.xml)

    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)

    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    strApplicationApprovalDate = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/APPLICATIONAPPROVALDATE")
    strTypeOfApplication = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
    ' EP2_1293 - Get ISLEGALREPTOBEUSED value.
    sIsLegalRepToBeUsed = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/ISLEGALREPTOBEUSED")

    blnTMAutoApprove = GetGlobalParamBoolean("TMAutoApprove")
    
    If blnTMAutoApprove = False Then
        If Len(strApplicationApprovalDate) = 0 Then
            errThrowError cstrFunctionName, 4814
        End If
    Else

        dteApplicationApprovalDate = CSafeDate(strApplicationApprovalDate)

    End If

    strAcceptedQuoteNumber = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/ACCEPTEDQUOTENUMBER")

    If Len(strAcceptedQuoteNumber) = 0 Then

        errThrowError cstrFunctionName, 4818
    End If

   'MAR550 Start
   bRARuleFailed = False
   bSMRuleFailed = False
   bKYCStatusFailed = False
   strTaskId = ""
   strCustomerName = ""
   
   'MAR550  Check for outstanding Risk Assessment rules
    
    blnCheckRARules = GetGlobalParamBoolean("TMPreventOfferWhenOutSCARules")
    
    If (blnCheckRARules = True) Then
        
        Set objRABO = gobjContext.CreateInstance(gstrRISK_ASSESSMENT & ".RiskAssessmentBO")
            
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    
        Set xmlRANode = xmlThisRequestDoc.createElement("RISKASSESSMENT")
        xmlTempRequest.appendChild xmlRANode
    
        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONNUMBER")
        xmlRANode.appendChild xmlNode
        xmlNode.Text = strApplicationNumber

        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlRANode.appendChild xmlNode
        xmlNode.Text = strApplicationFactFindNumber
    
        strResponse = objRABO.GetLatestRiskAssessment(xmlTempRequest.xml)
        Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
   
        ' Raise error if any error other than record not found
        If errGetOmigaErrorNumber(lngErrorNumber) <> oeRecordNotFound And lngErrorNumber <> 0 Then
            errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
        End If
    
        If Not xmlGetNode(xmlThisResponseDoc, ".//RISKASSESSMENTRULELIST") Is Nothing Then
            Set xmlRANodeList = xmlThisResponseDoc.selectNodes(".//RISKASSESSMENTRULELIST/RISKASSESSMENTRULE")

            For Each xmlRANode In xmlRANodeList

                iRuleResult = CInt(xmlGetNodeText(xmlRANode, "RARULERESULT"))

                If (iRuleResult <> 1) Then
                    If Not xmlRANode.selectSingleNode("RISKASSESSMENTRULEOVERRIDE") Is Nothing Then
                        'Check for override
                        Set xmlOverrideNode = xmlRANode.selectSingleNode("RISKASSESSMENTRULEOVERRIDE")
                    
                        strOverrideDate = xmlGetNodeText(xmlOverrideNode, "RAOVERRIDEDATETIME")
                        strOverrideReason = xmlGetNodeText(xmlOverrideNode, "RAOVERRIDEREASONCODE")
                    
                        If (Len(strOverrideDate) = 0) And (Len(strOverrideReason) = 0) Then
                            'No override exists
                            bRARuleFailed = True
                            Exit For    ' Outstanding Risk Assessment found
                        End If

                    End If
                End If
            Next
        End If

        If (bRARuleFailed = True) Then
    
            'Set up task id
            strTaskId = GetMandatoryGlobalParamString("TMOutstandingCARuleATOffer")
        
            'Set up error number
            strErrorNumber = 8534
        
        End If
  
    End If
    
    'MAR550  Check for outstanding Strategy Manager reasons
    
    blnCheckSMReasons = GetGlobalParamBoolean("TMPreventOfferWhenOutSSMRules")
    
    If ((blnCheckSMReasons = True) And (bRARuleFailed = False)) Then
        
        Set objCreditCheckBO = gobjContext.CreateInstance(gstrCREDIT_CHECK & ".CreditCheckBO")
        
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    
        Set xmlRANode = xmlThisRequestDoc.createElement("APPLICATIONCREDITCHECKDETAILS")
        xmlTempRequest.appendChild xmlRANode
    
        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONNUMBER")
        xmlRANode.appendChild xmlNode
        xmlNode.Text = strApplicationNumber

        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlRANode.appendChild xmlNode
        xmlNode.Text = strApplicationFactFindNumber
    
        strResponse = objCreditCheckBO.FindApplicationCreditCheckDetailsList(xmlTempRequest.xml)
        Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
   
        ' Raise error if any error other than record not found
        If errGetOmigaErrorNumber(lngErrorNumber) <> oeRecordNotFound And lngErrorNumber <> 0 Then
            errRaiseXMLResponseNode xmlThisResponseDoc.documentElement
        End If
    
        'Look for the latest Credit Check
        If Not xmlGetNode(xmlThisResponseDoc, ".//APPLICATIONCREDITCHECKLIST") Is Nothing Then
            Set xmlRANodeList = xmlThisResponseDoc.selectNodes(".//APPLICATIONCREDITCHECKLIST/APPLICATIONCREDITCHECK")

            iMaxSeqNo = 1
            For Each xmlRANode In xmlRANodeList
                iSeqNo = CInt(xmlGetNodeText(xmlRANode, "SEQUENCENUMBER"))
                If (iSeqNo > iMaxSeqNo) Then
                    iMaxSeqNo = iSeqNo
                End If
            Next
        
            For Each xmlRANode In xmlRANodeList
                iSeqNo = CInt(xmlGetNodeText(xmlRANode, "SEQUENCENUMBER"))
                
                If (iSeqNo = iMaxSeqNo) Then
                    
                    'Get all the latest strategy manager reasons
                    Set xmlCreditCheckNodeList = xmlRANode.selectNodes(".//CREDITCHECKREASONCODE")
                            
                    'Look for any reasons which are not overridden
                    For Each xmlCCNode In xmlCreditCheckNodeList
                        If xmlCCNode.selectSingleNode("SMOVERRIDEREASONRESULT") Is Nothing Then
                            'No override exists - set flag
                            bSMRuleFailed = True
                            Exit For   ' Outstanding Strategy Manager reason found
                        End If
             
                    Next
                End If
            Next
 
        End If
        
        If (bSMRuleFailed = True) Then

              'Set up task id
            strTaskId = GetMandatoryGlobalParamString("TMOutstandingSMRuleATOffer")
        
            'Set up error number
            strErrorNumber = 8535  ' Strategy Manager Reason codes outstanding
        End If
  
    End If

    'MAR550  Check that each customer has a KYC Status of Satisfied
    
    blnCheckKYCStatus = GetGlobalParamBoolean("TMCheckKYCStatusAtOffer")
    
    If ((blnCheckKYCStatus = True) And (bRARuleFailed = False) And (bSMRuleFailed = False)) Then
        
        Set xmlCustomersIn = xmlApplicationNode.selectNodes("CUSTOMER")
        
        'Check that each customer has a KYC status of 'Satisfied'
        For Each xmlCustomerIn In xmlCustomersIn
            
            'Get Customer details
                        
            Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
            Set xmlSearch = xmlThisRequestDoc.createElement("SEARCH")
            xmlTempRequest.appendChild xmlSearch
            
            Set xmlCustomer = xmlThisRequestDoc.createElement("CUSTOMER")
            xmlSetAttributeValue xmlCustomer, "NOLOCK", "1"
            Set xmlNode = xmlThisRequestDoc.createElement("CUSTOMERNUMBER")
            xmlCustomer.appendChild xmlNode
            xmlNode.Text = xmlGetAttributeText(xmlCustomerIn, "CUSTOMERNUMBER")
            Set xmlNode = xmlThisRequestDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlCustomer.appendChild xmlNode
            xmlNode.Text = xmlGetAttributeText(xmlCustomerIn, "CUSTOMERVERSIONNUMBER")
            
            xmlSearch.appendChild xmlCustomer
        
            strResponse = objCustomerBO.GetCustomerDetails(xmlTempRequest.xml)
            
            Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
            strKYCStatus = xmlGetNode(xmlThisResponseDoc, ".//CUSTOMERVERSION/CUSTOMERKYCSTATUS").Text
            
            If (strKYCStatus <> "Satisfied") Then
                bKYCStatusFailed = True
                
                strCustomerName = xmlGetNode(xmlThisResponseDoc, ".//CUSTOMERVERSION/FIRSTFORENAME").Text
                strCustomerName = strCustomerName & " " & xmlGetNode(xmlThisResponseDoc, ".//CUSTOMERVERSION/SURNAME").Text
                
                Exit For ' Customer found with unsatisfied KYC status
            
            End If
       
        Next
    
        If (bKYCStatusFailed = True) Then
        
            'Set up task id
            strTaskId = GetMandatoryGlobalParamString("TMUnsatifiedKYCStatusATOffer")
        
            'Set up error number
            strErrorNumber = 8503  ' KYC Status not satisfied

        End If
    
    End If
        
    'If a task is to be created and error reported, then do this now
    If ((bRARuleFailed = True) Or (bSMRuleFailed = True) Or (bKYCStatusFailed = True)) Then
        
        'Check to see if outstanding task exists
        Set xmlTempRequest = vxmlRequestNode.cloneNode(True)

        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        xmlThisRequestDoc.appendChild xmlTempRequest

        xmlSetAttributeValue xmlTempRequest, "OPERATION", "FindCaseTaskList"

        Set xmlTMCaseTaskNode = xmlGetMandatoryNode(xmlThisRequestDoc, "REQUEST/CASETASK")
        Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")

        xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
        xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
        xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"

        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTaskId
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"

        xmlTempRequest.removeChild xmlTMCaseTaskNode
        xmlTempRequest.appendChild xmlTMNewCaseTaskNode

        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True

        On Error Resume Next

        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

        On Error GoTo IssueOfferVbErr

        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)

        ' If no record is found then create the task
        If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then

            Set xmlThisRequestDoc = New FreeThreadedDOMDocument40

            xmlThisRequestDoc.validateOnParse = False
            xmlThisRequestDoc.setProperty "NewParser", True
            Set xmlNewRequest = vxmlRequestNode.cloneNode(False)
            xmlThisRequestDoc.appendChild xmlNewRequest

            Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
            xmlNewRequest.appendChild xmlTMNewCaseTaskNode
            xmlNewRequest.appendChild xmlApplicationNode.cloneNode(False)
            xmlNewRequest.setAttribute "OPERATION", "CreateAdhocCaseTask"

            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTaskId
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "0"

            CreateAdhocCaseTask xmlNewRequest

        ElseIf lngErrorNumber <> 0 Then
            errRaiseXMLResponseNode xmlThisResponseDoc.documentElement

        End If

        'Report the error
        errThrowError cstrFunctionName, strErrorNumber, "", strCustomerName
           
    End If
    
   'MAR550 End

'Step:

    'EP2_1293 - Check ISLEGALREPTOBEUSED flag.
    If sIsLegalRepToBeUsed = "1" Then

        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
         
         Set xmlAppLegalRepNode = xmlMakeNodeElementBased(vxmlRequestNode.selectSingleNode("//APPLICATION"), True, "APPLICATIONLEGALREP")
         xmlThisRequestDoc.appendChild xmlAppLegalRepNode
         
         strResponse = objAppThirdPartyBO.GetLegalRep(xmlThisRequestDoc.xml)
         
         Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
         lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
         
         If errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then
             CreateLegalRepTask vxmlRequestNode
             errThrowError cstrFunctionName, 8500, "LegalRepInactive"
         ElseIf lngErrorNumber <> 0 Then
             errThrowError cstrFunctionName, lngErrorNumber
         End If
         
         strDirectoryGUID = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLEGALREP/DIRECTORYGUID")
         If Len(strDirectoryGUID) = 0 Then
             CreateLegalRepTask vxmlRequestNode
             errThrowError cstrFunctionName, 8501, "LegalRepInactive"
         Else
             Set xmlPanelLegalRepDetails = New FreeThreadedDOMDocument40
             Set xmlPanelLegalRepDetails = GetPanelLegalRepDetails(vxmlRequestNode)
             strStatus = xmlGetAttributeText(xmlPanelLegalRepDetails.selectSingleNode("//PANELLEGALREP"), "STATUS")
         
             If IsValidationType("LegalRepStatus", strStatus, "I") = True Then
                 CreateLegalRepTask vxmlRequestNode
                 errThrowError cstrFunctionName, 8502, "LegalRepInactive"
             End If
             
         End If
    End If          '  ISLEGALREPTOBEUSED = 1 - EP2_1293
    
'Step:


    ' Get the appropriate Mortage subquote number from the quotation
    Set xmlQuotationRequestDoc = New FreeThreadedDOMDocument40
    Set xmlQuotationRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "QUOTATION")
    Set xmlQuotationNode = xmlGetNode(xmlQuotationRequestDoc, ".//QUOTATION")
    Set xmlNode = xmlQuotationRequestDoc.createElement("QUOTATIONNUMBER")
    xmlNode.Text = strAcceptedQuoteNumber
    xmlQuotationNode.appendChild xmlNode
        
    strResponse = objQuotation.GetData(xmlQuotationRequestDoc.xml)
    
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    strMortgageSubQuoteNo = xmlGetNodeText(xmlThisResponseDoc, ".//MORTGAGESUBQUOTENUMBER")
 
'Step:

    ' Check for Outstanding remodel quotation

    Set xmlRemodelQuotationRequest = vxmlRequestNode.cloneNode(True)
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.appendChild xmlRemodelQuotationRequest
    
    xmlSetAttributeValue xmlRemodelQuotationRequest, "OPERATION", "FindCaseTaskList"
    
    Set xmlTMCaseTaskNode = xmlGetMandatoryNode(xmlThisRequestDoc, "REQUEST/CASETASK")
   
    Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
    
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
    
    strTMRemodelMortgage = GetGlobalParamString("TMRemodelMortgage")
    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMRemodelMortgage
    
    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"

    xmlRemodelQuotationRequest.removeChild xmlTMCaseTaskNode
    xmlRemodelQuotationRequest.appendChild xmlTMNewCaseTaskNode
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    
    On Error Resume Next

    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
    
    On Error GoTo IssueOfferVbErr
    
    Set xmlThisRequestDoc = Nothing
    Set xmlTMNewCaseTaskNode = Nothing
    
    lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
    
    If lngErrorNumber <> 0 Then

        If errGetOmigaErrorNumber(lngErrorNumber) <> 500 Then
            errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        End If

    Else

        errThrowError cstrFunctionName, 4815

    End If
   
'Step:

    ' Check that the application is not underreview

    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "APPLICATIONREVIEWHISTORY")

    strResponse = objAppBO.IsAppUnderReview(xmlThisRequestDoc.xml)
    
    ' check for errors
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    blnApplicationUnderReview = xmlGetNodeAsBoolean(xmlThisResponseDoc.documentElement, "UNDERREVIEWINDICATOR")

    If blnApplicationUnderReview = True Then
        errThrowError cstrFunctionName, 4819
    End If
    
'Step:

    ' Validate the current accepted quotation

    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "BASICQUOTATIONDETAILS")

    Set xmlApplicationNode = xmlGetNode(xmlThisRequestDoc.documentElement, "BASICQUOTATIONDETAILS")
    
    Set xmlNode = xmlThisRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    If Not xmlNode Is Nothing Then
        xmlNode.Text = strMortgageSubQuoteNo  ' found from earlier call to get quotation data
    End If
    
    If Not xmlApplicationNode Is Nothing Then
        xmlApplicationNode.appendChild xmlNode
    End If
    
    strResponse = objAppQuoteBO.ValidateMortgageSubQuote(xmlThisRequestDoc.xml)

    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)

    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    blnValidMortgageSubQuote = xmlGetNodeAsBoolean(xmlThisResponseDoc.documentElement, "VALIDMORTGAGESUBQUOTE")

    If blnValidMortgageSubQuote = False Then
        errThrowError cstrFunctionName, 4815
    End If
    
'Step:
    
    'Remodel repayment costs and APR
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)
    Set xmlApplicationNode = xmlGetNode(xmlThisRequestDoc.documentElement, "APPLICATION")
    Set xmlQuoteNumberNode = xmlGetNode(xmlApplicationNode, "ACCEPTEDQUOTENUMBER")
    
    If xmlQuoteNumberNode Is Nothing Then
        Set xmlQuoteNumberNode = xmlThisRequestDoc.createElement("ACCEPTEDQUOTENUMBER")
        xmlApplicationNode.appendChild xmlQuoteNumberNode
    End If
    
    If Len(xmlQuoteNumberNode.Text) = 0 Then
        xmlQuoteNumberNode.Text = strAcceptedQuoteNumber
    End If
            
    strResponse = objAppQuoteBO.RemodelQuotation(xmlThisRequestDoc.xml)

    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)

    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
    Set xmlNode = xmlThisResponseDoc.selectSingleNode("//RESPONSE")
    
'Step:

    If Not xmlNode Is Nothing Then
        If xmlGetAttributeText(xmlNode, "TYPE") = "WARNING" Then
            'Check the warning raised is 7020
            ' Need to generate a new risk assessment task and halt processing ,first check to see
            'if one is outstanding if it is do not need to regenerate the task
            
            If errGetMessageText(7020) = xmlNode.selectSingleNode(".//MESSAGETEXT").Text Then
                
                'Check whether a task 'TMLTVRiskAssess', which is not actioned, exists

                'First get CaseActivityGuid
                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETCASEACTIVITY"
                Set xmlTMCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
                Set xmlNode = xmlThisRequestDoc.createElement("CASEACTIVITY")
                xmlNewRequest.appendChild xmlNode
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "CASEID"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "ACTIVITYID"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "ACTIVITYINSTANCE"
                
                strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
                Set xmlThisResponseDoc = xmlLoad(strResponse, TypeName(Me))
                errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
                Set xmlNode = xmlGetMandatoryNode(xmlThisResponseDoc.documentElement, "CASEACTIVITY")
                strCaseActivityGuid = xmlGetMandatoryAttributeText(xmlNode, "CASEACTIVITYGUID")
                
                'Now check for existance of task 'TMLTVRiskAssess'
                strTMLTVRiskAssessTask = GetGlobalParamString("TMLTVRiskAssess")
    
                strCondition = "CASEACTIVITYGUID = " & "0x" & strCaseActivityGuid
                strCondition = strCondition & " AND STAGEID ='" _
                                & xmlGetMandatoryAttributeText(xmlCaseTaskNode, "STAGEID") & "'"
                strCondition = strCondition & " AND CASESTAGESEQUENCENO ='" _
                                & xmlGetMandatoryAttributeText(xmlCaseTaskNode, "CASESTAGESEQUENCENO") & "'"
                strCondition = strCondition & " AND TASKID ='" & strTMLTVRiskAssessTask & "'"
                strCondition = strCondition & " AND TASKSTATUS ='10'"
                If Not adoCheckRecordExists("CASETASK", strCondition) Then
                    'Generate a new task
                    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
                    xmlThisRequestDoc.validateOnParse = False
                    xmlThisRequestDoc.setProperty "NewParser", True
                    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                    xmlThisRequestDoc.appendChild xmlNewRequest
                    Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
                    xmlNewRequest.appendChild xmlTMNewCaseTaskNode
                    xmlNewRequest.setAttribute "OPERATION", "CREATECASETASK"
                    
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMLTVRiskAssessTask
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "1"
                    'BMIDS815 GHun
                    xmlSetSysDateToNodeAttrib xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME"
                    'BMIDS815 End
                    
                    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
                    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
                End If
                GoTo WarningExit
            End If
        End If
    End If

'Step:
    
    blnTMAutoConditions = GetGlobalParamBoolean("TMAutoConditions")

    If blnTMAutoConditions Then
    
        'BC MAR1425 Begin
'        Set adoConn = New ADODB.Connection
'        Set adoCommand = New ADODB.Command
'        Set adoStream = New ADODB.Stream

'        'BC MAR1549 Begin
'        ' Build the request and call 'Conditions.DeleteAllApplicationConditions'
'
'        Set adoConn = New ADODB.Connection
'        Set adoCommand = New ADODB.Command
'        Set adoStream = New ADODB.Stream
'
'        Set adoParameter = adoCommand.CreateParameter("@pApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
'        adoCommand.Parameters.Append adoParameter
'
'        With adoConn
'            .ConnectionString = adoGetDbConnectString
'            .Open
'        End With
'
'        ' Use stream object to retrieve XML from the database
'        adoStream.Open
'
'        With adoCommand
'            .CommandText = "USP_DELETEUNMODIFIEDCONDITIONS"
'            .CommandType = adCmdStoredProc
'            .ActiveConnection = adoConn
'            .Properties("Output Stream") = adoStream
'            .Execute , , adExecuteStream
'        End With
'
'        adoConn.Close
'        adoStream.Close
'
'        'BC MAR1549 End
        'BC MAR1425 End
        
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEDEFAULTAPPLICATIONCONDITIONS"
    
        Set xmlTableNode = xmlDoc.createElement("APPLICATIONCONDITIONS")
        xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
        xmlSetAttributeValue xmlTableNode, "APPLICATIONFACTFINDNUMBER", strApplicationFactFindNumber
        xmlNewRequest.appendChild xmlTableNode
        
        strResponse = objAppProcBO.OmAppProcRequest(xmlNewRequest.xml)
        
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.loadXML (strResponse)
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
            
    End If
    
'Step:

    ' Print Offer document
    Dim strOutputDocument As String
    
    strOutputDocument = xmlGetAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT")
        
    If Len(strOutputDocument) > 0 Then
              
        Set xmlOut = New FreeThreadedDOMDocument40

        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
        
        Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
        xmlThisCaseTaskDoc.validateOnParse = False
        xmlThisCaseTaskDoc.setProperty "NewParser", True
        
        blnPrintingEnabled = True
            
        'Do the printing thing
        On Error Resume Next
    
        On Error GoTo IssueOfferVbErr

        If Not gobjEnquiry Is Nothing Then
            
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "GetPrintAttributes"
                      
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("FINDATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            xmlTempElem.setAttribute "HOSTTEMPLATEID", strOutputDocument
              
            strResponse = gobjEnquiry.omRequest(xmlThisCaseTaskElem.xml)
            Set xmlOut = xmlLoad(strResponse, cstrFunctionName)
            
            ' Error Check

            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True
            
            Set xmlTempElem = xmlOut.selectSingleNode("RESPONSE/ATTRIBUTES")
            
            If xmlTempElem.getAttribute("INACTIVEINDICATOR") = "1" Then
              errThrowError cstrFunctionName, 4820, "", "IssueOffer"
            End If
            
            If xmlTempElem.getAttribute("DEFAULTCOPIES") = "" Then
              errThrowError cstrFunctionName, 4823, "", "IssueOffer"
            End If
            
            strPrinterDestination = xmlGetAttributeText(xmlTempElem, "PRINTERDESTINATIONTYPE")
            
            ' EP2_1433 - use combovalidation types
            ' strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
'            If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
'                And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0) _
'                And (InStr(1, strPrinterDestination, "W", vbTextCompare) > 0)) Then     'MAR7 GHun
'               errThrowError cstrFunctionName, 4822, strPrinterDestination, "IssueOffer"
'            End If
            If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
            And IsValidationType("PrinterDestination", strPrinterDestination, "R") _
            And IsValidationType("PrinterDestination", strPrinterDestination, "W") _
            Then
                errThrowError cstrFunctionName, 4822, GetComboText("PrinterDestination", CInt(strPrinterDestination)), "IssueOffer"
            End If

                
            Set xmlPrintAttributeNode = xmlTempElem
            Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
            'Call OmTmBo.PrintDocumentForTask
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "PrintDocumentForTask"
            xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERID"
            xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "UNITID"
            xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERAUTHORITYLEVEL"
            
            Set xmlCaseTaskCopy = xmlCaseTaskNode.cloneNode(True)
            xmlThisCaseTaskNode.appendChild xmlCaseTaskCopy
                
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("APPLICATION")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            'use the request applicationnode
            Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
            
            xmlCopyAttribute xmlApplicationNode, xmlTempNode, "APPLICATIONFACTFINDNUMBER"
            
            ' PSC 04/07/2005 MAR5 - Start
            xmlSetAttributeValue xmlTempNode, "PRINTKFI", "1"
            xmlSetAttributeValue xmlTempNode, "QUOTATIONNUMBER", strAcceptedQuoteNumber
            xmlSetAttributeValue xmlTempNode, "MORTGAGESUBQUOTENUMBER", strMortgageSubQuoteNo
            ' PSC 04/07/2005 MAR5 - End

            Set xmlCustomerNode = xmlApplicationNode.selectSingleNode("./CUSTOMER")
            ' Note these Attributes are dummies as the OfferTemplate method in omPDM gets all the customer info
            ' again - they are needed here because omTM.PrintDocumentforTask wants them
             xmlCopyAttribute xmlCustomerNode, xmlTempNode, "CUSTOMERVERSIONNUMBER"
             
            xmlCopyAttributeValue xmlCustomerNode, xmlCaseTaskCopy, "CUSTOMERNUMBER", "CUSTOMERIDENTIFIER"
            
            xmlThisCaseTaskNode.appendChild xmlTempNode
            
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("PRINTATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DEFAULTCOPIES", "COPIES"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "PDMMETHOD", "METHODNAME"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE"     'MAR7 GHun
            
            ' EP2_1433
            ' xmlTempElem.setAttribute "DESTINATIONTYPE", Left(strPrinterDestination, 1)
            xmlTempElem.setAttribute "DESTINATIONTYPE", GetValidationTypeForValueID("PrinterDestination", strPrinterDestination)
            Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("PRINTER")
            xmlTempElem.setAttribute "COMPRESSIONMETHOD", xmlGetAttributeText(xmlPrinterNode, "COMPRESSIONMETHOD")  'MAR7 GHun
             
             ' EP2_1433 - use combovalidation types
            If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
            Or IsValidationType("PrinterDestination", strPrinterDestination, "W") _
            Then
                xmlTempElem.setAttribute "PRINTER", xmlGetAttributeText(xmlPrinterNode, "PRINTERNAME")
            Else    '"Remote" Or "File"
                 xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "REMOTEPRINTERLOCATION", "PRINTER"
            End If
            
            xmlThisCaseTaskNode.appendChild xmlTempNode
            
            PrintDocumentForTask xmlThisCaseTaskElem, vxmlResponseNode  'MAR7 GHun
                
        End If

    End If
    
    'PSC 03/02/2006 MAR1156 - Start
    ' Perform Pack Processing
    Dim strPackControlNumber As String
    strPackControlNumber = xmlGetAttributeText(xmlCaseTaskNode, "PACKCONTROLNUMBER")
    
    If Len(strPackControlNumber) > 0 Then
        SendPack vxmlRequestNode, strPackControlNumber
    End If
    'PSC 03/02/2006 MAR1156 - End
     
'Step:

    'BG Construct the element based request:

    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")

    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    
    Set xmlTempReq = New FreeThreadedDOMDocument40
    xmlTempReq.validateOnParse = False
    xmlTempReq.setProperty "NewParser", True
    Set xmlTempElem = xmlTempReq.createElement("APPLICATIONOFFER")
    
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONNUMBER")
    xmlTempChildElem.Text = strApplicationNumber
    xmlTempElem.appendChild xmlTempChildElem
    
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempChildElem.Text = strApplicationFactFindNumber
    xmlTempElem.appendChild xmlTempChildElem
       
    'AW 19/03/07    EP2_1897
    Set xmlTempChildElem = xmlTempReq.createElement("QUOTATIONNUMBER")
    xmlTempChildElem.Text = strAcceptedQuoteNumber
    xmlTempElem.appendChild xmlTempChildElem
    'AW 19/03/07    EP2_1897  -   End
    
    xmlNewRequest.appendChild xmlTempElem
    strResponse = objAppBO.CreateApplicationOffer(xmlNewRequest.xml)
    
    Set xmlCreateResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlCreateResponseDoc.documentElement, , True
                   
'Step:

    'BG Generate the received offer task or update an existing one with the correct applicationsequencenumber
    ' get current CASESTAGE detail

    Set xmlCaseTaskElem = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
   
    xmlCopyAttribute xmlCaseTaskElem, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttributeValue xmlApplicationElem, xmlElem, "APPLICATIONNUMBER", "CASEID"
    xmlCopyAttribute xmlCaseTaskElem, xmlElem, "ACTIVITYID"
    xmlThisRequestNode.appendChild xmlElem
    
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    xmlCopyAttribute xmlApplicationElem, xmlElem, "APPLICATIONPRIORITY"
    xmlThisRequestNode.appendChild xmlElem
    
    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
    xmlCurrentStageDoc.validateOnParse = False
    xmlCurrentStageDoc.setProperty "NewParser", True
    xmlCurrentStageDoc.async = False
    
    GetCurrentStage xmlThisRequestNode, xmlCurrentStageDoc
    
    strTMReceivedOffer = GetGlobalParamString("TMReceivedOffer")
    strValidationType = GetFirstComboValueId("TaskStatus", "I")
    
    strPattern = "RESPONSE/CASESTAGE/CASETASK[@TASKID='" & strTMReceivedOffer & "' and @TASKSTATUS='" & strValidationType & "']"
    Set xmlCaseTaskElem = xmlGetNode(xmlCurrentStageDoc, strPattern)
    Set xmlCaseStageElem = xmlGetNode(xmlCurrentStageDoc, "RESPONSE/CASESTAGE")
            
    If Not xmlCaseTaskElem Is Nothing Then
        
        'Call TaskManagementBO.UpdateCaseTask
        ' Build the Request and update the case task
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
        
        xmlSetAttributeValue xmlCaseTaskElem, "CONTEXT", xmlCreateResponseDoc.selectSingleNode("RESPONSE/OFFER/OFFERSEQUENCENO").Text
        xmlNewRequest.appendChild xmlCaseTaskElem
        
        strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
        errCheckXMLResponse strResponse, True
    
    End If
    
'Step:

    'Update LastOffered in Quotation
    
    strResponse = objQuotation.UpdateLastOfferedQuotation(xmlQuotationRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    ' RF 15/03/2006 MAR1392 Writing MOF message to ApplicationFirstTitle table is timing out:
    ' moved the "Automatically Approve Application" step to after the "First Title offer processing" step

'Step:

    ' First Title offer processing
    
    'MAR1259 If the global parameter is set to False, do not create the First Title tasks here.
    If (GetGlobalParamBoolean("TMFTCreateFTOfferInIssueOffer") = True) Then
    
        If IsValidationTypeInValidationList("TypeOfMortgage", "R", strTypeOfApplication) Then
        
            'MAR1533 check for taskid in combo groups
            'strTMIssueOffer = GetGlobalParamString("TMIssueOffer")
            'strTMReIssueOffer = GetGlobalParamString("TMReIssueOffer")
            strTMFTIssueOfferTaskID = GetGlobalParamString("TMFTIssueOfferTaskID")
            strTMReIssueOfferTaskID = GetGlobalParamString("TMFTReIssueOfferTaskID")    'MAR865 GHun
        
            strCurrentTaskId = xmlGetMandatoryAttributeText(xmlCaseTaskNode, "TASKID")
        
            'If strCurrentTaskId = strTMIssueOffer Then
            '    strCreateTaskID = strTMFTIssueOfferTaskID
            'ElseIf strCurrentTaskId = strTMReIssueOffer Then
            '    strCreateTaskID = strTMReIssueOfferTaskID
            'End If
            Dim ColValidIDsForIssueOffer As New Collection
            Dim ColValidIDsForReIssueOffer As New Collection
            GetValueIdsForValueName "TMIssueOfferTasks", strCurrentTaskId, ColValidIDsForIssueOffer
            GetValueIdsForValueName "TMReIssueOfferTasks", strCurrentTaskId, ColValidIDsForReIssueOffer
        
            If (ColValidIDsForIssueOffer.Count > 0) Then
                ' The task is in the Combo
                strCreateTaskID = strTMFTIssueOfferTaskID
            ElseIf (ColValidIDsForReIssueOffer.Count > 0) Then
                strCreateTaskID = strTMReIssueOfferTaskID
            End If
        
        
            'Generate a new task
            Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        
            xmlThisRequestDoc.validateOnParse = False
            xmlThisRequestDoc.setProperty "NewParser", True
        
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlThisRequestDoc.appendChild xmlNewRequest
        
            Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
            xmlNewRequest.appendChild xmlTMNewCaseTaskNode
            Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")    'MAR865 GHun
            xmlNewRequest.appendChild xmlApplicationNode.cloneNode(True)                    'MAR865 GHun
            xmlNewRequest.setAttribute "OPERATION", "CREATEADHOCCASETASK"
        
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
            xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strCreateTaskID
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
            xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "1"
            xmlSetSysDateToNodeAttrib xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME"
        
            CreateAdhocCaseTask xmlNewRequest
        End If
                
    End If
        
'Step:
    
    'Automatically Approve Application
    
    ' RF 15/03/2006 MAR1392 Writing MOF message to ApplicationFirstTitle table is timing out:
    ' moved this step to after the "First Title offer processing" step
    
    'MAR461 GHun
    'If blnTMAutoConditions Then
    If blnTMAutoApprove Then
    'MAR461 End
    
        strUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
        strUnitId = xmlGetAttributeText(vxmlRequestNode, "UNITID")
        
        Set xmlAppApproveRequestDoc = New FreeThreadedDOMDocument40
        
        Set xmlAppApproveRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "APPLICATIONFACTFIND")
        Set xmlAppNode = xmlAppApproveRequestDoc.selectSingleNode(".//APPLICATIONFACTFIND")
        
        Set xmlTempNode = xmlAppApproveRequestDoc.createElement("APPLICATIONAPPROVALDATE")
        xmlTempNode.Text = Format$(Now, "dd/mm/yyyy")
        xmlAppNode.appendChild xmlTempNode
        
        Set xmlTempNode = xmlAppApproveRequestDoc.createElement("APPLICATIONAPPROVALUSERID")
        xmlTempNode.Text = strUserId
        xmlAppNode.appendChild xmlTempNode
        
        Set xmlTempNode = xmlAppApproveRequestDoc.createElement("APPLICATIONAPPROVALUNITID")
        xmlTempNode.Text = strUnitId
        xmlAppNode.appendChild xmlTempNode
        
        strResponse = objAppBO.UpdateApplicationFactFind(xmlAppApproveRequestDoc.xml)
    
        Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    End If
    
    blnTMGenerateOfferExpiryTask = GetGlobalParamBoolean("TMGenerateOfferExpiryTask")
    
    If blnTMGenerateOfferExpiryTask Then
        
        Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
        
        strTMOfferExpiryTaskId = GetGlobalParamString("TMOfferExpiryTaskID")
        lngTMDaysToOfferExpiry = GetGlobalParamAmount("TMDaysToOfferExpiry")
    
        Set xmlOfferExpiryRequest = vxmlRequestNode.cloneNode(False)
    
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        xmlThisRequestDoc.appendChild xmlOfferExpiryRequest
        
        xmlSetAttributeValue xmlOfferExpiryRequest, "OPERATION", "FindCaseTaskList"
        
        Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
        
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMOfferExpiryTaskId
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
        xmlOfferExpiryRequest.appendChild xmlTMNewCaseTaskNode
    
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True
        
        On Error Resume Next
    
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        
        On Error GoTo IssueOfferVbErr
            
        lngErrorNumber = errCheckXMLResponseNode(xmlThisResponseDoc.documentElement, , False)
        
        Select Case True
            Case lngErrorNumber = 0
            
                Set xmlOfferExpiryCaseTask = xmlThisResponseDoc.selectSingleNode(".//CASETASK")
                Set xmlNewRequest = vxmlRequestNode.cloneNode(False)
                
                Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
                xmlThisRequestDoc.validateOnParse = False
                xmlThisRequestDoc.setProperty "NewParser", True
                xmlThisRequestDoc.appendChild xmlNewRequest
                
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "UpdateCaseTask"
                'MAR1599 GHun Any user needs to be able to update this task regardless of their authority level
                xmlSetAttributeValue xmlNewRequest, "USERAUTHORITYLEVEL", "99"
                
                Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
                xmlNewRequest.appendChild xmlTMNewCaseTaskNode
                
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "CASEID"
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "ACTIVITYID"
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "STAGEID"
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
                
                xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMOfferExpiryTaskId
                xmlCopyAttribute xmlOfferExpiryCaseTask, xmlTMNewCaseTaskNode, "TASKINSTANCE"
                
                dteDueDate = Now()
                dteDueDate = DateAdd("d", lngTMDaysToOfferExpiry, dteDueDate)
                xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME", Format(dteDueDate, "dd/mm/yyyy hh:mm:ss")
                
                Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
                xmlThisResponseDoc.validateOnParse = False
                xmlThisResponseDoc.setProperty "NewParser", True
                            
                RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
                errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
             
            Case errGetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound
            
                Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
                
                xmlThisRequestDoc.validateOnParse = False
                xmlThisRequestDoc.setProperty "NewParser", True
                Set xmlNewRequest = vxmlRequestNode.cloneNode(False)
                xmlThisRequestDoc.appendChild xmlNewRequest
                
                Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
                xmlNewRequest.appendChild xmlTMNewCaseTaskNode
                xmlNewRequest.appendChild xmlApplicationNode.cloneNode(False)
                xmlNewRequest.setAttribute "OPERATION", "CreateAdhocCaseTask"
                'MAR1599 GHun Any user needs to be able to create this task regardless of their authority level
                xmlNewRequest.setAttribute "USERAUTHORITYLEVEL", "99"
                
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
                xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
                xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMOfferExpiryTaskId
                xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
                
                dteDueDate = Now()
                dteDueDate = DateAdd("d", lngTMDaysToOfferExpiry, dteDueDate)
                xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME", Format(dteDueDate, "dd/mm/yyyy hh:mm:ss")
                
                CreateAdhocCaseTask xmlNewRequest
                        
            Case Else
                errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        End Select
    
    End If
    
    'MAR780 GHun The task must complete if we are not printing
    If blnPrintingEnabled = False Then
        Dim xmlTMUpdateCaseTaskRequest As IXMLDOMNode
        Dim xmlTMUpdateCaseTaskNode As IXMLDOMNode
            
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        
        Set xmlTMUpdateCaseTaskRequest = vxmlRequestNode.cloneNode(True)
    
        xmlThisRequestDoc.appendChild xmlTMUpdateCaseTaskRequest
    
        xmlSetAttributeValue xmlTMUpdateCaseTaskRequest, "OPERATION", "UpdateCaseTask"
            
        Set xmlTMUpdateCaseTaskNode = xmlGetMandatoryNode(xmlTMUpdateCaseTaskRequest, "CASETASK")
        xmlSetAttributeValue xmlTMUpdateCaseTaskNode, "TASKSTATUS", "40"
    
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        Set xmlThisRequestDoc = Nothing
    
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
        Set xmlThisResponseDoc = Nothing
        Set xmlTMUpdateCaseTaskRequest = Nothing
        Set xmlTMUpdateCaseTaskNode = Nothing
    End If
    'MAR780 End
    
WarningExit:

    gobjContext.SetComplete
                   
IssueOfferExit:
        
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTMCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlNode = Nothing
    Set xmlQuoteNumberNode = Nothing
    Set xmlQuotationNode = Nothing
    Set xmlNewRequest = Nothing
    Set xmlRemodelQuotationRequest = Nothing
    Set xmlTMNewCaseTaskNode = Nothing
    Set xmlTableNode = Nothing
    Set xmlApplicationElem = Nothing
    Set xmlTempReq = Nothing
    Set xmlTempChildElem = Nothing
    Set xmlCreateResponseDoc = Nothing
    Set xmlCaseTaskElem = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlCaseStageElem = Nothing
    Set xmlQuotationRequestDoc = Nothing
    Set xmlAppNode = Nothing
    Set xmlAppApproveRequestDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlPrintAttributeNode = Nothing
    Set xmlPrinterNode = Nothing
    Set xmlCaseTaskCopy = Nothing
    Set xmlOut = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlPanelLegalRepDetails = Nothing
    Set xmlAppLegalRepNode = Nothing
    
    Set objMsgTmBO = Nothing
    Set objAppBO = Nothing
    Set objAppQuoteBO = Nothing
    Set objQuotation = Nothing
    Set objAppProcBO = Nothing
    Set gobjEnquiry = Nothing
    Set objAppThirdPartyBO = Nothing
    
    ' PSC 26/10/2005 MAR300 - Start
    Set xmlTempRequest = Nothing
    Set xmlSearch = Nothing
    Set xmlCustomerList = Nothing
    Set xmlCustomersIn = Nothing
    Set xmlCustomerIn = Nothing
    Set xmlCustomer = Nothing
    Set xmlOfferExpiryRequest = Nothing
    Set xmlOfferExpiryCaseTask = Nothing
    ' PSC 26/10/2005 MAR300 - End
    
    'MAR1533
    Set ColValidIDsForIssueOffer = Nothing
    Set ColValidIDsForReIssueOffer = Nothing

    Set objNoTxBO = Nothing     'EP974/MAR1848

    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
    
    Exit Sub
    
' PSC 07/02/02 SYS3995
IssueOfferVbErr:

    gobjContext.SetAbort
    GoTo IssueOfferExit
End Sub
Public Function GetPanelLegalRepDetails(ByVal xmlRequestNode As IXMLDOMNode) As FreeThreadedDOMDocument40

    On Error GoTo GetPanelLegalRepDetailsVbErr

    Const strFunctionName As String = "GetPanelLegalRepDetails"
    
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    
    Dim strData As String
    Dim strAppNumber As String
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    
    'strAppNumber = xmlGetMandatoryNodeText(xmlRequestNode, "//APPLICATIONNUMBER")
    
    strAppNumber = xmlGetMandatoryAttributeText(xmlRequestNode.selectSingleNode("//APPLICATION"), "APPLICATIONNUMBER")
    
    Set adoParameter = adoCommand.CreateParameter("@pApplicationNumber", adBSTR, adParamInput, Len(strAppNumber), strAppNumber)
    adoCommand.Parameters.Append adoParameter
    
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.Open
    
    With adoCommand
        .CommandText = "USP_GETPANELLEGALREP"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "PANELLEGALREPLIST"
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = xmlLoad(strData, strFunctionName)

    Set GetPanelLegalRepDetails = xmlOut
    
 
GetPanelLegalRepDetailsExit:
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Function
    
GetPanelLegalRepDetailsVbErr:
    
    gobjContext.SetComplete
    
    GoTo GetPanelLegalRepDetailsExit
    
End Function
'
' JD BMIDS749 added.
' On issuing a KFI the application should be remodelled to pick up any base rate change
'
Private Sub RemodelKFI( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo RemodelKFIVbErr
    
    Const cstrFunctionName As String = "RemodelKFI"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlQuoteNumberNode As IXMLDOMElement
    Dim xmlQuotationNode As IXMLDOMElement
    
    Dim lngErrorNumber As Long
    Dim strResponse As String
    Dim blnPrintingEnabled As Boolean
    
    blnPrintingEnabled = False
    'SR 27/08/2004 : BMIDS815
    Dim strTMLTVRiskAssessTask  As String, strCondition As String
    Dim strApplicationNumber As String, strAFFNumber As String, strCaseActivityGuid As String
    Dim xmlTMNewCaseTaskNode As IXMLDOMElement, xmlTMCaseTaskNode As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim objBO As Object
    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    'SR 27/08/2004 : BMIDS815 - End
    
    ' check mandatory items '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    xmlCheckMandatoryAttribute vxmlRequestNode, "UNITID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    'SR 27/08/2004 : BMIDS815
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    'SR 27/08/2004 : BMIDS815 - End
    
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "TASKID"
        
    ' start of Application Data retrieval''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim objAppBO As Object
    
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)
    
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    strResponse = objAppBO.GetApplicationData(xmlThisRequestDoc.xml)
    
    Set xmlThisRequestDoc = Nothing
    
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    
    ' check for errors
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True

    Dim strActiveQuoteNumber As String
    strActiveQuoteNumber = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/ACTIVEQUOTENUMBER")

    If Len(strActiveQuoteNumber) = 0 Then
        
        errThrowError cstrFunctionName, 4818
        
    End If
    
    Set xmlThisResponseDoc = Nothing
    
    ' Get the appropriate Mortage subquote number from the quotation
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "QUOTATION")
    Set xmlQuotationNode = xmlGetNode(xmlThisRequestDoc, ".//QUOTATION")
    Set xmlNode = xmlThisRequestDoc.createElement("QUOTATIONNUMBER")
    xmlNode.Text = strActiveQuoteNumber
    xmlQuotationNode.appendChild xmlNode
    
    Dim objQuotation As Object
    Set objQuotation = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    strResponse = objQuotation.GetData(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True

    Dim strMortgageSubQuoteNo As String
    strMortgageSubQuoteNo = xmlGetNodeText(xmlThisResponseDoc, ".//MORTGAGESUBQUOTENUMBER")
    
    Set xmlQuotationNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlNode = Nothing
    Set xmlThisResponseDoc = Nothing
    
    ' Validate the current active quotation''''''''''''''''''''''''''''''''''''''''''''
    Dim objAppQuoteBO As Object
    Dim blnValidMortgageSubQuote As Boolean

    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "BASICQUOTATIONDETAILS")
           
    Set xmlApplicationNode = xmlGetNode(xmlThisRequestDoc.documentElement, "BASICQUOTATIONDETAILS")
    
    Set xmlNode = xmlThisRequestDoc.createElement("MORTGAGESUBQUOTENUMBER")
    If Not xmlNode Is Nothing Then
        xmlNode.Text = strMortgageSubQuoteNo
    End If
    
    If Not xmlApplicationNode Is Nothing Then
        xmlApplicationNode.appendChild xmlNode
    End If
    
    Set objAppQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    
    strResponse = objAppQuoteBO.ValidateMortgageSubQuote(xmlThisRequestDoc.xml)
    
    Set xmlThisRequestDoc = Nothing
    
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    blnValidMortgageSubQuote = xmlGetNodeAsBoolean(xmlThisResponseDoc.documentElement, "VALIDMORTGAGESUBQUOTE")
    
    If blnValidMortgageSubQuote = False Then
    
        errThrowError cstrFunctionName, 4815
    
    End If
    
    '  Remodel Quotation
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)
    Set xmlApplicationNode = xmlGetNode(xmlThisRequestDoc.documentElement, "APPLICATION")
    'Set xmlQuoteNumberNode = xmlGetNode(xmlApplicationNode, "ACTIVEQUOTENUMBER")
    
    'use active quote number as accepted quote number for the sake of RemodelQuotation
    Set xmlQuoteNumberNode = xmlThisRequestDoc.createElement("ACCEPTEDQUOTENUMBER")
    xmlApplicationNode.appendChild xmlQuoteNumberNode
    xmlQuoteNumberNode.Text = strActiveQuoteNumber
            
    strResponse = objAppQuoteBO.RemodelQuotation(xmlThisRequestDoc.xml)

    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    
    'SR 26/08/2004: BMIDS815
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
    Set xmlNode = xmlThisResponseDoc.selectSingleNode("//RESPONSE")
    If Not xmlNode Is Nothing Then
        If xmlGetAttributeText(xmlNode, "TYPE") = "WARNING" Then
            'Check the warning raised is 7020
            If errGetMessageText(7020) = xmlNode.selectSingleNode(".//MESSAGETEXT").Text Then
                '======================================================================
                'Check whether a task 'TMLTVRiskAssess', which is not actioned, exists
                '=======================================================================
                'First get CaseActivityGuid
                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETCASEACTIVITY"
                Set xmlTMCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
                Set xmlNode = xmlThisRequestDoc.createElement("CASEACTIVITY")
                xmlNewRequest.appendChild xmlNode
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "SOURCEAPPLICATION"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "CASEID"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "ACTIVITYID"
                xmlCopyAttribute xmlTMCaseTaskNode, xmlNode, "ACTIVITYINSTANCE"
                
                strResponse = objBO.TmRequest(xmlNewRequest.xml)
                Set xmlThisResponseDoc = xmlLoad(strResponse, TypeName(Me))
                errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
                Set xmlNode = xmlGetMandatoryNode(xmlThisResponseDoc.documentElement, "CASEACTIVITY")
                strCaseActivityGuid = xmlGetMandatoryAttributeText(xmlNode, "CASEACTIVITYGUID")
            
                'Check whether a task 'TMLTVRiskAssess', which is not actioned, exists
                strTMLTVRiskAssessTask = GetGlobalParamString("TMLTVRiskAssess")
        
                strCondition = "CASEACTIVITYGUID = " & "0x" & strCaseActivityGuid
                strCondition = strCondition & " AND STAGEID ='" _
                                & xmlGetMandatoryAttributeText(xmlCaseTaskNode, "STAGEID") & "'"
                strCondition = strCondition & " AND CASESTAGESEQUENCENO ='" _
                                & xmlGetMandatoryAttributeText(xmlCaseTaskNode, "CASESTAGESEQUENCENO") & "'"
                strCondition = strCondition & " AND TASKID ='" & strTMLTVRiskAssessTask & "'"
                strCondition = strCondition & " AND TASKSTATUS ='10'"
                If Not adoCheckRecordExists("CASETASK", strCondition) Then
                   'Generate a new task
                    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
                    xmlThisRequestDoc.validateOnParse = False
                    xmlThisRequestDoc.setProperty "NewParser", True
                    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                    xmlThisRequestDoc.appendChild xmlNewRequest
                    Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
                    xmlNewRequest.appendChild xmlTMNewCaseTaskNode
                    xmlNewRequest.setAttribute "OPERATION", "CREATECASETASK"
                    
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
                    xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMLTVRiskAssessTask
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
                    xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "1"
                    'BMIDS815 GHun
                    xmlSetSysDateToNodeAttrib xmlTMNewCaseTaskNode, "TASKDUEDATEANDTIME"
                    'BMIDS815 End
                    
                    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
                    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, vxmlResponseNode, True
                End If
                GoTo Warning7020Exit
            End If
        End If
    End If
    'SR 26/08/2004: BMIDS815 - End
    
    Set xmlQuoteNumberNode = Nothing
    Set objAppQuoteBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = Nothing

    
    ' Print KFI document'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strOutputDocument As String
    
    strOutputDocument = xmlGetAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT")
        
    If Len(strOutputDocument) > 0 Then
              
        '
        ' Printing the document
        '
        Dim strTaskName As String
        Dim strPrinterDestination As String
        Dim strApplicant As String, strCustomerIdentifier As String
        Dim xmlOut As New FreeThreadedDOMDocument40
        xmlOut.validateOnParse = False
        xmlOut.setProperty "NewParser", True
        Dim xmlThisCaseTaskDoc As New FreeThreadedDOMDocument40
        xmlThisCaseTaskDoc.validateOnParse = False
        xmlThisCaseTaskDoc.setProperty "NewParser", True
        Dim xmlThisCaseTaskElem As IXMLDOMElement
        Dim xmlThisCaseTaskNode As IXMLDOMNode
        Dim xmlCustomerNodeList As IXMLDOMNodeList
        Dim xmlCustomerNode As IXMLDOMNode
        Dim strTmpResponse As String
        Dim xmlTempNode As IXMLDOMNode
        Dim xmlTempElem As IXMLDOMElement
        Dim xmlPrintAttributeNode As IXMLDOMNode
        Dim xmlPrinterNode As IXMLDOMNode
        Dim xmlCaseTaskCopy As IXMLDOMNode
        Dim gobjEnquiry As Object
        
                blnPrintingEnabled = True
            
            ' Do the printing thing
                On Error Resume Next
                Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
                
                On Error GoTo RemodelKFIVbErr
    
                If Not gobjEnquiry Is Nothing Then
                
                    Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
                    Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
                    xmlThisCaseTaskElem.setAttribute "OPERATION", "GetPrintAttributes"
                              
                    Set xmlTempElem = xmlThisCaseTaskDoc.createElement("FINDATTRIBUTES")
                    Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
                    xmlTempElem.setAttribute "HOSTTEMPLATEID", strOutputDocument
                      
                    strResponse = gobjEnquiry.omRequest(xmlThisCaseTaskElem.xml)
                    Set xmlOut = xmlLoad(strResponse, cstrFunctionName)
                    '
                    ' Error Check
                    '
                    Set xmlTempNode = xmlOut.firstChild
                    errCheckXMLResponseNode xmlTempNode, , True
                    
                    Set xmlTempElem = xmlOut.selectSingleNode("RESPONSE/ATTRIBUTES")
                    
                    If xmlTempElem.getAttribute("INACTIVEINDICATOR") = "1" Then
                      errThrowError cstrFunctionName, 4820, "", "RemodelKFI"
                    End If
                    
                    If xmlTempElem.getAttribute("DEFAULTCOPIES") = "" Then
                      errThrowError cstrFunctionName, 4823, "", "RemodelKFI"
                    End If
                    
                    strPrinterDestination = xmlTempElem.getAttribute("PRINTERDESTINATIONTYPE")
            
                    ' EP2_1433 - use combovalidation types
                    ' strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
'                    If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
'                    And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0)) Then
'                       errThrowError cstrFunctionName, 4822, strPrinterDestination, "RemodelKFI"
'                    End If
                    ' IK - don't get this
                    If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
                    And IsValidationType("PrinterDestination", strPrinterDestination, "R") _
                    Then
                        errThrowError cstrFunctionName, 4822, GetComboText("PrinterDestination", CInt(strPrinterDestination)), "RemodelKFI"
                    End If
                    
                    Set xmlPrintAttributeNode = xmlTempElem
                    
                    Set xmlThisCaseTaskElem = Nothing
                    Set xmlThisCaseTaskNode = Nothing
                    Set xmlThisCaseTaskDoc = Nothing
                    Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
                    Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
                    xmlThisCaseTaskElem.setAttribute "OPERATION", "PrintDocumentForTask"
                    xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERID"
                    xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "UNITID"
                    xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERAUTHORITYLEVEL"
                    
                    Set xmlCaseTaskCopy = xmlCaseTaskNode.cloneNode(True)
                    xmlThisCaseTaskNode.appendChild xmlCaseTaskCopy
                    
                    Set xmlTempElem = xmlThisCaseTaskDoc.createElement("APPLICATION")
                    Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
                    'use the request applicationnode
                    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
                    
                    xmlCopyAttribute xmlApplicationNode, xmlTempNode, "APPLICATIONFACTFINDNUMBER"
                    
    
                    Set xmlCustomerNode = xmlApplicationNode.selectSingleNode("./CUSTOMER")
    '   Note these Attributes are dummies as the method in omPDM gets all the customer info
    '   again - they are needed here because omTM.PrintDocumentforTask wants them
                     xmlCopyAttribute xmlCustomerNode, xmlTempNode, "CUSTOMERVERSIONNUMBER"
                     
                    xmlCopyAttributeValue xmlCustomerNode, xmlCaseTaskCopy, "CUSTOMERNUMBER", "CUSTOMERIDENTIFIER"
                     
                    
                    xmlThisCaseTaskNode.appendChild xmlTempNode
                    
                    Set xmlTempElem = xmlThisCaseTaskDoc.createElement("PRINTATTRIBUTES")
                    Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
                    xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
                    xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DEFAULTCOPIES", "COPIES"
                    xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "PDMMETHOD", "METHODNAME"
                    xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
                    xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE"     'MAR7 GHun
                    
                    ' EP2_1433
                    ' xmlTempElem.setAttribute "DESTINATIONTYPE", Left(strPrinterDestination, 1)
                    xmlTempElem.setAttribute "DESTINATIONTYPE", GetValidationTypeForValueID("PrinterDestination", strPrinterDestination)
                    
                    xmlCopyAttributeValue xmlPrinterNode, xmlTempNode, "COMPRESSIONMETHOD", "COMPRESSIONMETHOD" 'MAR7 GHun
             
                     ' EP2_1433 - use combovalidation types
                    If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
                    Or IsValidationType("PrinterDestination", strPrinterDestination, "W") _
                    Then
                    ' If strPrinterDestination = "Local" Or strPrinterDestination = "Workstation Printer" Then    'MAR7 GHun
                        Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("PRINTER")
                        xmlTempElem.setAttribute "PRINTER", xmlGetAttributeText(xmlPrinterNode, "PRINTERNAME")
                    Else    '"Remote" Or "File"
                        xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "REMOTEPRINTERLOCATION", "PRINTER"
                    End If
                   
                    xmlThisCaseTaskNode.appendChild xmlTempNode
                    
                    PrintDocumentForTask xmlThisCaseTaskElem, vxmlResponseNode  'MAR7 GHun
                    
                End If
                Set xmlOut = Nothing
                Set xmlThisCaseTaskDoc = Nothing
                Set xmlThisCaseTaskElem = Nothing
                Set xmlThisCaseTaskNode = Nothing
                Set xmlCustomerNodeList = Nothing
                Set xmlCustomerNode = Nothing
                Set xmlTempNode = Nothing
                Set xmlTempElem = Nothing
                Set xmlPrintAttributeNode = Nothing
                Set xmlPrinterNode = Nothing
                Set gobjEnquiry = Nothing
                Set xmlCaseTaskCopy = Nothing
    End If
    
    
    ' Update task as complete''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' PrintDocumentForTask updates the task as complete so only do this if printing
    ' is not available
    If blnPrintingEnabled = False Then
        Dim xmlTMUpdateCaseTaskRequest As IXMLDOMNode
        Dim xmlTMUpdateCaseTaskNode As IXMLDOMNode
            
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        
        Set xmlTMUpdateCaseTaskRequest = vxmlRequestNode.cloneNode(True)
    
        xmlThisRequestDoc.appendChild xmlTMUpdateCaseTaskRequest
    
        xmlSetAttributeValue xmlTMUpdateCaseTaskRequest, "OPERATION", "UpdateCaseTask"
            
        Set xmlTMUpdateCaseTaskNode = xmlGetMandatoryNode(xmlTMUpdateCaseTaskRequest, "CASETASK") 'JD BMIDS838
        xmlSetAttributeValue xmlTMUpdateCaseTaskNode, "TASKSTATUS", "40" 'JD BMIDS838
    
        Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.setProperty "NewParser", True
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        Set xmlThisRequestDoc = Nothing
    
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        
        Set xmlThisResponseDoc = Nothing
        Set xmlTMUpdateCaseTaskRequest = Nothing
        Set xmlTMUpdateCaseTaskNode = Nothing
        Set xmlTMNewCaseTaskNode = Nothing 'SR 27/08/2004 : BMIDS815
        Set xmlTMCaseTaskNode = Nothing 'SR 27/08/2004 : BMIDS815
    End If

Warning7020Exit:
    gobjContext.SetComplete
                   
RemodelKFIExit:
    
    Set objAppBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RemodelKFIVbErr:

    gobjContext.SetAbort
    GoTo RemodelKFIExit

End Sub


Private Sub ApproveRecommendApplication(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Updates the application with the approval/recommended data and updates the case task
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo ApproveRecommendApplicationExit
    
    Const cstrFunctionName As String = "ApproveRecommendApplication"
    
    Dim objApplicationBO As Object
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As FreeThreadedDOMDocument40
    Dim xmlTaskManagerBO As Object
    Dim xmlAppRec As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
     
    ' validate REQUEST
    Set xmlTempRequest = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPROVREC", True, "APPLICATION")
    
    Dim strResponse As String
    
    ' Update the application data
    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    
    strResponse = objApplicationBO.ApproveRecommendApplication(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
        
    Set objApplicationBO = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTempRequest = Nothing

    ' Update the case task data
    Set xmlElement = vxmlRequestNode.cloneNode(True)
    xmlElement.setAttribute "OPERATION", "UpdateCaseTask"
    Set xmlAppRec = xmlGetMandatoryNode(xmlElement, "APPROVREC")
    xmlElement.removeChild xmlAppRec
    
    Set xmlTaskManagerBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    
    strResponse = xmlTaskManagerBO.TmRequest(xmlElement.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
        
    Set xmlTaskManagerBO = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlElement = Nothing
    
ApproveRecommendApplicationExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'BMIDS682 return info for Address Targeting
Private Sub RunCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    RunCreditCheck
' author:       Mark Coleman
' date:         22/08/2002
' description:  Runs a credit score, full bureau download and imports the bureau data into Omiga
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:   vxmlResponseNode
'       XML RESPONSE node
' notes:        The Credit Check and Bureau Download are run in a new transaction. This is so
'               that the Credit_Check task can be committed independently from the Import_Bureau_Data.
'               If the subsequent ImportBureauData process fails, an adhoc case task can be created
'               so that it can be re-run later.
'------------------------------------------------------------------------------------------

On Error GoTo RunCreditCheckExit
    
    Const cstrFunctionName  As String = "RunCreditCheck"
        
    Dim objNTxBO            As IOmTmNTxBO
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlTempRequest      As IXMLDOMElement

    'BMIDS658
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCustomerListNode As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim objNoTxBO           As omTmNoTxBO
    Dim strResponse         As String
    'BMIDS658 End
    'BMIDS682
    Dim xmlHeaderNode     As IXMLDOMNode
    Dim xmlTargetReqNode     As IXMLDOMNode
    Dim xmlCCN1Node     As IXMLDOMNode
    Dim bAddressTargeting As Boolean
    'BMIDS730
    Dim xmlNewRequest       As IXMLDOMNode
    Dim objMsgTmBO          As Object
  
    'Initialise ======================================================================
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
     'Validate ========================================================================
     Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
     Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    'BMIDS682 These will only be there for an address target request
     Set xmlHeaderNode = xmlGetNode(vxmlRequestNode, "REQHEADER")
     Set xmlTargetReqNode = xmlGetNode(vxmlRequestNode, "TARGETREQUEST")
     Set xmlCCN1Node = xmlGetNode(vxmlRequestNode, "CCN1LIST")
    
     'BMIDS658 Find the customers on the application outside of the transaction to reduce locking
     Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
     
     Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
     xmlTempRequest.setAttribute "OPERATION", "FINDCUSTOMERSFORCREDITCHECK"
     xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
     
     Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
     strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
     xmlTempDoc.loadXML strResponse
     errCheckXMLResponseNode xmlTempDoc.documentElement, , True
     Set xmlCustomerListNode = xmlGetMandatoryNode(xmlTempDoc.documentElement, "CUSTOMERLIST")
    
     Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
     xmlTempRequest.appendChild xmlCaseTaskNode.cloneNode(True)
     xmlTempRequest.appendChild xmlApplicationNode
     xmlTempRequest.appendChild xmlCustomerListNode
     'BMIDS658 End
    'BMIDS682 Add our Address Target Request if/when appropriate
    If Not xmlHeaderNode Is Nothing Then
        xmlTempRequest.appendChild xmlHeaderNode
    End If
    If Not xmlTargetReqNode Is Nothing Then
        xmlTempRequest.appendChild xmlTargetReqNode
    End If
    If Not xmlCCN1Node Is Nothing Then
        xmlTempRequest.appendChild xmlCCN1Node
    End If
    
    'Do Credit Score and Bureau Download =============================================
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNTxBO")
    'BMIDS682 Check the response
    strResponse = objNTxBO.RunExperianCreditCheck(xmlTempRequest)
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
   
    AttachResponseData vxmlResponseNode, xmlTempDoc.documentElement
     
    bAddressTargeting = xmlGetNodeAsBoolean(vxmlResponseNode, ".//ADDRESSTARGETING")
    
    'BMIDS682 If Address Targeting is Required Don't save at this time
    'Only ImportBureauData if we don't need to do addresstargeting
    If (bAddressTargeting = False) Then
    
        ' BMIDS730 If we've got this far, need to Update the Case Task status
        ' Build the Request and update the case task
        Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
        xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

        xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(xmlTempRequest, "USERID")
        xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(xmlTempRequest, "UNITID")
        xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
        xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
        xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)

        Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
        strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
        errCheckXMLResponse strResponse, True
    
        'Import Bureau data ==============================================================
        'Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
        xmlTempRequest.setAttribute "CREDITCHECK", "1"
        ImportBureauData xmlTempRequest
    End If
    
RunCreditCheckExit:
   
    Set objNTxBO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set objNoTxBO = Nothing
    'BMIDS730
    Set xmlNewRequest = Nothing
    Set objMsgTmBO = Nothing

    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'BMIDS758  Add new functions
Private Sub RunEnquiryCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                                  ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    RunEnquiryCreditCheck
' author:       Helen Aldred
' date:         27/07/2004
' description:  Runs an enquiry 'footprint' credit score, full bureau download and imports the bureau data into Omiga.
'
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       vxmlResponseNode
'               XML RESPONSE node
' notes:        The Credit Check and Bureau Download are run in a new transaction. This is so
'               that the Credit_Check task can be committed independently from the Import_Bureau_Data.
'               If the subsequent ImportBureauData process fails, an adhoc case task can be created
'               so that it can be re-run later.
'------------------------------------------------------------------------------------------

On Error GoTo RunEnquiryCreditCheckExit
    
    Const cstrFunctionName  As String = "RunEnquiryCreditCheck"
    
    'Set the EnquiryOnly attribute in the request.
    xmlSetAttributeValue vxmlRequestNode, "ENQUIRYCREDITCHECK", "TRUE"
  
    RunCreditCheck vxmlRequestNode, vxmlResponseNode
      
RunEnquiryCreditCheckExit:
  
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub RunUpgradeToFullAppCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                                           ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    RunUpgradeToFullAppCreditCheck
' author:       Helen Aldred
' date:         27/07/2004
' description:  Upgrades an enquiry 'footprint' credit score to a Full Application 'footprint'.
'
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       vxmlResponseNode
'               XML RESPONSE node
'------------------------------------------------------------------------------------------

On Error GoTo RunUpgradeToFullAppCreditCheckExit
    
    Const cstrFunctionName  As String = "RunUpgradeToFullAppCreditCheck"
        
    Dim objNTxBO            As IOmTmNTxBO
    Dim xmlTempDoc          As FreeThreadedDOMDocument40

    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlNewRequest       As IXMLDOMNode
    Dim objMsgTmBO          As Object
    
    Dim strResponse         As String
  
    'Initialise ======================================================================
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'Validate ========================================================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    
    'Upgrade Credit Score to a full application 'footprint' ==========================
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNTxBO")
    
    strResponse = objNTxBO.RunUpgradeToFullAppCreditCheck(vxmlRequestNode)
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
   
    AttachResponseData vxmlResponseNode, xmlTempDoc.documentElement
   
    ' If we've got this far, need to Update the Case Task status
    ' Build the Request and update the case task
    Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(xmlTempRequest, "USERID")
    xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(xmlTempRequest, "UNITID")
    xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
    xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
      
   
RunUpgradeToFullAppCreditCheckExit:
   
    Set objNTxBO = Nothing
    Set xmlTempDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNewRequest = Nothing
    Set objMsgTmBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub


' PSC 04/11/2005 MAR329
Private Sub RunRiskAssessment(ByVal vxmlRequestNode As IXMLDOMNode, Optional ByVal vxmlResponseNode = Nothing)

    On Error GoTo RunRiskAssessmentExit
    
    Const cstrFunctionName As String = "RunRiskAssessment"
    
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlCaseTakElem As IXMLDOMElement
    
    'SR 04/11/2005 : MAR269
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCurrentCaseStage As IXMLDOMElement
    Dim xmlRiskResponseDoc As FreeThreadedDOMDocument40 'MAR1173 GHun
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempElem2 As IXMLDOMElement
    Dim xmlInterfaceElem As IXMLDOMElement     ' MAR630
    Dim intFailLevel As Integer
    Dim intUWCount As Integer
    Dim intProcCount As Integer
    Dim intFraudCount As Integer
    'SR 04/11/2005 : MAR269 - End

    Dim objBO As Object
    
    Dim strApplicationNumber As String
    Dim strAFFNumber As String
    Dim strStageId As String 'SR 04/11/2005 : MAR269
    Dim strResponse As String
    Dim strInterfaceType As String          ' MAR630
    Dim strTaskId As String                 ' MAR630
    Dim strCaseAssessment As String         ' MAR630
    
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
    
    'SR 04/11/2005 : MAR269 - check whether StageId was passed in (as an attrib to Application node)
    '                if no, call GetCurrentStage
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    
    'MAR630
    strTaskId = ""
    
    If xmlGetAttributeText(vxmlRequestNode, "TASKMANAGER") <> "" Then
    
        Set xmlCaseActivity = xmlGetNode(vxmlRequestNode, ".//CASEACTIVITY")
        'MAR630 Use the CASETASK node and rename it.
        If xmlCaseActivity Is Nothing Then
            Set xmlCaseActivity = xmlGetNode(vxmlRequestNode, ".//CASETASK").cloneNode(True)
            xmlChangeNodeName xmlCaseActivity, "CASETASK", "CASEACTIVITY"
        End If
        
        strTaskId = xmlGetAttributeText(xmlCaseActivity, "TASKID")
        
        strStageId = xmlGetAttributeText(xmlApplicationElem, "STAGEID")
        If strStageId = "" Then
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlNewRequest.appendChild xmlCaseActivity.cloneNode(True)
            GetCurrentStage xmlNewRequest, xmlTempResponseDoc
    
            Set xmlCurrentCaseStage = xmlTempResponseDoc.selectSingleNode("//CURRENTCASESTAGE")
            strStageId = xmlGetAttributeText(xmlTempResponseDoc.documentElement, "STAGEID")
        End If
    
    End If
      
    'Create Request and run RiskAssessment
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)

    If xmlGetAttributeText(xmlApplicationElem, "STAGEID") = "" Then
        xmlSetAttributeValue xmlApplicationElem, "STAGEID", strStageId
    End If
    'SR 04/11/2005 : MAR269 - End
    
    ' PSC 20/11/01 SYS2867/SYS2988
    xmlNewRequest.appendChild xmlMakeNodeElementBased(xmlApplicationElem, False, "RISKASSESSMENT")
    
    Set objBO = gobjContext.CreateInstance("omRA.RiskAssessmentBO")
    strResponse = objBO.RunRiskAssessment(xmlNewRequest.xml)
    
    ' ik_MAR452_20051111
    ' xmlLoad strResponse, xmlTempResponseDoc
    Set xmlRiskResponseDoc = xmlLoad(strResponse, cstrFunctionName) 'MAR1173 GHun
    ' ik_MAR452_20051111_ends
    
    Dim lngErrNo As Long
    lngErrNo = errCheckXMLResponseNode(xmlRiskResponseDoc.documentElement, , False) 'MAR1173 GHun
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, 4813
    Else   'SR 04/11/2005 : MAR269
        'MAR1173 GHun
        intFailLevel = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "FAILLEVEL")
        intUWCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "UNDERWRITERCOUNT")
        intProcCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "PROCESSORCOUNT")
        intFraudCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "FRAUDCOUNT") 'JD MAR1615
        'MAR1173 End
        'MAR1142 DRC Allow web apps to RA create tasks
        ' If xmlGetAttributeText(vxmlRequestNode, "TASKMANAGER") <> "" And intFailLevel > 1 Then
        'MAR1142 DRC
            If intFailLevel > 1 Then
            'Call HandleInterfaceResponse here
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "HANDLEINTERFACERESPONSE"
            xmlNewRequest.appendChild xmlApplicationElem.cloneNode(True)
            'MAR 1142 DRC - Case Activity can be created in Handle Interface
            If Not xmlCaseActivity Is Nothing Then
                xmlNewRequest.appendChild xmlCaseActivity.cloneNode(True)
            End If
            'MAR 1142 DRC
            'MAR630  Create the MESSAGESUBTYPELIST node under INTERFACE
            Set xmlInterfaceElem = xmlTempResponseDoc.createElement("INTERFACE")
            
            'MAR630 Get the value ID for interface type Case Assessement "CA"
            strInterfaceType = GetFirstComboValueId("InterfaceType", "CA")
            xmlSetAttributeValue xmlInterfaceElem, "INTERFACETYPE", strInterfaceType
             
            xmlSetAttributeValue xmlInterfaceElem, "MESSAGETYPE", intFailLevel
            
            xmlSetAttributeValue xmlInterfaceElem, "CREATETASKFLAG", "1"
            'MAR 1142 DRC - No messagesubtypes for decline tasks
            If intFailLevel < 3 Then
             If intUWCount > 0 Or intProcCount > 0 Or intFraudCount > 0 Then
                 Set xmlTempElem2 = xmlTempResponseDoc.createElement("MESSAGESUBTYPELIST")
                 xmlSetAttributeValue xmlTempElem2, "CREATETASKFLAG", 1
                 xmlInterfaceElem.appendChild xmlTempElem2
             End If
                         
             If intUWCount > 0 Then
                 Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                 xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "10"
                 xmlTempElem2.appendChild xmlTempElem
             End If
             
             If intProcCount > 0 Then
                 Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                 xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "20"
                 xmlTempElem2.appendChild xmlTempElem
             End If
             
             If intFraudCount > 0 Then
                 Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                 xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "30"
                 xmlTempElem2.appendChild xmlTempElem
             End If
             End If 'MAR 1142 DRC
             xmlNewRequest.appendChild xmlInterfaceElem
 
            
            strResponse = OmTmRequest(xmlNewRequest.xml)
            Set xmlTempResponseDoc = xmlLoad(strResponse, cstrFunctionName)
            errCheckXMLResponseNode xmlTempResponseDoc.documentElement, , True 'MAR1143 GHun Raise the error
        End If

        ' Now update Case Task if task = Case Assessment
        
        strCaseAssessment = GetMandatoryGlobalParamString("TMRiskAssess")
        If (Not vxmlRequestNode.selectSingleNode(".//CASETASK") Is Nothing) And (strTaskId = strCaseAssessment) Then
            
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
            
            Set xmlCaseTakElem = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
            xmlSetAttributeValue xmlCaseTakElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
            xmlSetAttributeValue xmlCaseTakElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
            xmlSetAttributeValue xmlCaseTakElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
            xmlSetAttributeValue xmlCaseTakElem, "TASKSTATUS", 40
            
            xmlNewRequest.appendChild xmlCaseTakElem
            
            Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
            strResponse = objBO.TmRequest(xmlNewRequest.xml)
            
            errCheckXMLResponse strResponse, True
        End If
        
        ' PSC 04/11/2005 MAR329- Start
        ' Send the response back if required
        If Not vxmlResponseNode Is Nothing Then
        
            Dim xmlRAChild As IXMLDOMNode
            
            For Each xmlRAChild In xmlRiskResponseDoc.documentElement.childNodes    'MAR1173 GHun
                ' ik_MAR452_20051111
                vxmlResponseNode.appendChild xmlRAChild.cloneNode(True)
            Next
            
            ' ik_MAR452_20051111
            Set xmlRAChild = Nothing
        End If
        ' PSC 04/11/2005 MAR329 - End

    End If 'SR 04/11/2005 : MAR269 - End

RunRiskAssessmentExit:
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTakElem = Nothing
    Set objBO = Nothing
    'SR 04/11/2005 : MAR269
    Set xmlCaseActivity = Nothing
    Set xmlCurrentCaseStage = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlRiskResponseDoc = Nothing    'MAR1173 GHun
    Set xmlTempElem = Nothing
    Set xmlTempElem2 = Nothing
    'SR 04/11/2005 : MAR269 - End
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
Private Sub ValidateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ValidateValuationReportExit
    
    Const cstrFunctionName As String = "ValidateValuationReport"
    
    Dim strValResponse As String
    Dim xmlCaseNode  As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlCaseTaskReq As IXMLDOMElement
    Dim xmlCaseResp As IXMLDOMNode
    Dim xmlValRequest As IXMLDOMElement
    Dim xmlValuation As IXMLDOMElement
    Dim xmlReqValuation As IXMLDOMElement
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDOMCaseResp As FreeThreadedDOMDocument40
    Dim xmlDOMValuationResp As FreeThreadedDOMDocument40
    Dim strReturnNodeName1 As String
    Dim strReturnNodeName2 As String
    
    Dim objAppProcBO As Object
    Dim objValRulesBO As Object
    Dim objRB As Object
    
    ' Setup default response block
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Set xmlResponseNode = xmlThisResponseDoc.appendChild(xmlThisResponseDoc.createElement("RESPONSE"))
    
    ' Find Valuation details in request
    If vxmlRequestNode.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequestNode
    Else
        Set xmlRequestNode = xmlGetMandatoryNode(vxmlRequestNode, ".//REQUEST")
    End If
    
    Set xmlValuation = xmlGetMandatoryNode(xmlRequestNode, "./VALUATION")

    ''''''''''''''''''''''''''''
    ' Do Validation
    
    ' Create Validation Request block
    Set xmlValRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlValRequest.setAttribute "OPERATION", "ValidateValuationReport"

    ' Append the Valuation details passed in
    xmlValRequest.appendChild xmlValuation.cloneNode(True)
    
    ' Make the call
    
    'MAR1434 switch between different rule calls
    If GetGlobalParamBoolean("ValuationUseAppProcRules") = True Then
        Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
        
        If objAppProcBO Is Nothing Then
            errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
        End If
        
        strReturnNodeName1 = ".//APPSTATUS"
        strReturnNodeName2 = ".//LTV"
        xmlThisResponseDoc.loadXML objAppProcBO.OmAppProcRequest(xmlValRequest.xml)
        Set objAppProcBO = Nothing
    Else
        
        strReturnNodeName1 = ".//VALUATION"
        strReturnNodeName2 = ".//TASKS"
        'Need to call omRB to get all data required for ValRules
        Dim xmlRBDoc As FreeThreadedDOMDocument40
        Set xmlRBDoc = New FreeThreadedDOMDocument40
        Dim xmlRBReqNode As IXMLDOMElement
        Dim xmlRBAppNode As IXMLDOMElement
        Dim xmlRbResponseNode As IXMLDOMNode
        Dim xmlChildNode As IXMLDOMNode
        Dim strRetXML As String
        Set xmlRBReqNode = xmlGetRequestNode(vxmlRequestNode)
        
        xmlRBReqNode.setAttribute "COMBOLOOKUP", "NO"
        xmlRBReqNode.setAttribute "RB_TEMPLATE", "APValnRBTemplate"
        Set xmlRBAppNode = xmlRBDoc.createElement("APPLICATION")
        xmlRBReqNode.appendChild xmlRBAppNode
        xmlRBAppNode.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlValuation, "APPLICATIONNUMBER")
        xmlRBAppNode.setAttribute "APPLICATIONFACTFINDNUMBER", xmlGetAttributeText(xmlValuation, "APPLICATIONFACTFINDNUMBER")
        xmlRBAppNode.setAttribute "_SCHEMA_", "APPLICATION"
        
        Set objRB = gobjContext.CreateInstance("omRB.omRequestDO")
        If objRB Is Nothing Then
            errThrowError cstrFunctionName, oeObjectNotCreatable, "omRB.omRequestDO"
        End If
        strRetXML = objRB.OmDataRequest(xmlRBReqNode.xml)
        Set objRB = Nothing
        
        xmlRBDoc.loadXML strRetXML
        Set xmlRbResponseNode = xmlRBDoc.selectSingleNode("RESPONSE")
        'Append RB response to valuation rules request
        For Each xmlChildNode In xmlRbResponseNode.childNodes
            xmlValRequest.appendChild xmlChildNode.cloneNode(True)
        Next
        
        Set objValRulesBO = gobjContext.CreateInstance("omValuationRules.ValuationRulesBO")
        
        If objValRulesBO Is Nothing Then
            errThrowError cstrFunctionName, oeObjectNotCreatable, "omValuationRules.ValuationRulesBO"
        End If
        
        xmlValRequest.setAttribute "OPERATION", "ProcessValuationReport"
        xmlThisResponseDoc.loadXML objValRulesBO.RunValuationRules(xmlValRequest.xml)
        Set objValRulesBO = Nothing
    End If
   
   ' parse response
    Call errCheckXMLResponseNode(xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' AQR 2271 DRC - add response from valuation call to this response node
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlAppStatusNode As IXMLDOMNode
    Dim xmlLTVNode As IXMLDOMNode
    'Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(strReturnNodeName1)
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    'Set xmlLTVNode = xmlThisResponseDoc.selectSingleNode(".//LTV")
    Set xmlLTVNode = xmlThisResponseDoc.selectSingleNode(strReturnNodeName2)
    If Not xmlLTVNode Is Nothing Then
        vxmlResponseNode.appendChild xmlLTVNode
    End If
      
    
    ''''''''''''''''''''''''''''
    
    Set xmlCaseNode = xmlGetMandatoryNode(xmlRequestNode, "./CASETASK")

    ' Update Case Details if previous validation is successful
    Set xmlCaseTaskReq = xmlGetRequestNode(vxmlRequestNode)
    Set xmlReqValuation = xmlCaseNode.cloneNode(True)
    xmlReqValuation.setAttribute "TASKSTATUS", cstrTaskComplete
    xmlCaseTaskReq.appendChild xmlReqValuation
    
    xmlCaseTaskReq.setAttribute "OPERATION", "UpdateCaseTask"
    Set xmlDOMCaseResp = New FreeThreadedDOMDocument40
    xmlDOMCaseResp.validateOnParse = False
    xmlDOMCaseResp.setProperty "NewParser", True
    
    xmlDOMCaseResp.loadXML gobjTmBO.TmRequest(xmlCaseTaskReq.xml)
    errCheckXMLResponseNode xmlDOMCaseResp.selectSingleNode("RESPONSE"), , True
    
    ' Append to the response block
    AttachResponseData xmlDOMCaseResp.documentElement, xmlResponseNode

ValidateValuationReportExit:
    Set xmlCaseTaskReq = Nothing
    Set xmlResponseNode = Nothing
    Set xmlAppStatusNode = Nothing
    Set xmlLTVNode = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub CreateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode)
    On Error GoTo CreateValuationReportExit
    Dim strValResponse As String
    Dim xmlOperationNode As IXMLDOMElement
    Dim xmlValElem As IXMLDOMElement
    Dim xmlValRequest As IXMLDOMElement
    Dim objOmTmBo As Object
    Dim objAppProcBO As Object
    
    Const cstrFunctionName As String = "CreateValuationReport"
    
    ' Create the Valuation Report
    Set xmlValElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION")
    Set xmlValRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlValRequest.setAttribute "OPERATION", "CreateValuationReport"
    xmlValRequest.appendChild xmlValElem.cloneNode(True)
    
    ' Do the Call
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
    
    If objAppProcBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    strValResponse = objAppProcBO.OmAppProcRequest(xmlValRequest.xml)
    Set objAppProcBO = Nothing
    
    ' Check the response
    errCheckXMLResponse strValResponse, True

    ' Now Update the case task
    Dim xmlCaseReq As IXMLDOMElement
    Dim xmlTableNode As IXMLDOMNode
    Dim strCaseResp As String
    
    Set xmlCaseReq = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlCaseReq, "OPERATION", "UPDATECASETASK"
    xmlCaseReq.appendChild vxmlRequestNode.selectSingleNode(".//CASETASK")
    Set xmlTableNode = xmlCaseReq.selectSingleNode(".//CASETASK")
    xmlSetAttributeValue xmlTableNode, "TASKSTATUS", 20
    
    Set objOmTmBo = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    strCaseResp = objOmTmBo.TmRequest(xmlCaseReq.xml)
    
    ' Check the response
    errCheckXMLResponse strCaseResp, True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateValuationReportExit:
    
    Set objOmTmBo = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCurrEmployersRefErr:

    gobjContext.SetAbort
    
    GoTo CreateValuationReportExit
    
End Sub

Private Sub CreateValuationReportNoInst(ByVal vxmlRequestNode As IXMLDOMNode)
'This routine has been written to create a valuation report which has not been
'created as a result of a valuer instruction.

    On Error GoTo CreateValuationReportNoInstErr
    
    'Declare Variables
    Dim xmlExtractDOM As FreeThreadedDOMDocument40
    Dim xmlValuerResponse As FreeThreadedDOMDocument40
    Dim xmlValuerRequest As FreeThreadedDOMDocument40
    Dim xmlComboRequest As FreeThreadedDOMDocument40
    Dim xmlComboResponse As FreeThreadedDOMDocument40
    Set xmlExtractDOM = New FreeThreadedDOMDocument40
    xmlExtractDOM.validateOnParse = False
    xmlExtractDOM.setProperty "NewParser", True
    Set xmlValuerRequest = New FreeThreadedDOMDocument40
    xmlValuerRequest.validateOnParse = False
    xmlValuerRequest.setProperty "NewParser", True
    Set xmlValuerResponse = New FreeThreadedDOMDocument40
    xmlValuerResponse.validateOnParse = False
    xmlValuerResponse.setProperty "NewParser", True
    Set xmlComboRequest = New FreeThreadedDOMDocument40
    xmlComboRequest.validateOnParse = False
    xmlComboRequest.setProperty "NewParser", True
    Set xmlComboResponse = New FreeThreadedDOMDocument40
    xmlComboResponse.validateOnParse = False
    xmlComboResponse.setProperty "NewParser", True
    
    Dim xmlValElem As IXMLDOMElement
    Dim xmlInstElem As IXMLDOMElement
    Dim xmlValRequest As IXMLDOMElement
    Dim xmlInstRequest As IXMLDOMElement
    Dim xmlCaseReq As IXMLDOMElement
    Dim xmlValuerReqElem As IXMLDOMElement
    Dim xmlValuerList As IXMLDOMElement
    Dim xmlValuerPanelID As IXMLDOMElement
    Dim xmlValuerValType As IXMLDOMElement
    Dim xmlList As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlOldValNode As IXMLDOMNode 'MAR1396
        
    Dim objOmTmBo As Object
    Dim objAppProcBO As Object
    Dim objThirdPartyBO As Object
    Dim objComboDO As Object
    
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlValuationType As IXMLDOMNode
         
    Dim strCaseResp As String
    Dim strReportResponse As String
    Dim strInstResponse As String
    Dim intSeqNo As Integer
    Dim strValuerResponse As String
    Dim strDirectoryGUID As String
    Dim strComboList As String
    Dim strStandardValuer As String
    
    Const cstrFunctionName As String = "CreateValuationReportNoInst"
    
    'amend request node
    xmlChangeNodeName vxmlRequestNode, "APPLICATION", "VALUATION"
        
    'set up connection to Application Processing & Third Party Business Objects
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
    
    'DPF 06/11/2002 - BMIDS00852
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        
    If objThirdPartyBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO"
    End If
    
    If objAppProcBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    'DPF 06/11/2002 - BMIDS00852
    If objComboDO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrBASE_COMPONENT & ".ComboDO"
    End If
      
    'DPF 5/11/2002 - BMIDS00852 - use Combo Assist to pull back valuer id for Validation type 'L'
    strStandardValuer = "L"
    Set xmlList = xmlComboRequest.createElement("LIST")
    xmlComboRequest.appendChild xmlList
    Set xmlNode = xmlComboRequest.createElement("LISTNAME")
    xmlNode.Text = "ValuationType"
    xmlList.appendChild xmlNode
    
    strComboList = objComboDO.GetComboList(xmlList.xml)
    Set xmlComboResponse = xmlLoad(strComboList, "CreateValuationReportNoInst")
        
    Set xmlValuationType = xmlComboResponse.selectSingleNode("//LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE = '" & strStandardValuer & "']")
    'DPF - END OF BMIDS00852
      
    'PB 29/06/2006 EP435 Begin
    'Retrieve the Directory GUID of a default Valuer set up in supervisor
    'Set xmlValuerReqElem = xmlGetRequestNode(vxmlRequestNode)
    'xmlValuerReqElem.setAttribute "OPERATION", "FindPanelValuerList"
    'Set xmlValuerList = xmlValuerRequest.createElement("VA_PANELVALUERLIST")
    'Set xmlValuerValType = xmlValuerRequest.createElement("VALUATIONTYPE")
    'Set xmlValuerPanelID = xmlValuerRequest.createElement("PANELID")
    'xmlValuerPanelID.Text = "Default"
    'xmlValuerValType.Text = xmlValuationType.selectSingleNode("VALUEID").Text 'DPF 05/11/2002 - BMIDS00852
    'xmlValuerList.appendChild xmlValuerValType
    'xmlValuerList.appendChild xmlValuerPanelID
    'xmlValuerReqElem.appendChild xmlValuerList
    'xmlValuerRequest.appendChild xmlValuerReqElem
    'PB EP435 End
    
    'strValuerResponse = objThirdPartyBO.FindPanelValuerList(xmlValuerRequest.xml)
    
    'Check Response
    'errCheckXMLResponse strValuerResponse, True
        
    'Assign directory guid retrieved into Variable
    'xmlValuerResponse.loadXML (strValuerResponse)
    'AW     20/12/2002  BM0190
    'strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "VA_PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
    'strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
        
    'MAR1396 - find values from previous valuerinstruction if present to pass on to new one
    Set xmlInstElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION").cloneNode(True)
    xmlChangeNodeName xmlInstElem, "VALUATION", "VALUERINSTRUCTION"
    xmlInstElem.setAttribute "_ORDERBY_", "INSTRUCTIONSEQUENCENO DESC"
    Set xmlInstRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlInstRequest.setAttribute "OPERATION", "FindValuerInstructionList"
    xmlInstRequest.appendChild xmlInstElem.cloneNode(True)
    
    strInstResponse = objAppProcBO.OmAppProcRequest(xmlInstRequest.xml)
    Set xmlExtractDOM = xmlLoad(strInstResponse, "FindValuerInstructionList")
    'The first in the list will be the latest as we did an ORDERBY
    Set xmlOldValNode = xmlExtractDOM.documentElement.selectNodes("VALUERINSTRUCTION").Item(0)
    
    
    'Create 'Dummy' Valuer Instruction record
    Set xmlInstElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION").cloneNode(True)
    xmlChangeNodeName xmlInstElem, "VALUATION", "VALUERINSTRUCTION"
    'PB 29/06/2006 EP435 Begin
    'xmlInstElem.setAttribute "VALUERPANELNO", "Default"
    'PB EP435 End
    xmlInstElem.setAttribute "DIRECTORYGUID", strDirectoryGUID
    'MAR1396 if we have a previous valuerinstruction then copy some values across
    If Not xmlOldValNode Is Nothing Then
        xmlInstElem.setAttribute "VALUATIONTYPE", xmlGetAttributeText(xmlOldValNode, "VALUATIONTYPE")
        xmlInstElem.setAttribute "APPOINTMENTDATE", xmlGetAttributeText(xmlOldValNode, "APPOINTMENTDATE")
        xmlInstElem.setAttribute "DATEOFINSTRUCTION", xmlGetAttributeText(xmlOldValNode, "DATEOFINSTRUCTION")
        'AW EP1196  - Copy valuer details
        xmlInstElem.setAttribute "DIRECTORYGUID", xmlGetAttributeText(xmlOldValNode, "DIRECTORYGUID")
        xmlInstElem.setAttribute "VALUATIONSTATUS", xmlGetAttributeText(xmlOldValNode, "VALUATIONSTATUS")
        xmlInstElem.setAttribute "VALUERPANELNO", xmlGetAttributeText(xmlOldValNode, "VALUERPANELNO")
    End If
    Set xmlInstRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlInstRequest.setAttribute "OPERATION", "CreateValuerInstructions"
    xmlInstRequest.appendChild xmlInstElem.cloneNode(True)
    
    strInstResponse = objAppProcBO.OmAppProcRequest(xmlInstRequest.xml)
    
    'Check Response
    errCheckXMLResponse strInstResponse, True
    
    'Extract Instruction Sequence No of record created
    Set xmlExtractDOM = xmlLoad(strInstResponse, "CreateValuationReportNoInst")
    intSeqNo = xmlExtractDOM.documentElement.Attributes.getNamedItem("INSTRUCTIONSEQUENCENO").Text
              
    '-----------------------------------------------------------------------------
        
    'Call omAppproc.ValuationBO.CreateValuationReportUsingTemplate
    'passing in initial request data & Instruction Seq No of record just created
    Set xmlValElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION")
    xmlValElem.setAttribute "INSTRUCTIONSEQUENCENO", intSeqNo
    Set xmlValRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlValRequest.setAttribute "OPERATION", "CreateValuationReportUsingTemplate"
    xmlValRequest.appendChild xmlValElem.cloneNode(True)
            
    strReportResponse = objAppProcBO.OmAppProcRequest(xmlValRequest.xml)
    Set objAppProcBO = Nothing
        
    'Check Response
    errCheckXMLResponse strReportResponse, True
          
    '-----------------------------------------------------------------------------
      
    'Update Case Task
    Set xmlCaseReq = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlCaseReq, "OPERATION", "UPDATECASETASK"
    xmlCaseReq.appendChild vxmlRequestNode.selectSingleNode(".//CASETASK")
    Set xmlTableNode = xmlCaseReq.selectSingleNode("//CASETASK")
    xmlSetAttributeValue xmlTableNode, "TASKSTATUS", 20
    xmlSetAttributeValue xmlTableNode, "CONTEXT", intSeqNo
    
    Set objOmTmBo = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    strCaseResp = objOmTmBo.TmRequest(xmlCaseReq.xml)
    
    'Check Response
    errCheckXMLResponse strCaseResp, True
        
    'Complete Transaction
    gobjContext.SetComplete
        
CreateValuationReportNoInstExit:
    
    Set xmlExtractDOM = Nothing
    Set xmlValuerRequest = Nothing
    Set xmlValuerResponse = Nothing
    Set xmlComboRequest = Nothing
    Set xmlComboResponse = Nothing
    
    Set xmlValElem = Nothing
    Set xmlInstElem = Nothing
    Set xmlValRequest = Nothing
    Set xmlInstRequest = Nothing
    Set xmlCaseReq = Nothing
    Set xmlValuerReqElem = Nothing
    Set xmlValuerList = Nothing
    Set xmlValuerPanelID = Nothing
    Set xmlValuerValType = Nothing
    Set xmlList = Nothing
    Set xmlNode = Nothing
        
    Set objOmTmBo = Nothing
    Set objThirdPartyBO = Nothing
    Set objComboDO = Nothing
    
    Set xmlTableNode = Nothing
    Set xmlValuationType = Nothing
            
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub

CreateValuationReportNoInstErr:

    gobjContext.SetAbort
    MsgBox "Source:  " & Err.Source & " Description:  " & Err.Description & " Number:  " & Err.Number
    GoTo CreateValuationReportNoInstExit
    
End Sub

Private Sub CompleteValuerInstructions(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
                  
' header ----------------------------------------------------------------------------------
' description:
'   Updates the ststus of the "Assign Valuer" task to complete and creates an
'   "Instruct Valuer" task
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CompleteValuerInstructionsExit
    
    Const strFunctionName As String = "CompleteValuerInstructions"
    
    Dim xmlResponseDoc As FreeThreadedDOMDocument40, xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMElement
    Dim objTaskManagerBO As Object
    Dim objAppProcBO As Object
    Dim objComboDO As Object
    Dim xmlCaseTask As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCaseTaskList As IXMLDOMNodeList
    Dim xmlStageTask As IXMLDOMElement
    Dim xmlTask As IXMLDOMElement, xmlCaseActivity As IXMLDOMElement
    Dim xmlAttribute As IXMLDOMAttribute
    Dim xmlApplication As IXMLDOMElement
    Dim xmlValuerInstruction As IXMLDOMElement
    
    Dim strAssignValuerId As String, strInstructValuerId As String, strResponse As String
    Dim intCaseStageSeqNo As Integer, lngErrNo As Long, strCondition As String, strTaskStatus As String
    Dim intTaskInstance As Integer, intTemp As Integer
     
    ' validate REQUEST
    xmlCheckMandatoryNode vxmlRequestNode, "CASETASK"
    
    'Get the Ids for the tasks from Global Parameters
    strAssignValuerId = GetMandatoryGlobalParamString("TMAssignValuer")
    strInstructValuerId = GetMandatoryGlobalParamString("TMInstructValuer")
    
    Set objTaskManagerBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    Set objAppProcBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    
    'Get Current stage details
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    Set xmlCaseTask = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETCURRENTSTAGE"
    Set xmlCaseActivity = xmlResponseDoc.createElement("CASEACTIVITY")
    
    xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "CASEID"
    xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "ACTIVITYID"
    
    xmlTempRequest.appendChild xmlCaseActivity
    
    strResponse = objTaskManagerBO.TmRequest(xmlTempRequest.xml)
    Set xmlCurrentStageDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlCurrentStageDoc.documentElement, vxmlResponseNode, True
    
    Set xmlNode = xmlCurrentStageDoc.selectSingleNode("//CASESTAGE")
    intCaseStageSeqNo = xmlGetAttributeText(xmlNode, "CASESTAGESEQUENCENO")
    
    ' Check whether any records with AssignValuer task exist
    strCondition = "//CASETASK[@TASKID='" & strAssignValuerId & "']"
    Set xmlCaseTaskList = xmlCurrentStageDoc.selectNodes(strCondition)
    
    Dim blnCreateTask As Boolean
    
    If xmlCaseTaskList.length >= 1 Then
        intTaskInstance = 0
        intTemp = 0
        ' If there are more than one CaseTask nodes, take the one with the highest taskInstance value
        For Each xmlNode In xmlCaseTaskList
            intTemp = xmlGetAttributeText(xmlNode, "TASKINSTANCE")
            If intTemp > intTaskInstance Then
                Set xmlCaseTask = xmlNode
                intTaskInstance = intTemp
                strTaskStatus = xmlGetAttributeText(xmlNode, "TASKSTATUS")
            End If
        Next xmlNode
        
        strTaskStatus = xmlGetAttributeText(xmlCaseTask, "TASKSTATUS")
        'if the task status is 10, update the task status as pending
         'GD SYS 2092
         If strTaskStatus = "10" Or strTaskStatus = "20" Then
            Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
            xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", omiga4TASKSTATUSComplete
            xmlTempRequest.appendChild xmlCaseTask
            
            strResponse = objTaskManagerBO.TmRequest(xmlTempRequest.xml)
            Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
        'GD SYS 2092
        'Else
            
            'If strTaskStatus = 20 Then
                ' Exit process
            '    GoTo CompleteValuerInstrExit2
            'Else
            '    blnCreateTask = True
            'End If
        End If
    Else ' No tasks with taskid 'AssignValuer' exist
        blnCreateTask = True
    End If
    
    If blnCreateTask Then
        'Get the details of the Assign Valuer task in order for it to be created
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "FINDSTAGETASKLIST"
        Set xmlStageTask = xmlResponseDoc.createElement("STAGETASK")
        xmlCopyAttribute xmlCaseTask, xmlStageTask, "STAGEID"
        'SR 25/04/01 : SYS2194 - Get the details of AssignValuer task
        xmlSetAttributeValue xmlStageTask, "TASKID", strAssignValuerId
        xmlTempRequest.appendChild xmlStageTask
        
        strResponse = objTaskManagerBO.TmRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
        
        ' Call CreateAdhocCaseTask
        Set xmlStageTask = Nothing
        Set xmlStageTask = xmlGetMandatoryNode(xmlResponseDoc, ".//STAGETASK")
        
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "CREATEADHOCCASETASK"
        
        For Each xmlAttribute In xmlStageTask.Attributes
            xmlCopyAttribute xmlStageTask, xmlCaseTask, xmlAttribute.nodeName
        Next
        xmlCaseTask.setAttribute "CASESTAGESEQUENCENO", intCaseStageSeqNo
        xmlSetAttributeValue xmlCaseTask, "ACTIVITYINSTANCE", "1"
        
        Set xmlAttribute = xmlGetAttributeNode(xmlCaseTask, "TASKINSTANCE")
        If Not xmlAttribute Is Nothing Then
            xmlCaseTask.removeAttributeNode xmlAttribute
        End If

        xmlCaseTask.setAttribute "TASKSTATUS", omiga4TASKSTATUSPending
        xmlTempRequest.appendChild xmlCaseTask
        xmlTempRequest.appendChild xmlApplication.cloneNode(True)
        
        CreateAdhocCaseTask xmlTempRequest
    End If
    
    'Get the InstructionSequenceNumber of the latest ValuerInstruction
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETVALUERINSTRUCTIONS"
    Set xmlValuerInstruction = xmlResponseDoc.createElement("VALUERINSTRUCTIONS")
    
    xmlCopyAttribute xmlApplication, xmlValuerInstruction, "APPLICATIONNUMBER"
    xmlCopyAttribute xmlApplication, xmlValuerInstruction, "APPLICATIONFACTFINDNUMBER"
    xmlTempRequest.appendChild xmlValuerInstruction
    
    strResponse = objAppProcBO.OmAppProcRequest(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
    
    Dim intInstrutionSeqNo As Integer, strValType As String, intValType As Integer, strCompanyname As String
    
    Set xmlValuerInstruction = xmlResponseDoc.selectSingleNode("//VALUERINSTRUCTIONS")
    intInstrutionSeqNo = xmlGetAttributeText(xmlValuerInstruction, "INSTRUCTIONSEQUENCENO")
    'GD SYS 2092
    strCompanyname = xmlGetAttributeText(xmlValuerInstruction, "COMPANYNAME")
    intValType = xmlGetAttributeText(xmlValuerInstruction, "VALUATIONTYPE")
    strValType = objComboDO.GetComboText("ValuationType", CStr(intValType))
    'Get the details of new valuer report task to be created
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "FINDSTAGETASKLIST"
    Set xmlStageTask = xmlResponseDoc.createElement("STAGETASK")
    Set xmlCaseTask = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
    xmlCopyAttribute xmlCaseTask, xmlStageTask, "STAGEID"
    xmlSetAttributeValue xmlStageTask, "TASKID", strInstructValuerId
    xmlTempRequest.appendChild xmlStageTask
    
    strResponse = objTaskManagerBO.TmRequest(xmlTempRequest.xml)
    Set xmlResponseDoc = xmlLoad(strResponse, strFunctionName)
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
    
    Set xmlStageTask = Nothing
    Set xmlStageTask = xmlResponseDoc.selectSingleNode("//STAGETASK")
    Set xmlTask = xmlResponseDoc.selectSingleNode("//TASK")
    
    '-------------------------------------------------------------------------------------------
    'Check whether Instruct_Valuer records exist on the current stage. If so get the taskStatus
    'from the record with highest taskInstance. Create a new task if the task does not exist or
    'taskStatus is complete.
    '-------------------------------------------------------------------------------------------
    intTaskInstance = 0
    intTemp = 0
    
    strCondition = "//CASETASK[@TASKID='" & strInstructValuerId & "']"
    strTaskStatus = ""
    Set xmlCaseTaskList = xmlCurrentStageDoc.selectNodes(strCondition)
    Dim blnTaskComplete As Boolean
    
    If xmlCaseTaskList.length > 0 Then
        For Each xmlNode In xmlCaseTaskList
            intTemp = xmlGetAttributeText(xmlNode, "TASKINSTANCE")
            If intTemp > intTaskInstance Then
                Set xmlCaseTask = xmlNode
                intTaskInstance = intTemp
                strTaskStatus = xmlGetAttributeText(xmlNode, "TASKSTATUS")
            End If
        Next xmlNode
        blnTaskComplete = objComboDO.IsItemInValidation("TaskStatus", strTaskStatus, "C")
    End If
    
    If xmlCaseTaskList.length = 0 Or blnTaskComplete Then
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "CREATEADHOCCASETASK"
        Set xmlCaseTask = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
        xmlTempRequest.appendChild xmlCaseTask
        
        If Not xmlStageTask Is Nothing Then
            For Each xmlAttribute In xmlStageTask.Attributes
                xmlCopyAttribute xmlStageTask, xmlCaseTask, xmlAttribute.nodeName
            Next
        End If
        
        If Not xmlTask Is Nothing Then
            For Each xmlAttribute In xmlTask.Attributes
                xmlCopyAttribute xmlTask, xmlCaseTask, xmlAttribute.nodeName
            Next
        End If
        
        xmlCaseTask.setAttribute "TASKSTATUS", 10
        xmlCaseTask.setAttribute "CASESTAGESEQUENCENO", intCaseStageSeqNo
        
        Set xmlAttribute = xmlGetAttributeNode(xmlCaseTask, "TASKINSTANCE")
        If Not xmlAttribute Is Nothing Then
            xmlCaseTask.removeAttributeNode xmlAttribute
        End If
        
        xmlSetAttributeValue xmlCaseTask, "ACTIVITYINSTANCE", "1"
        xmlSetAttributeValue xmlCaseTask, "CONTEXT", intInstrutionSeqNo
        
        'GD SYS 2092
        xmlSetAttributeValue xmlCaseTask, "CASETASKNAME", strValType & " " & "Valuation" & " " & strCompanyname
       
        xmlTempRequest.appendChild xmlApplication
        
        CreateAdhocCaseTask xmlTempRequest
    End If
    
    '-----------------------------------------------------------
    'APWP3 - DPF 27/08/2002 - Create Valuation Report - Start
    
    'Declare variables
    Dim strReportResponse As String
    Dim xmlCreateRepElem As IXMLDOMElement
    Dim xmlApplicationElem As IXMLDOMElement
    
    'set up connection to Application Processing Business Object
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
        
    If objAppProcBO Is Nothing Then
        errThrowError strFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    'Build up request block
    Set xmlCreateRepElem = xmlGetRequestNode(vxmlRequestNode).cloneNode(True)
    xmlCreateRepElem.setAttribute "OPERATION", "CreateValuationReportUsingTemplate"
    Set xmlApplicationElem = xmlApplication.cloneNode(True)
    xmlChangeNodeName xmlApplicationElem, "APPLICATION", "VALUATION"
    xmlApplicationElem.setAttribute "INSTRUCTIONSEQUENCENO", intInstrutionSeqNo
    xmlCreateRepElem.appendChild xmlApplicationElem.cloneNode(True)
    
    'Call to App Processing BO to create valuation report record & close connection
    strReportResponse = objAppProcBO.OmAppProcRequest(xmlCreateRepElem.xml)
    
    'Check Response
    errCheckXMLResponse strReportResponse, True
    
    'reset variables used in this section
    Set xmlCreateRepElem = Nothing
    Set xmlApplicationElem = Nothing
    
    'End of APWP3 amendment
    '-----------------------------------------------------------
    
CompleteValuerInstrExit2:

    Set xmlResponseDoc = Nothing
    Set xmlTempRequest = Nothing
    Set objTaskManagerBO = Nothing
    Set objAppProcBO = Nothing
    Set xmlCaseTask = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlCaseTaskList = Nothing
    Set xmlStageTask = Nothing
    Set xmlTask = Nothing
    Set xmlAttribute = Nothing
    Set xmlApplication = Nothing
    Set xmlValuerInstruction = Nothing
    
CompleteValuerInstructionsExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub CreateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo CreateReportOnTitleErr

    Const strFunctionName As String = "CreateReportOnTitle"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "REPORTONTITLE"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement

    ' Call referencesBO.CreateReportOnTitle
    Set objBO = gobjContext.CreateInstance("omROT.omRotBO")
    xmlThisResponseDoc.loadXML objBO.OmRotRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    Debug.Print "CreateOmRotRequest response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    'Call omTMBO.UpdateCaseTask
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"

    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    ' create CASETASK element
    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
    ' clone search attributes
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
    xmlElem.setAttribute "TASKSTATUS", "20" 'Pending

    xmlThisRequestNode.appendChild xmlElem

    ' run Tm UpdateCaseTask request
    xmlThisResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlThisResponseDoc.xml
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete

CreateReportOnTitleExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    Exit Sub
    
CreateReportOnTitleErr:

    gobjContext.SetAbort
    
    GoTo CreateReportOnTitleExit
    
End Sub

Private Sub GetTaskContactDetails( _
                ByVal vxmlRequestNode As IXMLDOMNode, _
                ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetTaskContactDetailsExit
    Const strFunctionName As String = "GetTaskContactDetails"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim strContacttype  As String
    Dim strContactTypeName As String
    Dim strCustomerIdentifier  As String
    Dim strCustomerVersionNumber  As String
    Dim strContext  As String
    Dim strTmpResponse As String
    Dim gobjEnquiry As Object
      
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    Dim xmlChildNode As IXMLDOMNode
        
    Set xmlElement = vxmlRequestNode
    Set xmlParentNode = xmlIn.appendChild(xmlElement)
   
    
    Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("CASETASK")
    
    strContacttype = xmlGetAttributeText(xmlCaseTaskNode, "CONTACTTYPE")
    strContext = xmlGetAttributeText(xmlCaseTaskNode, "CONTEXT")
    
    strCustomerIdentifier = xmlGetAttributeText(xmlCaseTaskNode, "CUSTOMERIDENTIFIER")
    strCustomerVersionNumber = "1"  ' hardcoded for the time being
    '
    ' Get the Contact type name from the combo using the TaskContactType ID
    '
    ' AQR SYS2675  DRC change combo lookup to TaskContactType
    strContactTypeName = GetComboText("TaskContactType", CInt(strContacttype))
    
    Select Case Left(Trim$(strContactTypeName), 3)
    
        Case "Emp" '- Employer
        
            Set xmlChildNode = xmlIn.createElement("EMPLOYMENT")
            xmlParentNode.appendChild xmlChildNode
            
            Set xmlTempNode = xmlIn.createElement("CUSTOMERNUMBER")
            xmlTempNode.Text = strCustomerIdentifier
            xmlChildNode.appendChild xmlTempNode
            
            Set xmlTempNode = xmlIn.createElement("CUSTOMERVERSIONNUMBER")
            xmlTempNode.Text = strCustomerVersionNumber
            xmlChildNode.appendChild xmlTempNode
            
            Set xmlTempNode = xmlIn.createElement("EMPLOYMENTSEQUENCENUMBER")
            xmlTempNode.Text = strContext
            xmlChildNode.appendChild xmlTempNode
            '
            ' Calling CustomerEmploymentBO.GetEmploymentDetails
            '
            Set gobjEnquiry = gobjContext.CreateInstance("omCE.CustomerEmploymentBO")
            strTmpResponse = gobjEnquiry.GetEmploymentDetails(xmlElement.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            '
            ' Error Check
            '
            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True
            
            Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/EMPLOYMENT/NAMEANDADDRESSDIRECTORY")
            
            If xmlParentNode Is Nothing Then
            '
            ' check for ThirdParty if Self Employed
            '
              Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/EMPLOYMENT/THIRDPARTY")
            End If
             
        Case "Acc" '- Accountant

           

            Set xmlChildNode = xmlIn.createElement("EMPLOYMENT")
            xmlParentNode.appendChild xmlChildNode

            Set xmlTempNode = xmlIn.createElement("CUSTOMERNUMBER")
            xmlTempNode.Text = strCustomerIdentifier
            xmlChildNode.appendChild xmlTempNode

            Set xmlTempNode = xmlIn.createElement("CUSTOMERVERSIONNUMBER")
            xmlTempNode.Text = strCustomerVersionNumber
            xmlChildNode.appendChild xmlTempNode

            Set xmlTempNode = xmlIn.createElement("EMPLOYMENTSEQUENCENUMBER")
            xmlTempNode.Text = strContext
            xmlChildNode.appendChild xmlTempNode
            '
            ' Calling CustomerEmploymentBO.GetEmploymentAccountant
            '
            Set gobjEnquiry = gobjContext.CreateInstance("omCE.CustomerEmploymentBO")
            strTmpResponse = gobjEnquiry.GetEmploymentAccountant(xmlElement.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            '
            ' Error Check
            '
            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True
            
            Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/ACCOUNTANT/NAMEANDADDRESSDIRECTORY")
            
        Case "Len" '- Lender
            
            Set xmlChildNode = xmlIn.createElement("ACCOUNT")
            xmlParentNode.appendChild xmlChildNode
            Set xmlTempNode = xmlIn.createElement("ACCOUNTGUID")
            xmlTempNode.Text = strContext
            xmlChildNode.appendChild xmlTempNode
            '
            ' Calling CustomerFinancialBO.GetAccountDetails
            '
            Set gobjEnquiry = gobjContext.CreateInstance("omCF.CustomerFinancialBO")
            strTmpResponse = gobjEnquiry.GetAccountDetails(xmlElement.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            '
            ' Error Check
            '
            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True

            Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/ACCOUNT/NAMEANDADDRESSDIRECTORY")
            '
            ' check for NameandAddressDirectory
            '
            If xmlParentNode Is Nothing Then
            '
            ' check for ThirdParty
            '
              Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/ACCOUNT/THIRDPARTY")

            End If

        Case "Lan"  ' -Landlord

            Set xmlChildNode = xmlIn.createElement("TENANCY")
            xmlParentNode.appendChild xmlChildNode

            Set xmlTempNode = xmlIn.createElement("CUSTOMERNUMBER")
            xmlTempNode.Text = strCustomerIdentifier
            xmlChildNode.appendChild xmlTempNode

            Set xmlTempNode = xmlIn.createElement("CUSTOMERVERSIONNUMBER")
            xmlTempNode.Text = strCustomerVersionNumber
            xmlChildNode.appendChild xmlTempNode

            Set xmlTempNode = xmlIn.createElement("CUSTOMERADDRESSSEQUENCENUMBER")
            xmlTempNode.Text = strContext
            xmlChildNode.appendChild xmlTempNode
            '
            ' Calling CustomerBO.GetTenancy
            '
            Set gobjEnquiry = gobjContext.CreateInstance("omCust.CustomerBO")
            strTmpResponse = gobjEnquiry.GetTenancy(xmlElement.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
    '
    ' Error Check
    '
            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True
       
            Set xmlParentNode = xmlOut.selectSingleNode("RESPONSE/TENANCY/NAMEANDADDRESSDIRECTORY")

        Case Else

            '
            ' raise an error for invalid Contact Type
            '
            errThrowError strFunctionName, oeRecordNotFound, "Unrecognised Task Contact Type ID " & strContacttype
        
    End Select
    '
    ' Build up the Response
    '
    If Not xmlParentNode Is Nothing Then
        
        Set xmlChildNode = xmlParentNode.selectSingleNode("CONTACTDETAILS")
        Set xmlElement = xmlOut.createElement("CONTACTDETAILS")
        xmlElement.setAttribute "CONTACTNAME", xmlGetNodeText(xmlChildNode, "CONTACTFORENAME") & " " & _
                                               xmlGetNodeText(xmlChildNode, "CONTACTSURNAME")
                                                                 ' Company name is up the tree
        xmlElement.setAttribute "CONTACTCOMPANY", xmlGetNodeText(xmlParentNode, "COMPANYNAME")
       
        xmlElement.setAttribute "TELEPHONENUMBER", xmlGetNodeText(xmlChildNode, "TELEPHONENUMBER")
        xmlElement.setAttribute "FAXNUMBER", xmlGetNodeText(xmlChildNode, "FAXNUMBER")
        xmlElement.setAttribute "EMAILADDRESS", xmlGetNodeText(xmlChildNode, "EMAILADDRESS")
                
        'JR - Omiplus24, removed FaxNumber and Replaced TelephoneNumber with below
        Dim xmlCTNode As IXMLDOMNode
        Dim xmlCTNodeList As IXMLDOMNodeList
        
        Set xmlCTNodeList = xmlChildNode.selectNodes(".//CONTACTTELEPHONEDETAILS")
        For Each xmlCTNode In xmlCTNodeList
            Set xmlCTNode = xmlCreateAttributeBasedResponse(xmlCTNode, True)
            xmlElement.appendChild xmlCTNode
        Next
        'End
        
        vxmlResponseNode.appendChild xmlElement
    End If
 
GetTaskContactDetailsExit:
    
    Set xmlCaseTaskNode = Nothing
    Set gobjEnquiry = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlElement = Nothing
    Set xmlParentNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlTempNode = Nothing
    'JR - Omiplus24
    Set xmlCTNode = Nothing
    Set xmlCTNodeList = Nothing

    
    
    errCheckError strFunctionName

End Sub


Private Sub ProcessTask(ByVal vxmlRequest As IXMLDOMNode)

    On Error GoTo ProcessTaskExit
    Const strFunctionName As String = "ProcessTask"
    
    Dim strInterface As String
    Dim strOutputDocument As String
    'Dim strTaskID As String
    Dim strTaskName As String
    Dim strPrinterDestination As String
    Dim strApplicant As String
    Dim strCustomerIdentifier As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskElem As IXMLDOMElement
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim strTmpResponse As String
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlPrintAttributeNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    
    ' PSC 15/09/2005 MAR32 - Start
    Dim objAppManager As Object
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCustRole As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    
    Dim strResponse As String
    ' PSC 15/09/2005 MAR32 - End

    Dim gobjEnquiry As Object
    ' TW 03/11/2005 MAR211
    Dim strPackControlNumber As String
    ' TW 03/11/2005 MAR211 End
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
    xmlThisCaseTaskDoc.validateOnParse = False
    xmlThisCaseTaskDoc.setProperty "NewParser", True
    xmlThisCaseTaskDoc.async = False
    Set xmlCaseTaskNode = vxmlRequest.selectSingleNode("./CASETASK")
            
    'MAR1300 GHun Validate Process Task Authority
    Set xmlTempElem = xmlGetRequestNode(vxmlRequest)
    xmlTempElem.setAttribute "OPERATION", "ValidateProcessTaskAuthority"
    xmlTempElem.appendChild xmlCaseTaskNode.cloneNode(True)
       
    xmlOut.loadXML gobjTmBO.TmRequest(xmlTempElem.xml)
    errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True
    'MAR1300 End
            
    strTaskName = xmlGetAttributeText(xmlCaseTaskNode, "TASKNAME")
    strInterface = xmlGetAttributeText(xmlCaseTaskNode, "INTERFACE")
    strOutputDocument = xmlGetAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT")
       
    'MAR1244 Code to get Customer information moved
    
    Set xmlApplicationNode = vxmlRequest.selectSingleNode("./APPLICATION")
    
    strApplicant = xmlGetAttributeText(xmlCaseTaskNode, "APPLICANT")
    strCustomerIdentifier = xmlGetAttributeText(xmlCaseTaskNode, "CUSTOMERIDENTIFIER")

    ' PSC 15/09/2005 MAR32 - Start
    If (strApplicant = "1") And (Len(strCustomerIdentifier) > 0) Then
        Set xmlCustomerNode = xmlApplicationNode.selectSingleNode("CUSTOMER[@CUSTOMERNUMBER='" & strCustomerIdentifier & "']")
                
        If Not xmlCustomerNode Is Nothing Then
            xmlCopyAttribute xmlCustomerNode, xmlTempNode, "CUSTOMERVERSIONNUMBER"
        Else
            Set xmlTempRequest = vxmlRequest.cloneNode(False)
            Set xmlCustRole = xmlThisCaseTaskDoc.createElement("CUSTOMERROLE")
            xmlTempRequest.appendChild xmlCustRole
            Set xmlNode = xmlThisCaseTaskDoc.createElement("APPLICATIONNUMBER")
            xmlNode.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
            xmlCustRole.appendChild xmlNode
            Set xmlNode = xmlThisCaseTaskDoc.createElement("APPLICATIONFACTFINDNUMBER")
            xmlNode.Text = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
            xmlCustRole.appendChild xmlNode
                    
            Set objAppManager = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
            strResponse = objAppManager.FindCustomerRoleList(xmlTempRequest.xml)
            Set xmlTempResponse = xmlLoad(strResponse, strFunctionName)
            errCheckXMLResponseNode xmlTempResponse.documentElement, , True
            Set xmlCustomerNode = xmlTempResponse.selectSingleNode("RESPONSE/CUSTOMERROLELIST/CUSTOMERROLE[CUSTOMERNUMBER='" & strCustomerIdentifier & "']/CUSTOMERVERSIONNUMBER")
                    
            If Not xmlCustomerNode Is Nothing Then
                xmlSetAttributeValue xmlTempNode, "CUSTOMERVERSIONNUMBER", xmlCustomerNode.Text
            End If
        End If
    End If
    ' PSC 15/09/2005 MAR32 - End
              
    If Len(strInterface) > 0 Then
        '
        ' Call OmTmRequest with the interface & the casetask elemenet
        '
        Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
        Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
        xmlThisCaseTaskElem.setAttribute "OPERATION", strInterface
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "USERID"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "UNITID"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "USERAUTHORITYLEVEL"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "ADMINSYSTEMSTATE"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "CHANNELID"  'MAR1632 GHun
        xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
        xmlThisCaseTaskNode.appendChild xmlApplicationNode
        
        Set xmlPrinterNode = vxmlRequest.selectSingleNode("./PRINTER")
        
        ' PSC 08/01/2003 BM0138 - Start
        If Not xmlPrinterNode Is Nothing Then
            xmlThisCaseTaskNode.appendChild xmlPrinterNode
        End If
        ' PSC 08/01/2003 BM0138 - Start
        
        '
        '
        strTmpResponse = OmTmRequest(xmlThisCaseTaskElem.xml)
        Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
        '
        ' Error Check
        '
         Set xmlTempNode = xmlOut.firstChild
         errCheckXMLResponseNode xmlTempNode, , True
 
    End If
    
    If Len(strOutputDocument) > 0 And Len(strInterface) = 0 Then
        '
        ' Call OmPrintBO.GetPrintAtributes
        '
        ' PSC 26/11/01 SYS3195 - Start
        On Error Resume Next
        Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        On Error GoTo ProcessTaskExit

        If Not gobjEnquiry Is Nothing Then
        
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "GetPrintAttributes"
                      
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("FINDATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            '
            'AQR 2266 DRC - changed Atrribute name from DOCUMENTID to HOSTTEMPLATEID
            '
            xmlTempElem.setAttribute "HOSTTEMPLATEID", strOutputDocument
                
            strTmpResponse = gobjEnquiry.omRequest(xmlThisCaseTaskElem.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            '
            ' Error Check
            '
            Set xmlTempNode = xmlOut.firstChild
            errCheckXMLResponseNode xmlTempNode, , True
            
            Set xmlTempElem = xmlOut.selectSingleNode("RESPONSE/ATTRIBUTES")
            
            If xmlTempElem.getAttribute("INACTIVEINDICATOR") = "1" Then
              errThrowError strFunctionName, 4820, "", strTaskName
            End If
            
            If Len(xmlTempElem.getAttribute("DEFAULTCOPIES")) = 0 Then
              errThrowError strFunctionName, 4823, "", strTaskName
            End If
            
            strPrinterDestination = xmlGetAttributeText(xmlTempElem, "PRINTERDESTINATIONTYPE")
            ' EP2_1433 - use combovalidation types
            ' strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
'            If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
'                And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0) _
'                And (InStr(1, strPrinterDestination, "W", vbTextCompare) > 0)) Then     'MAR7 GHun
'               errThrowError cstrFunctionName, 4822, strPrinterDestination, "IssueOffer"
'            End If
            If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
            And IsValidationType("PrinterDestination", strPrinterDestination, "R") _
            And IsValidationType("PrinterDestination", strPrinterDestination, "W") _
            Then
                errThrowError strFunctionName, 4822, GetComboText("PrinterDestination", CInt(strPrinterDestination)), strFunctionName
            End If
            
            Set xmlPrintAttributeNode = xmlTempElem
'
'           Call OmTmBo.PrintDocumentForTask
'
            Set xmlThisCaseTaskElem = Nothing
            Set xmlThisCaseTaskNode = Nothing
            Set xmlThisCaseTaskDoc = Nothing
            'BMIDS697 Clear xmlThisCaseTaskDoc
            Set xmlThisCaseTaskDoc = New FreeThreadedDOMDocument40
            xmlThisCaseTaskDoc.validateOnParse = False
            xmlThisCaseTaskDoc.setProperty "NewParser", True
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "PrintDocumentForTask"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "USERID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "UNITID"
            xmlCopyAttribute vxmlRequest, xmlThisCaseTaskNode, "USERAUTHORITYLEVEL"
            xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
            
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("APPLICATION")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            Set xmlApplicationNode = vxmlRequest.selectSingleNode("APPLICATION")
            xmlCopyAttribute xmlApplicationNode, xmlTempNode, "APPLICATIONFACTFINDNUMBER"
            
            xmlThisCaseTaskNode.appendChild xmlTempNode
            
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("PRINTATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            'AQR SYS2266 - DRC Print Attribute names have changed
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DEFAULTCOPIES", "COPIES"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "PDMMETHOD", "METHODNAME"
            
            'AQR SYS4467 - DR Need to add more print attributes for DMS
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "HOSTTEMPLATENAME"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "HOSTTEMPLATEDESCRIPTION"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "PRINTERDESTINATIONTYPE"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "MAXCOPIES"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE" 'MAR7 GHun
            
            'AQR SYS2266 - DRC
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
            
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "EMAILRECIPIENTTYPE"   'EP2_1903 GHun
            
            ' EP2_1433
            ' xmlTempElem.setAttribute "DESTINATIONTYPE", Left(strPrinterDestination, 1)
            xmlTempElem.setAttribute "DESTINATIONTYPE", GetValidationTypeForValueID("PrinterDestination", strPrinterDestination)
            
            ' EP2_1433 - use combovalidation types
            If IsValidationType("PrinterDestination", strPrinterDestination, "L") _
            Or IsValidationType("PrinterDestination", strPrinterDestination, "W") _
            Then
            ' If strPrinterDestination = "Local" Or strPrinterDestination = "Workstation Printer" Then    'MAR7 GHun
                Set xmlPrinterNode = vxmlRequest.selectSingleNode("PRINTER")
                xmlTempElem.setAttribute "PRINTER", xmlGetAttributeText(xmlPrinterNode, "PRINTERNAME")
            'RF 16/02/2006 MAR1251 Start - Allow template destination of DMS only
            ' EP2_1433 - no CURRENT requirement for "Document Store Only"
            'EP2_1903 GHun Do nothing for email fulfilment
            ElseIf GetComboText("PrinterDestination", CInt(strPrinterDestination)) = "Document Store Only" _
                Or IsValidationType("PrinterDestination", strPrinterDestination, "EF") Then
                ' do nothing
            'RF 16/02/2006 MAR1251 End
            Else    '"Remote" Or "File"
                 xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "REMOTEPRINTERLOCATION", "PRINTER"
            End If
           
            xmlThisCaseTaskNode.appendChild xmlTempNode
            
            strTmpResponse = OmTmRequest(xmlThisCaseTaskElem.xml)
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            '
            ' Error Check
            '
             Set xmlTempNode = xmlOut.firstChild
             errCheckXMLResponseNode xmlTempNode, , True
        End If
        ' PSC 26/11/01 SYS3195 - End
    End If
    
    ' TW 03/11/2005 MAR211
    strPackControlNumber = xmlGetAttributeText(xmlCaseTaskNode, "PACKCONTROLNUMBER")
    
    ' PSC 03/02/2006 MAR1156 - Start
    If Len(strPackControlNumber) > 0 And Len(strInterface) = 0 Then
        SendPack vxmlRequest, strPackControlNumber
    End If
    ' PSC 03/02/2006 MAR1156 - End
    
    ' TW 03/11/2005 MAR211 End
    
    If Len(strOutputDocument) = 0 And Len(strInterface) = 0 Then
    '
    ' Call TaskManagementBO.UpdateCaseTask
    '
        Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
        Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
        xmlThisCaseTaskElem.setAttribute "OPERATION", "UpdateCaseTask"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERID"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "UNITID"
        xmlCopyAttribute vxmlRequest, xmlThisCaseTaskElem, "USERAUTHORITYLEVEL"
        xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "40"
        xmlThisCaseTaskNode.appendChild xmlCaseTaskNode

        
        ' run Tm UpdateCaseTask request
        xmlOut.async = False
        RunTmRequest xmlThisCaseTaskDoc, xmlOut
    
        Debug.Print "UpdateCaseTask response:"
        Debug.Print xmlOut.xml
        ' parse response
        errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True
        
    End If
    
ProcessTaskExit:

    Set xmlOut = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlPrintAttributeNode = Nothing
    Set xmlPrinterNode = Nothing
    
    ' PSC 15/09/2005 MAR32 - Start
    Set objAppManager = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCustRole = Nothing
    Set xmlNode = Nothing
    Set xmlTempResponse = Nothing
    ' PSC 15/09/2005 MAR32 - End

    errCheckError strFunctionName
    
End Sub

Private Sub ValidateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo ValidateReportOnTitleErr
    
    Const cstrFunctionName As String = "ValidateReportOnTitle"
                               
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "APPLICATION"
    
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlROTNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlThisTaskNode  As IXMLDOMNode
      
    Dim blnReissueOfferInd As Boolean
    
    Dim strTMReprocessCreditCheck As String
    Dim strTMRevAndDecis As String
    Dim strTaskDueDateAndTime As String
    Dim strResponse As String
    Dim strMandatoryIndicator As String 'KW
        
    blnReissueOfferInd = False
    
'Stage1:
    
    ' Call omRotBO.ValidateReportOnTitle
    Set objBO = gobjContext.CreateInstance("omROT.omRotBO")
    xmlThisResponseDoc.loadXML objBO.OmRotRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True

    Set xmlROTNode = xmlThisResponseDoc.selectSingleNode(".//ROT")
    If Not xmlROTNode Is Nothing Then
        blnReissueOfferInd = xmlGetAttributeAsBoolean(xmlROTNode, "REISSUEOFFERIND")
        vxmlResponseNode.appendChild xmlROTNode
    End If

    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
'Stage2:

    If blnReissueOfferInd = True Then
        
        strTMReprocessCreditCheck = GetGlobalParamString("TMReprocessCreditCheck")
        
        'Get TaskDueDateAndTime Used when Creating CaseTask
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskList"
        
        Set xmlElem = xmlThisRequestDoc.createElement("TASK")
        xmlElem.setAttribute "TASKID", strTMReprocessCreditCheck
        
        xmlThisRequestNode.appendChild xmlElem
        
        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlThisRequestNode.xml)
        
        xmlThisResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        
        Set xmlThisTaskNode = xmlThisResponseDoc.selectSingleNode(".//TASK")
        If Not xmlThisTaskNode Is Nothing Then ' error ???
            SetCaseTaskDueDateAndTime xmlThisTaskNode, xmlThisTaskNode
        End If
        
        strTaskDueDateAndTime = xmlGetAttributeText(xmlThisTaskNode, "TASKDUEDATEANDTIME", "")
        
        'KW
        'Get StageTask to find mandatory indicator
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageTask"
        
        Set xmlElem = xmlThisRequestDoc.createElement("STAGETASK")
        xmlElem.setAttribute "TASKID", strTMReprocessCreditCheck
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        
        xmlThisRequestNode.appendChild xmlElem
        
        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlThisRequestNode.xml)
        
        xmlThisResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        
        Set xmlThisTaskNode = xmlThisResponseDoc.selectSingleNode(".//STAGETASK")
        
        strMandatoryIndicator = xmlGetAttributeText(xmlThisTaskNode, "MANDATORYFLAG")
        'KW End
        
        ' create CreateCaseTask request
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation attribute

        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"

        ' create CaseTask element
        Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        
        'MV - 27/02/2003 - BM0362
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
        xmlSetAttributeValue xmlElem, "TASKID", strTMReprocessCreditCheck
        xmlSetAttributeValue xmlElem, "TASKDUEDATEANDTIME", strTaskDueDateAndTime
        'KRW
        xmlSetAttributeValue xmlElem, "MANDATORYINDICATOR", strMandatoryIndicator
        'KW End
        xmlThisRequestNode.appendChild xmlElem

        ' run Tm CreateCaseTask request
        xmlThisResponseDoc.async = False
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

        ' parse response
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        
        'BMIDS668 DRC
        
        strTMRevAndDecis = GetGlobalParamString("TMRevAndDecis")
        
        'Get TaskDueDateAndTime Used when Creating CaseTask
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskList"

        Set xmlElem = xmlThisRequestDoc.createElement("TASK")
        xmlElem.setAttribute "TASKID", strTMRevAndDecis
    
        xmlThisRequestNode.appendChild xmlElem

        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlThisRequestNode.xml)
    
        xmlThisResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
        Set xmlThisTaskNode = xmlThisResponseDoc.selectSingleNode(".//TASK")
        If Not xmlThisTaskNode Is Nothing Then ' error ???
            SetCaseTaskDueDateAndTime xmlThisTaskNode, xmlThisTaskNode
        End If
           
        strTaskDueDateAndTime = xmlGetAttributeText(xmlThisTaskNode, "TASKDUEDATEANDTIME", "")
        
        'KW
        'Get StageTask to find mandatory indicator
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageTask"
        
        Set xmlElem = xmlThisRequestDoc.createElement("STAGETASK")
        xmlElem.setAttribute "TASKID", strTMReprocessCreditCheck
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        
        xmlThisRequestNode.appendChild xmlElem
        
        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlThisRequestNode.xml)
        
        xmlThisResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        
        Set xmlThisTaskNode = xmlThisResponseDoc.selectSingleNode(".//STAGETASK")
        
        strMandatoryIndicator = xmlGetAttributeText(xmlThisTaskNode, "MANDATORYFLAG")
        'KW End
        
        ' create CreateCaseTask request
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"
        ' create CaseTask element
        Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        
        'MV - 27/02/2003 - BM0362
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
        xmlSetAttributeValue xmlElem, "TASKID", strTMRevAndDecis
        xmlSetAttributeValue xmlElem, "TASKDUEDATEANDTIME", strTaskDueDateAndTime
        'KRW
        xmlSetAttributeValue xmlElem, "MANDATORYINDICATOR", strMandatoryIndicator
        'KW End
        xmlThisRequestNode.appendChild xmlElem

        ' run Tm CreateCaseTask request
        xmlThisResponseDoc.async = False
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

        ' parse response
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        
        'BMIDS6668 - End
    
    End If
    
'Stage3:
    
    'DB BMIDS00862 - solicitors details not being captured as part of ROT for bmids.
    'If blnSolBankAccMatch = True Then
        'Call omTMBO.UpdateCaseTask
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKINSTANCE"
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
        ' DB SYS3520 03/01/02 - Added required primary key.
        xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASESTAGESEQUENCENO"
        ' DB End

        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        ' copy request node
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set operation attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
        ' create CASETASK element
        Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
        ' clone search attributes
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"    'EP2_2262 GHun
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
        ' DB SYS3520 03/01/02 - Added required primary key.
        xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
        ' DB End
        xmlElem.setAttribute "TASKSTATUS", "40" 'Complete

        xmlThisRequestNode.appendChild xmlElem

        ' run Tm UpdateCaseTask request
        xmlThisResponseDoc.async = False
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc

        Debug.Print "UpdateCaseTask response:"
        Debug.Print xmlThisResponseDoc.xml
        ' parse response
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    'DB BMIDS00862
    'End If
    
    ' PSC 06/12/01 SYS3290
    gobjContext.SetComplete
        
ValidateReportOnTitleExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
    
ValidateReportOnTitleErr:

    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    
    GoTo ValidateReportOnTitleExit

End Sub


Private Sub ReceivedOffer( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo ReceivedOfferExit
    Const cstrFunctionName As String = "ReceivedOffer"
    
    Dim objAppBO As Object
    Dim objBO As Object
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTmRequest As IXMLDOMNode
    Dim strResponse As String
    Dim xmlTempReq As FreeThreadedDOMDocument40
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempChildElem As IXMLDOMElement
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strOfferSeqNo As String
    
    
    'BG Get mandatory values
    
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    Set xmlCaseTaskElem = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
    strOfferSeqNo = xmlGetMandatoryAttributeText(xmlCaseTaskElem, "CONTEXT")
    
    
    'BG Construct element based request to call into omApp
                  
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    
    Set xmlTempReq = New FreeThreadedDOMDocument40
    xmlTempReq.validateOnParse = False
    xmlTempReq.setProperty "NewParser", True
    Set xmlTempElem = xmlTempReq.createElement("APPLICATIONOFFER")
    
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONNUMBER")
    xmlTempChildElem.Text = strApplicationNumber
    xmlTempElem.appendChild xmlTempChildElem
    
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempChildElem.Text = strApplicationFactFindNumber
    xmlTempElem.appendChild xmlTempChildElem
    
    Set xmlTempChildElem = xmlTempReq.createElement("OFFERSEQUENCENO")
    xmlTempChildElem.Text = strOfferSeqNo
    xmlTempElem.appendChild xmlTempChildElem
    
    xmlNewRequest.appendChild xmlTempElem
    
    'BG call omApp.ApplicationBO_UpdateApplicationOffer method with element based request.
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    strResponse = objAppBO.UpdateApplicationOffer(xmlNewRequest.xml)
    
    
    'DRC SYS4236 - Update CaseTask to complete if successful

    Dim lngErrNo As Long
    lngErrNo = errCheckXMLResponse(strResponse, False)
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, 4824
    End If
    
    ' Now update Case Task
    Set xmlTmRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlCaseTaskElem = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    xmlSetAttributeValue xmlTmRequest, "OPERATION", "UPDATECASETASK"
    
    xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", 40
    
    xmlTmRequest.appendChild xmlCaseTaskElem
    
    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    strResponse = objBO.TmRequest(xmlTmRequest.xml)
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    
    
    
ReceivedOfferExit:
        
    Set objAppBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'BMIDS00005 CRWP2 BM065 UpdateCRSContactLog
Private Sub UpdateCRSContactLog( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:  Calls CustomerBO.UpdateCRSContactLog and updates the task status
' pass: vxmlRequestNode     XML REQUEST node containing application and casetask details
'       vxmlResponseNode    XML RESPONSE node
'------------------------------------------------------------------------------------------

    On Error GoTo UpdateCRSContactLogExit
    
    Const strFunctionName As String = "UpdateCRSContactLog"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlContactLog       As IXMLDOMElement
    Dim xmlApplication      As IXMLDOMElement
    Dim xmlNewRequest       As IXMLDOMNode
    Dim xmlCaseTask         As IXMLDOMNode
    Dim objCustomerBO       As Object
    Dim objMsgTmBO          As Object
    Dim strValidationType   As String
    Dim strResponse         As String
    
    ' Create the Request node for calling CustomerBO
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlContactLog = xmlDoc.createElement("CONTACTLOG")
    xmlNewRequest.appendChild xmlContactLog
    ' BMIDS00058 GHun Changed xmlApplication element name from APPLICATION to APPLICATIONNUMBER
    Set xmlApplication = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlApplication.Text = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")
    xmlContactLog.appendChild xmlApplication
        
    ' Call CustomerBO to get the contacts and pass them to the admin system for updating
    'BMIDS00005 replaced "omCust" with constant
    Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    strResponse = objCustomerBO.UpdateCRSContactLog(xmlNewRequest.xml)
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    ' update the task as complete
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    ' Get the task status for validation type "CP"
    strValidationType = GetFirstComboValueId("TaskStatus", "CP")

    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")

    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(xmlNewRequest, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(xmlNewRequest, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlNewRequest, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", strValidationType
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)

    'BMIDS00005 replaced "MsgTm" with constant
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
UpdateCRSContactLogExit:

    Set xmlDoc = Nothing
    Set xmlContactLog = Nothing
    Set xmlApplication = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTask = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub
'BMIDS00005 End

'BMIDS00026 IWP2 BM060 UpdateCRSCustomer
Private Sub UpdateCRSCustomer( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:  Calls CustomerBO.UpdateCRSCustomer and updates the task status
' pass: vxmlRequestNode     XML REQUEST node containing application and casetask details
'       vxmlResponseNode    XML RESPONSE node
'------------------------------------------------------------------------------------------

    On Error GoTo UpdateCRSCustomerExit
    
    Const strFunctionName As String = "UpdateCRSCustomer"
    
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlCustomer         As IXMLDOMElement
    Dim xmlApplication      As IXMLDOMElement
    Dim xmlNewRequest       As IXMLDOMNode
    Dim xmlCaseTask         As IXMLDOMNode
    Dim objCustomerBO       As Object
    Dim objMsgTmBO          As Object
    Dim strValidationType   As String
    Dim strResponse         As String
    
    ' Create the Request node for calling CustomerBO
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlCustomer = xmlDoc.createElement("CUSTOMERS")
    xmlNewRequest.appendChild xmlCustomer
    Set xmlApplication = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlApplication.Text = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("APPLICATION"), "APPLICATIONNUMBER")
    xmlCustomer.appendChild xmlApplication
        
    ' Call CustomerBO to get the customer details and pass them to the admin system for updating
    'BMIDS00026 replaced "omCust" with constant
    Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    strResponse = objCustomerBO.UpdateCRSCustomer(xmlNewRequest.xml)
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    ' update the task as complete
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    ' Get the task status for validation type "CP"
    strValidationType = GetFirstComboValueId("TaskStatus", "CP")

    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")

    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(xmlNewRequest, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(xmlNewRequest, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlNewRequest, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", strValidationType
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)

    'BMIDS00026 replaced "MsgTm" with constant
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
UpdateCRSCustomerExit:

    Set xmlDoc = Nothing
    Set xmlCustomer = Nothing
    Set xmlApplication = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTask = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub
'BMIDS00026 End

'BMIDS00025 MDC 10/06/2002
Private Sub RunHunterInterface(ByVal vxmlRequestNode As IXMLDOMNode)

On Error GoTo RunHunterInterfaceExit
    
Const cstrFunctionName As String = "RunHunterInterface"
    
Dim xmlApplicationElem As IXMLDOMElement
Dim xmlNewRequest As IXMLDOMElement
Dim xmlCaseTaskElem As IXMLDOMElement

Dim objBO As Object

Dim strApplicationNumber As String
Dim strAFFNumber As String
Dim strResponse As String
Dim strValidationType As String
    
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
    
    'Create Request and run Hunter Interface
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApplicationElem = vxmlRequestNode.selectSingleNode(".//APPLICATION")
    
    xmlNewRequest.appendChild xmlApplicationElem
    
    Set objBO = gobjContext.CreateInstance(gstrHUNTERINTERFACE_COMPONENT & ".omHIBO")
    strResponse = objBO.GenerateRecords(xmlNewRequest.xml)
    
    Dim lngErrNo As Long
    'MO - 08/11/2002 - BMIDS00752 - Made change to pass the error description from hunter
    Dim strErrDescription As String
    
    On Error Resume Next
    
    errCheckXMLResponse strResponse, True
    
    lngErrNo = Err.Number
    strErrDescription = Err.Description
    
    On Error GoTo RunHunterInterfaceExit
    
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, oeTmHunterInterfaceFailed, strErrDescription
    End If
    
    ' Get the task status for validation type "CP"
    strValidationType = GetFirstComboValueId("TaskStatus", "CP")
    
    ' Now update Case Task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
    Set xmlCaseTaskElem = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
    xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", strValidationType
    
    xmlNewRequest.appendChild xmlCaseTaskElem
    
    'BMIDS00142 MDC 28/06/2002 - Code review changes
'    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    'BMIDS00142 MDC 28/06/2002 - End
    
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    
    errCheckXMLResponse strResponse, True
    
RunHunterInterfaceExit:
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    Set objBO = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'BMIDS00025 MDC 10/06/2002 - End

'BMIDS00006 MDC 17/07/2002
Private Sub ImportCustomersIntoApplication(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ImportCustomersIntoApplicationExit

Const cstrFunctionName As String = "ImportCustomersIntoApplication"
    
Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
Dim xmlThisRequestNode As IXMLDOMElement
Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
Dim xmlCaseActivityRequestNode As IXMLDOMNode
Dim xmlImportAccountsResponseNode As IXMLDOMNode
Dim xmlElem As IXMLDOMElement
Dim xmlSrceNode As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim objBO As Object
    
Dim strResponse As String
    
    'Validate the Request
    Set xmlCaseActivityRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASEACTIVITY")
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseActivityRequestNode, "ACTIVITYID"

    'Convert Request to Element based for call to Phase 1 component
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "IMPORTCUSTOMERSINTOAPPLICATION", True)
    
    'Call ApplicationManagerBO.ImportCustomersIntoApplication
    Set objBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strResponse = objBO.ImportCustomersIntoApplication(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    Set objBO = Nothing
    
    Debug.Print "ImportCustomersIntoApplication response:"
    Debug.Print xmlThisResponseDoc.xml
    
    ' parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONNUMBER"
    xmlCheckMandatoryNode xmlThisResponseDoc, "RESPONSE/APPLICATION/APPLICATIONFACTFINDNUMBER"
    
    Set xmlImportAccountsResponseNode = _
        xmlThisResponseDoc.selectSingleNode("RESPONSE")

    ' create CreateActivity request =================================================
    Set xmlThisRequestNode = xmlThisRequestDoc.selectSingleNode("REQUEST")
    xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
    xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "CreateActivity"
    
    ' clone CASEACTIVITY node from original request
    Set xmlNode = _
        xmlThisRequestNode.appendChild(xmlCaseActivityRequestNode.cloneNode(False))
    
    ' get APPLICATIONNUMBER node from CreateApplicationAndCustomer response
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    
    ' create CASEID attribute from APPLICATIONNUMBER
    xmlSetAttributeValue xmlNode, "CASEID", xmlSrceNode.Text
    
    ' create APPLICATION node from CreateApplicationAndCustomer response
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    Set xmlSrceNode = xmlImportAccountsResponseNode.selectSingleNode("APPLICATION/APPLICATIONFACTFINDNUMBER")
    xmlElem.setAttribute xmlSrceNode.nodeName, xmlSrceNode.Text
    
    'xmlThisRequestNode.appendChild xmlCaseActivityRequestNode.cloneNode(False)
    xmlThisRequestNode.appendChild xmlElem.cloneNode(False)
    
    ' run CreateActivity ============================================================
    
    Debug.Print "CreateActivity request:"
    Debug.Print xmlThisRequestNode.xml
    
    CreateActivity xmlThisRequestNode, vxmlResponseNode
    
ImportCustomersIntoApplicationExit:
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlCaseActivityRequestNode = Nothing
    Set xmlImportAccountsResponseNode = Nothing
    Set xmlElem = Nothing
    Set xmlSrceNode = Nothing
    Set xmlNode = Nothing
    Set objBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'BMIDS00025 MDC 10/06/2002 - End

'BMIDS00336 MDC 23/08/2002
Private Sub ImportBureauData(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    ImportBureauData
' author:       Mark Coleman
' date:         22/08/2002
' description:  Imports data saved from a Full Bureau download into Omiga
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
' notes:        If this method has been run as part of the Credit_Check task, an adhoc
'               case task will be created for Import_Bureau_Data so that the task can be run
'               again independently at a later date.
'------------------------------------------------------------------------------------------
On Error GoTo ImportBureauDataErr

Const cstrFunctionName As String = "ImportBureauData"

Dim objNTxBO As IOmTmNTxBO

Dim xmlTempRequest As IXMLDOMElement
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlStageTask As IXMLDOMNode
Dim xmlCaseTask As IXMLDOMElement
Dim xmlCaseActivity As IXMLDOMElement
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlNode As IXMLDOMNode
Dim xmlTaskElem As IXMLDOMElement

Dim strImportBureauDataTaskId As String
Dim strResponse As String
Dim lngErr As Long
Dim intCaseStageSeqNo As Integer

'BMIDS658
Dim objNoTxBO           As omTmNoTxBO
Dim xmlApplicationNode  As IXMLDOMNode
Dim xmlCustomerListNode As IXMLDOMNode
Dim xmlBureauData       As IXMLDOMNode
'BMIDS658 End

    'Initialise
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.validateOnParse = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False
    
    'Validate
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")

    'BMIDS658 Call GetAndDeduplicateBureauRecords outside of the omAIP.ImportBureauData transaction,
    'so the locks created by reading the data are released before any updates occur
    
    'The application node will already be element based when called from RunCreditCheck
    If Not xmlApplicationNode.hasChildNodes Then
        Set xmlApplicationNode = xmlMakeNodeElementBased(xmlApplicationNode, True, "APPLICATION")
    End If
    
    'The customer list may already exist if called from RunCreditCheck otherwise get it
    'from the database
    Set xmlCustomerListNode = xmlGetNode(vxmlRequestNode, "CUSTOMERLIST")
    If xmlCustomerListNode Is Nothing Then
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        xmlTempRequest.setAttribute "OPERATION", "FINDCUSTOMERSFORCREDITCHECK"
        xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
        
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
        strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
        xmlResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
        Set xmlCustomerListNode = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "CUSTOMERLIST")
    End If
    
    'Import the bureau data =======================================================================
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlTempRequest.setAttribute "OPERATION", "GETBUREAURECORDS"
    xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
    xmlTempRequest.appendChild xmlCustomerListNode.cloneNode(True)
    
    If objNoTxBO Is Nothing Then
        Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
    End If
    
    strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
    xmlResponseDoc.loadXML strResponse
    errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    Set xmlBureauData = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "BUREAUDATA")
    xmlTempRequest.removeAttribute "OPERATION"
    
    xmlTempRequest.appendChild xmlBureauData
    xmlTempRequest.appendChild xmlCaseTask.cloneNode(True)
    'BMIDS658 End

    'Call NTxBO.ImportBureauData
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
    objNTxBO.ImportBureauData xmlTempRequest
    
    'BM0376 Call NTxBO.RunIncomeCalcs
    objNTxBO.RunIncomeCalcs vxmlRequestNode
    'BM0376 End

ImportBureauDataExit:
    Set objNTxBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlStageTask = Nothing
    Set xmlCaseTask = Nothing
    Set xmlAttribute = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlNode = Nothing
    Set xmlTaskElem = Nothing
    Set objNoTxBO = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerListNode = Nothing
    Set xmlBureauData = Nothing
    
    If lngErr <> 0 Then
        'INR BMIDS593 The original error description is written to the event log.
        'Add to the error message to notify the user.
        errThrowError cstrFunctionName, lngErr, "See application event log for details."
    End If
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
ImportBureauDataErr:
    
    'If necessary create an adhoc task
    If xmlGetAttributeAsBoolean(vxmlRequestNode, "CREDITCHECK") Then
        'This has been run from a Credit Check task so generate an adhoc case task for ImportBureauData
        strImportBureauDataTaskId = GetMandatoryGlobalParamString("TMImportBureauData")
    
        'Get details for the ImportBureauData task
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETTASKDETAIL"
        Set xmlTaskElem = xmlResponseDoc.createElement("TASK")
        xmlSetAttributeValue xmlTaskElem, "TASKID", strImportBureauDataTaskId
        xmlTempRequest.appendChild xmlTaskElem
        
        strResponse = gobjTmBO.TmRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
        Set xmlTaskElem = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "//TASK")
        
        'Get Current stage details --------------------------------------------------
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "GETCURRENTSTAGE"
        
        Set xmlCaseActivity = xmlResponseDoc.createElement("CASEACTIVITY")
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "CASEID"
        xmlCopyAttribute xmlCaseTask, xmlCaseActivity, "ACTIVITYID"
    
        xmlTempRequest.appendChild xmlCaseActivity
    
        strResponse = gobjTmBO.TmRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
    
        Set xmlNode = xmlGetMandatoryNode(xmlResponseDoc.documentElement, "//CASESTAGE")
        intCaseStageSeqNo = xmlGetAttributeText(xmlNode, "CASESTAGESEQUENCENO")


        ' Find the Stage Task List for the current stage -----------------------------
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "FINDSTAGETASKLIST"
        Set xmlStageTask = xmlResponseDoc.createElement("STAGETASK")
        xmlCopyAttribute xmlCaseTask, xmlStageTask, "STAGEID"
        xmlSetAttributeValue xmlStageTask, "TASKID", strImportBureauDataTaskId
        xmlTempRequest.appendChild xmlStageTask
        
        strResponse = gobjTmBO.TmRequest(xmlTempRequest.xml)
        Set xmlResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlResponseDoc.documentElement, , True
        
        ' Call CreateAdhocCaseTask --------------------------------------------------
        Set xmlStageTask = xmlGetMandatoryNode(xmlResponseDoc, ".//STAGETASK")
        
        Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlSetAttributeValue xmlTempRequest, "OPERATION", "CREATEADHOCCASETASK"
        
        For Each xmlAttribute In xmlStageTask.Attributes
            xmlCopyAttribute xmlStageTask, xmlCaseTask, xmlAttribute.nodeName
        Next
        xmlCaseTask.setAttribute "CASESTAGESEQUENCENO", intCaseStageSeqNo
        xmlSetAttributeValue xmlCaseTask, "ACTIVITYINSTANCE", "1"
        
        xmlCopyAttributeValue xmlTaskElem, xmlCaseTask, "TASKNAME", "CASETASKNAME"
        xmlCopyAttributeValue xmlTaskElem, xmlCaseTask, "TASKNAME", "TASKNAME"
        xmlCopyAttributeValue xmlTaskElem, xmlCaseTask, "INTERFACE", "INTERFACE"
        
        Set xmlAttribute = xmlGetAttributeNode(xmlCaseTask, "TASKINSTANCE")
        If Not xmlAttribute Is Nothing Then
            xmlCaseTask.removeAttributeNode xmlAttribute
        End If

        xmlCaseTask.setAttribute "TASKSTATUS", omiga4TASKSTATUSIncomplete 'omiga4TASKSTATUSPending
        xmlTempRequest.appendChild xmlCaseTask
        'BMIDS658 Make sure APPLICATIONPRIORITY is available as an attribute
        xmlSetAttributeValue xmlApplicationNode, "APPLICATIONPRIORITY", xmlGetMandatoryNodeText(xmlApplicationNode, "APPLICATIONPRIORITY")
        xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
        
        CreateAdhocCaseTask xmlTempRequest
        
        lngErr = oeTmCCOKBureauImportFailed
    Else
        'BMIDS658 Write an error to the event log if something went wrong
        App.LogEvent "Error: " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description, vbLogEventTypeError
        'Run from ImportBureauData task so just raise error
        lngErr = oeTmBureauDataImportFailed
    End If
    
    GoTo ImportBureauDataExit
    
End Sub
'BMIDS00336 MDC 23/08/2002 - End

' PSC 21/10/2002 BMIDS00676 - Start
Private Sub RunGenericInterface(ByVal vxmlRequestNode As IXMLDOMNode, _
                           ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo RunGenericInterfaceExit

Const cstrFunctionName As String = "RunGenericInterface"
        
    Dim objAppManagerBO As Object
    Dim objAdminIFBO As Object
    Dim objMsgTmBO As Object
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlListNode As IXMLDOMNode
    Dim xmlCustomers As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String

    ' Validate Request
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' Get the customers for this application
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    xmlTempRequest.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    xmlTempRequest.appendChild xmlNode
    
    Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
    strResponse = objAppManagerBO.FindCustomersForApplication(xmlTempRequest.xml)
    Set xmlTempDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlTempDoc.documentElement, , True
    
    ' Run generic interface
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "RunGenericInterface"
    Set xmlNode = xmlTempDoc.createElement("APPLICATION")
    xmlTempRequest.appendChild xmlNode
    xmlCopyAttribute xmlApplication, xmlNode, "APPLICATIONNUMBER"
    xmlCopyAttribute xmlApplication, xmlNode, "APPLICATIONFACTFINDNUMBER"
    xmlCopyAttribute xmlCaseTask, xmlNode, "TASKID"
    
    Set xmlListNode = xmlTempDoc.createElement("CUSTOMERLIST")
    xmlNode.appendChild xmlListNode
    
    ' PSC 24/10/2002 BMIDS00710
    Set xmlCustomers = xmlGetMandatoryNodeList(xmlTempDoc, "RESPONSE/CUSTOMER/CUSTOMERVERSION")
    
    For Each xmlCustomer In xmlCustomers
        Set xmlNode = xmlTempDoc.createElement("CUSTOMER")
        xmlSetAttributeValue xmlNode, "CUSTOMERNUMBER", xmlGetMandatoryNodeText(xmlCustomer, "CUSTOMERNUMBER")
        ' PSC 24/10/2002 BMIDS00710
        xmlSetAttributeValue xmlNode, "CUSTOMERVERSIONNUMBER", xmlGetMandatoryNodeText(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        xmlListNode.appendChild xmlNode
    Next
    
    ' PSC 28/10/2002 BMIDS00736 - Start
    Set xmlNode = xmlTempDoc.createElement("CASETASK")
    xmlTempRequest.appendChild xmlNode
    xmlCopyAttributeValue xmlCaseTask, xmlNode, "STAGEID", "STAGEID"
    xmlCopyAttributeValue xmlCaseTask, xmlNode, "CASESTAGESEQUENCENO", "CASESTAGESEQUENCENO"
    xmlCopyAttributeValue xmlCaseTask, xmlNode, "TASKINSTANCE", "TASKINSTANCE"
    ' PSC 28/10/2002 BMIDS00736 - End

    Set objAdminIFBO = gobjContext.CreateInstance(gstrADMIN_INTERFACE & ".AdminInterfaceBO")
    
    strResponse = objAdminIFBO.REQUEST(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
    
    ' Set task to complete
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
    Set xmlNode = xmlTempRequest.appendChild(xmlCaseTask.cloneNode(True))

    xmlSetAttributeValue xmlNode, "USERID", xmlGetMandatoryAttributeText(xmlTempRequest, "USERID")
    xmlSetAttributeValue xmlNode, "UNITID", xmlGetMandatoryAttributeText(xmlTempRequest, "UNITID")
    xmlSetAttributeValue xmlNode, "USERAUTHORITYLEVEL", xmlGetMandatoryAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlNode, "TASKSTATUS", GetFirstComboValueId("TaskStatus", "CP")
    
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True
        
RunGenericInterfaceExit:

    Set objAppManagerBO = Nothing
    Set objAdminIFBO = Nothing
    Set objMsgTmBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNode = Nothing
    Set xmlApplication = Nothing
    Set xmlCaseTask = Nothing
    Set xmlListNode = Nothing
    Set xmlCustomers = Nothing
    Set xmlCustomer = Nothing
    Set xmlTempDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
' PSC 21/10/2002 BMIDS00676 - End


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

'MO - 15/11/2002 - BMIDS00814 - START
Private Sub TransferApplicationOwnership(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
                           
On Error GoTo TransferApplicationOwnershipExit

Const cstrFunctionName As String = "TransferApplicationOwnership"
        
    Dim objAppBO As Object
    Dim objMsgTmBO As Object
    Dim xmlCreateUserHistoryDOM As FreeThreadedDOMDocument40
    Dim xmlUserHistoryNode As IXMLDOMNode
    Dim strCreateUserHistoryResponse As String
    Dim xmlNode As IXMLDOMNode
    Dim xmlCurrentStageResponse As FreeThreadedDOMDocument40
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlGetStageTaskDOM As FreeThreadedDOMDocument40
    Dim xmlStageTaskNode As IXMLDOMNode
    Dim xmlGetStageTaskResponse As FreeThreadedDOMDocument40
    Dim xmlStageTaskForCaseTaskId As IXMLDOMNode
    Dim xmlTransferTaskListDOM As FreeThreadedDOMDocument40
    Dim xmlTransferCaseTaskNode As IXMLDOMNode
    Dim xmlUpdateCaseTaskResponse As FreeThreadedDOMDocument40
    Dim strOldUserId As String
    Dim strOldUnitId As String
    Dim strNewUserId As String
    Dim strNewUnitId As String
    Dim strCurrentStageId As String
    Dim xmlNodeList As IXMLDOMNodeList  'BM0177 MDC 20/12/2002
            
    'extract values from the XML
    strNewUserId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@NEWUSERID")
    strNewUnitId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@NEWUNITID")
    strOldUserId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@OLDUSERID")
    strOldUnitId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@OLDUNITID")
            
    'BUILD THE CREATE USER HISTORY REQUEST
    Set xmlCreateUserHistoryDOM = New FreeThreadedDOMDocument40
    xmlCreateUserHistoryDOM.validateOnParse = False
    xmlCreateUserHistoryDOM.setProperty "NewParser", True
    'clone the request node
    xmlCreateUserHistoryDOM.appendChild vxmlRequestNode.cloneNode(False)
    Set xmlUserHistoryNode = xmlCreateUserHistoryDOM.createElement("USERHISTORY")
    xmlCreateUserHistoryDOM.documentElement.appendChild xmlUserHistoryNode
    
    Set xmlNode = xmlCreateUserHistoryDOM.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryNodeText(vxmlRequestNode, "CASEACTIVITY/@CASEID")
    xmlUserHistoryNode.appendChild xmlNode
    
    Set xmlNode = xmlCreateUserHistoryDOM.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryNodeText(vxmlRequestNode, "CASEACTIVITY/@CASEID")
    xmlUserHistoryNode.appendChild xmlNode
    
    Set xmlNode = xmlCreateUserHistoryDOM.createElement("USERID")
    xmlNode.Text = strNewUserId
    xmlUserHistoryNode.appendChild xmlNode
    
    Set xmlNode = xmlCreateUserHistoryDOM.createElement("UNITID")
    xmlNode.Text = strNewUnitId
    xmlUserHistoryNode.appendChild xmlNode
    
    'CALL CREATE USER HISTORY TO MODIFY THE OWNER OF THE CASE
    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    strCreateUserHistoryResponse = objAppBO.CreateUserHistory(xmlCreateUserHistoryDOM.xml)
    
    'CHECK THE RESPONSE RETURNED AND RAISE ERRORS
    errCheckXMLResponse strCreateUserHistoryResponse, True
    
    'GET THE CURRENT STAGE OF THE APPLICATION
    Set xmlCurrentStageResponse = New FreeThreadedDOMDocument40
    xmlCurrentStageResponse.validateOnParse = False
    xmlCurrentStageResponse.setProperty "NewParser", True
    GetCurrentStage vxmlRequestNode, xmlCurrentStageResponse
        
    strCurrentStageId = xmlGetMandatoryNodeText(xmlCurrentStageResponse, "RESPONSE/CASESTAGE/@STAGEID")
        
    'GET ALL THE STAGE TASKS FOR THE CURRENT STAGE
    'BUILD GETSTAGETASK REQUEST
    Set xmlGetStageTaskDOM = New FreeThreadedDOMDocument40
    xmlGetStageTaskDOM.validateOnParse = False
    xmlGetStageTaskDOM.setProperty "NewParser", True
    'close the request node
    xmlGetStageTaskDOM.appendChild vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlGetStageTaskDOM.documentElement, "OPERATION", "GETSTAGETASKLIST"
    Set xmlStageTaskNode = xmlGetStageTaskDOM.createElement("STAGETASK")
    xmlGetStageTaskDOM.documentElement.appendChild xmlStageTaskNode
    xmlSetAttributeValue xmlStageTaskNode, "STAGEID", strCurrentStageId
    
    'Get the stagetask for this case task
    Set xmlGetStageTaskResponse = New FreeThreadedDOMDocument40
    xmlGetStageTaskResponse.validateOnParse = False
    xmlGetStageTaskResponse.setProperty "NewParser", True
    RunTmRequest xmlGetStageTaskDOM, xmlGetStageTaskResponse
    
    'Create a DOM to hold the case tasks that are going to be transferred
    Set xmlTransferTaskListDOM = New FreeThreadedDOMDocument40
    xmlTransferTaskListDOM.validateOnParse = False
    xmlTransferTaskListDOM.setProperty "NewParser", True
    xmlTransferTaskListDOM.appendChild vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTransferTaskListDOM.documentElement, "OPERATION", "UPDATECASETASK"
    
    'BM0177 MDC 20/12/2002 - Only do UpdateCaseTask if caseTasks are found
    Set xmlNodeList = xmlCurrentStageResponse.selectNodes("/RESPONSE/CASESTAGE/CASETASK")
    If xmlNodeList.length > 0 Then
        'LOOP THROUGH ALL THE TASKS THAT REMAIN FOR THIS APPLICATION AND REASSIGN THEM
        'For Each xmlCaseTask In xmlCurrentStageResponse.selectNodes("/RESPONSE/CASESTAGE/CASETASK")
        For Each xmlCaseTask In xmlNodeList
        
            'is this task 'incomplete?
            'check validation type of 'I'
            If IsValidationType("TaskStatus", xmlGetAttributeAsLong(xmlCaseTask, "TASKSTATUS"), "I") = True Then
                
                'does this task belong to the old user?
                If xmlGetAttributeText(xmlCaseTask, "OWNINGUNITID") = strOldUnitId And xmlGetAttributeText(xmlCaseTask, "OWNINGUSERID") = strOldUserId Then
                    
                    'get the stage task for this case task from the get stage task list response we obtained earlier
                    Set xmlStageTaskForCaseTaskId = xmlGetMandatoryNode(xmlGetStageTaskResponse.documentElement, "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlCaseTask, "TASKID") & "']")
                    
                    'dont add to our list of tasks to transfer if its got a specific user id set against
                    ' the stage task or if its got a unit id and its different to the one we are trying to
                    ' set
                    If (xmlGetAttributeText(xmlStageTaskForCaseTaskId, "TASKUSERID") = "") Or (xmlGetAttributeText(xmlStageTaskForCaseTaskId, "TASKUNITID") <> strNewUnitId) Then
                        
                        Set xmlTransferCaseTaskNode = xmlCaseTask.cloneNode(True)
                        xmlSetAttributeValue xmlTransferCaseTaskNode, "OWNINGUSERID", strNewUserId
                        xmlSetAttributeValue xmlTransferCaseTaskNode, "OWNINGUNITID", strNewUnitId
                        xmlTransferTaskListDOM.documentElement.appendChild xmlTransferCaseTaskNode
                    
                    End If
                    
                End If
                
            End If
            
        Next
        
        'BM0177 MDC 13/01/2003
        If Not xmlGetNode(xmlTransferTaskListDOM, "//CASETASK") Is Nothing Then
            Set xmlUpdateCaseTaskResponse = New FreeThreadedDOMDocument40
            xmlUpdateCaseTaskResponse.validateOnParse = False
            xmlUpdateCaseTaskResponse.setProperty "NewParser", True
            RunTmRequest xmlTransferTaskListDOM, xmlUpdateCaseTaskResponse
        End If
        'BM0177 MDC 13/01/2003 - End
    End If
    'BM0177 MDC 20/12/2002 - End
    
TransferApplicationOwnershipExit:
        
    Set objAppBO = Nothing
    Set xmlCreateUserHistoryDOM = Nothing
    Set xmlUserHistoryNode = Nothing
    Set xmlNode = Nothing
    Set xmlCaseTask = Nothing
    Set xmlGetStageTaskDOM = Nothing
    Set xmlStageTaskNode = Nothing
    Set xmlGetStageTaskResponse = Nothing
    Set xmlTransferTaskListDOM = Nothing
    Set xmlTransferCaseTaskNode = Nothing
    Set xmlUpdateCaseTaskResponse = Nothing
    Set xmlNodeList = Nothing   'BM0177 MDC 20/12/2002
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'MO - 15/11/2002 - BMIDS00814 - END

' PSC 05/12/2002 BM0167 - Start
Private Sub UnapproveApplication(ByVal vxmlRequestNode As IXMLDOMNode, _
                                 ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UnapproveApplicationExit

Const cstrFunctionName As String = "UnapproveApplication"
        
    Dim objApplicationBO As Object
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    
    Dim strResponse As String

    ' Validate Request
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' Set the approval fields to null
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    Set xmlRoot = xmlTempDoc.createElement("APPLICATIONFACTFIND")
    xmlTempRequest.appendChild xmlRoot
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONAPPROVALDATE")
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONAPPROVALUSERID")
    xmlRoot.appendChild xmlNode
    Set xmlNode = xmlTempDoc.createElement("APPLICATIONAPPROVALUNITID")
    xmlRoot.appendChild xmlNode
    
    Set objApplicationBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    
    strResponse = objApplicationBO.UpdateApplicationFactFind(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
     
UnapproveApplicationExit:

    Set objApplicationBO = Nothing
    Set xmlTempRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlNode = Nothing
    Set xmlApplication = Nothing
    Set xmlCaseTask = Nothing
    Set xmlTempDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
' PSC 05/12/2002 BM0167 - End

'BMIDS670 DRC 11/12/2003
Private Function IsFreezeUnfreezeDataTask(ByVal strTaskKind As String, _
                                          ByVal strThisTaskId As String) As Boolean
On Error GoTo IsFreezeUnfreezeDataTaskExit
Dim ColValidIDs As New Collection
Const cstrFunctionName As String = "IsFreezeUnfreezeDataTask"
Dim thisValID As Integer
Dim blnReturn As Boolean

blnReturn = False
'Look up the Combo table to see if the Task belongs to the specified Group,
'and get it's valueID if it does
'
GetValueIdsForValueName "TMFreezeUnFreezeDataTask", strThisTaskId, ColValidIDs
' The task is in the Combo
If (ColValidIDs.Count > 0) Then
' A validation type of "F" means freeze ONLY
' An Unfreeze task, the default, unfreezes the data when incomplete,
' but freezes the data when it is complete. It has no Validation type

    If (strTaskKind = "UnFreeze") Then
        thisValID = ColValidIDs.Item(1)
        If (GetValidationTypeForValueID("TMFreezeUnFreezeDataTask", thisValID) <> "F") Then
            blnReturn = True
        End If
    Else
        blnReturn = True
    End If
End If

IsFreezeUnfreezeDataTask = blnReturn

IsFreezeUnfreezeDataTaskExit:

   Set ColValidIDs = Nothing
   errCheckError cstrFunctionName, TypeName(Me)
End Function
'BMIDS670 DRC 11/12/2003

' PSC 14/08/2005 MAR32 - Start
Private Sub FindCaseTaskListLite(ByVal vxmlRequestNode As IXMLDOMNode, _
                                 ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo FindCaseTaskListLiteExit
Const cstrFunctionName As String = "FindCaseTaskListLite"
                        
    
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xslDoc As FreeThreadedDOMDocument40
    Dim xmlTransformedDoc As FreeThreadedDOMDocument40
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTasks As IXMLDOMNodeList
    Dim xmlCaseTask As IXMLDOMNode
    Dim cintValueIds As Collection

    Dim strAttribValue As String
    Dim strDueDateStart As String
    Dim strDueDateEnd As String
    Dim strDateFilter As String
    Dim strSLAExpiryFilter As String
    Dim strFilter As String
    Dim strInClause As String
    Dim strTaskType As String
    Dim strSLAExpiry As String
    Dim strXslPath As String
    
    Dim dteSearchEndDate As Date

    Dim intIndex As Integer
    
    Dim blnLoadedOk As Boolean

    
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlRequestCaseTaskNode = xmlRequestCaseTaskNode.cloneNode(True)
    If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUESTARTDATEANDTIME") Then
        strAttribValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUESTARTDATEANDTIME")
        If Not IsDate(strAttribValue) Then
            errThrowError _
                cstrFunctionName, _
                oeXMLInvalidAttributeValue, _
                "TASKDUESTARTDATEANDTIME = " & strAttribValue & " (not a valid date)"
         End If
         strDueDateStart = strAttribValue
    End If
    
    If xmlAttributeValueExists(xmlRequestCaseTaskNode, "TASKDUEENDDATEANDTIME") Then
        strAttribValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKDUEENDDATEANDTIME")
        If Not IsDate(strAttribValue) Then
            errThrowError _
                cstrFunctionName, _
                oeXMLInvalidAttributeValue, _
                "TASKDUEENDDATEANDTIME = " & strAttribValue & " (not a valid date)"
         End If
        strDueDateEnd = strAttribValue
        ' Add a day to the date as we want to include the end date in search
        dteSearchEndDate = CSafeDate(strDueDateEnd)
        dteSearchEndDate = DateAdd("d", 1, dteSearchEndDate)
        strDueDateEnd = Format$(dteSearchEndDate, "DD/MM/YYYY")
    End If
    
    strSLAExpiry = xmlGetAttributeText(xmlRequestCaseTaskNode, "SLAEXPIRYWITHIN")
    strTaskType = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKTYPE")
    
    Set cintValueIds = New Collection
        
    If xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKSTATUS") = "I" Then
   
        xmlRequestCaseTaskNode.Attributes.removeNamedItem ("TASKSTATUS")
   
        GetValueIdsForValidationType "TaskStatus", "I", cintValueIds
    
        For intIndex = 1 To cintValueIds.Count
            If Len(strInClause) > 0 Then
                strInClause = strInClause & ","
            End If
            strInClause = strInClause & CStr(cintValueIds.Item(intIndex))
        Next
    
        strFilter = "TASKSTATUS IN (" & strInClause & ")"
    End If
    
    ' Set up condition for task due date start and end
    Select Case True
        Case Len(strDueDateStart) > 0 And Len(strDueDateEnd) > 0
            strDateFilter = "TASKDUEDATEANDTIME BETWEEN '" & strDueDateStart & _
                            "' AND '" & strDueDateEnd & "'"
        Case Len(strDueDateStart) > 0
             strDateFilter = "TASKDUEDATEANDTIME > '" & strDueDateStart & "'"
        
        Case Len(strDueDateEnd) > 0
             strDateFilter = "TASKDUEDATEANDTIME < '" & strDueDateEnd & "'"
    End Select
    
    If Len(strDateFilter) > 0 Then
        If Len(strFilter) > 0 Then
            strFilter = strFilter & " AND "
        End If
        
        strFilter = strFilter & strDateFilter
    End If
    
    ' Set up condition for SLA Expiry
    If Len(strSLAExpiry) > 0 Then
        ' PSC 26/10/2005 MAR284 - Start
        If CSafeInt(strSLAExpiry) >= 0 Then
            strSLAExpiryFilter = "SLAEXPIRYWITHIN >=0 AND SLAEXPIRYWITHIN <= " & strSLAExpiry
        Else
            strSLAExpiryFilter = "SLAEXPIRYWITHIN <=0 AND SLAEXPIRYWITHIN >= " & strSLAExpiry
        End If
        ' PSC 26/10/2005 MAR284 - End
        
        If Len(strFilter) > 0 Then
            strFilter = strFilter & " AND "
        End If
        
        strFilter = strFilter & strSLAExpiryFilter
        xmlRequestCaseTaskNode.Attributes.removeNamedItem ("SLAEXPIRYWITHIN")
    End If
        
    ' If we are not doing an all task types together with an SLA Expiry search process as normal
    ' otherwise create a temporary response node so that we can transform it into the
    ' correct xml format
    If Len(strTaskType) > 0 Or Len(strSLAExpiry) = 0 Then
        Set xmlTempResponseNode = vxmlResponseNode
    Else
       Set xmlTempResponseNode = vxmlResponseNode.ownerDocument.createElement("RESPONSE")
    End If
     
    adoGetAsXML xmlRequestCaseTaskNode, xmlTempResponseNode, "CASETASKDETAILLITE", strFilter, "TASKDUEDATEANDTIME"
    
    If xmlTempResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
     
    ' If we are doing an all task type and SLA expiry serach transform to just return
    ' unique case ids
    If Len(strTaskType) = 0 And Len(strSLAExpiry) > 0 Then
    
        Set xslDoc = New FreeThreadedDOMDocument40
        xslDoc.validateOnParse = False
        xslDoc.setProperty "NewParser", True
        xslDoc.async = False
        
        Set xmlTransformedDoc = New FreeThreadedDOMDocument40
        xmlTransformedDoc.validateOnParse = False
        xmlTransformedDoc.setProperty "NewParser", True
        
        Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
        xmlTempResponseDoc.validateOnParse = False
        xmlTempResponseDoc.setProperty "NewParser", True

        strXslPath = App.Path & "\FindCaseTaskListLite.xslt"
        strXslPath = Replace(strXslPath, "DLL", "XML", 1, 1, vbTextCompare)
        blnLoadedOk = xslDoc.Load(strXslPath)
    
        If Not blnLoadedOk Then
            errThrowError cstrFunctionName, oeSchemaNotLoaded, "Error in Loading XSLT File - " & strXslPath
        End If
        
        xmlTempResponseDoc.loadXML (xmlTempResponseNode.xml)
        
        xmlTempResponseDoc.transformNodeToObject xslDoc.documentElement, xmlTransformedDoc
        
        ' PSC 26/10/2005 MAR284
        Set xmlCaseTasks = xmlTransformedDoc.selectNodes("CASETASKS/CASETASK")
        
        For Each xmlCaseTask In xmlCaseTasks
            vxmlResponseNode.appendChild xmlCaseTask
        Next
    End If


FindCaseTaskListLiteExit:

    Set xmlRequestCaseTaskNode = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlTempResponseNode = Nothing
    Set xslDoc = Nothing
    Set xmlTransformedDoc = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlCaseTasks = Nothing
    Set xmlCaseTask = Nothing
    Set cintValueIds = Nothing
    
    gobjContext.SetComplete

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

' PSC 14/08/2005 MAR32 - Start
Private Sub ProcessTASTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo ProcessTASTaskExit
    Const cstrFunctionName As String = "ProcessTASTask"

    Dim objNTxBO As IOmTmNTxBO
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.ProcessTASTask vxmlRequestNode, vxmlResponseNode
    
ProcessTASTaskExit:
    Set objNTxBO = Nothing
    gobjContext.SetComplete

End Sub

Private Sub TASMoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo TASMoveToNextStageExit
    Const cstrFunctionName As String = "TASMoveToNextStage"

    Dim objNTxBO As IOmTmNTxBO
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.TASMoveToNextStage vxmlRequestNode, vxmlResponseNode
    
TASMoveToNextStageExit:
    
    Set objNTxBO = Nothing
    gobjContext.SetComplete

End Sub

Private Sub RunAutoApplicationExpiry(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo RunAutoApplicationExpiryErr
    Const cstrFunctionName As String = "RunAutoApplicationExpiry"

    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseStage As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim objNoTxBO As omTmNoTxBO
    Dim objBO As Object
    
    Dim strExceptionReason As String
    Dim strCompleteValue As String
    Dim strResponse As String
    Dim intCancelStage As Integer
    
    Set xmlCaseTaskIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    Set xmlTempResponse = New FreeThreadedDOMDocument40
    xmlTempResponse.validateOnParse = False
    xmlTempResponse.setProperty "NewParser", True
        
    strExceptionReason = GetFirstComboValueId("ExceptionReason", "CE")
    intCancelStage = GetMandatoryGlobalParamAmount("CancelledStageValue")
    
    ' Move to Stage
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStage"
    Set xmlCaseStage = xmlTempResponse.createElement("CASESTAGE")
    xmlTempRequest.appendChild xmlCaseStage
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "CASEID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYINSTANCE"
    xmlSetAttributeValue xmlCaseStage, "STAGEID", CStr(intCancelStage)
    xmlSetAttributeValue xmlCaseStage, "EXCEPTIONREASON", strExceptionReason
    xmlTempRequest.appendChild xmlApplication.cloneNode(True)
    Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
    strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
        
RunAutoApplicationExpiryExit:

    Set xmlCaseTask = Nothing
    Set xmlCaseTaskIn = Nothing
    Set xmlApplication = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseStage = Nothing
    Set xmlTempResponse = Nothing
    Set objNoTxBO = Nothing
    Set objBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RunAutoApplicationExpiryErr:

    gobjContext.SetAbort
    GoTo RunAutoApplicationExpiryExit

End Sub
' PSC 14/08/2005 MAR32 - End
'SR 21/10/2005 : MAR268 - new method 'RunHunterExperianInterface'
Private Sub RunHunterExperianInterface(ByVal vxmlRequestNode As IXMLDOMNode, _
                                       ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName  As String = "RunHunterExperianInterface"
    
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlTemp             As IXMLDOMElement
    Dim objMsgTmBO          As Object
    
    Dim objExperianHI As Object
    Dim strApplNumber As String
    Dim strAFFNumber As String
    Dim strTaskStatus As String
    Dim strResponse As String
    Dim lngErrNo As Long

    'MAR339
    Dim xmlHUNONode         As IXMLDOMNode
    Dim xmlERR1Node         As IXMLDOMNode
    Dim strErrorCode        As String
    Dim strErrorMessage     As String
    Dim strMessage          As String
    Dim strReturnCode       As String
    
On Error GoTo RunHunterExperianInterfaceError

    'Initialise
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'MAR339
    'Initialise the task status to Incomplete.
    'Status will be set to Complete when a successful Experian Hunter call is made.
    strTaskStatus = GetFirstComboValueId("TaskStatus", "I") 'Task Incomplete
    
    'Validate
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    strApplNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
    
    Set objExperianHI = gobjContext.CreateInstance("omExperianHI.ExperianHIBO")
    strResponse = objExperianHI.RunExperianHI(xmlTempRequest.xml)
    
    'MAR339 Check for errors
    'Check for an ErrorCode in the response and raise an error if found.
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    'A call to Experian will result in a HUNO node or an ERR1 node.
    'Check for the error block and raise an error if found.
    
    xmlTempDoc.loadXML strResponse

    Set xmlHUNONode = xmlTempDoc.selectSingleNode(".//HUNO")
    Set xmlERR1Node = xmlTempDoc.selectSingleNode(".//ERR1")
    
    If Not xmlERR1Node Is Nothing Then
        strErrorCode = xmlERR1Node.selectSingleNode(".//CODE").Text
        strMessage = xmlERR1Node.selectSingleNode(".//MESSAGE").Text
        strErrorMessage = ""
        
        If (strErrorCode <> "") Then
            strErrorMessage = "Error Code " & strErrorCode & ". "
        End If
       
        If (strMessage <> "") Then
            strErrorMessage = strErrorMessage & "Error Message: " & strMessage
        End If
        
        errThrowError cstrFunctionName, 7030, strErrorMessage
    
    End If
    
    'Check the return code in the HUNO block and raise an error if necessary.
    If Not xmlHUNONode Is Nothing Then
        'MAR533 RETURNCODE is returned as mixed case
        strReturnCode = xmlHUNONode.selectSingleNode(".//ReturnCode").Text
        
        If (strReturnCode = "9999") Then
            'Raise error and do not complete the task (status will stay at Incomplete)
            errThrowError cstrFunctionName, 7001
        
        ElseIf (strReturnCode = "0000") Then
        
            'Success - set the task to Complete
            strTaskStatus = GetFirstComboValueId("TaskStatus", "CP")  'Task complete
        
        End If
   
    End If
    
    'Now Update the CaseTask
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
    
    Set xmlTemp = xmlCaseTaskNode.cloneNode(True)
    xmlSetAttributeValue xmlTemp, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlTemp, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlTemp, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlTemp, "TASKSTATUS", strTaskStatus
    xmlTempRequest.appendChild xmlTemp
    
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
RunHunterExperianInterfaceExit:
    Set xmlTempDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTemp = Nothing
    Set xmlHUNONode = Nothing
    Set xmlERR1Node = Nothing
    
    Set objExperianHI = Nothing
    Set objMsgTmBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RunHunterExperianInterfaceError:

    gobjContext.SetAbort
    
    GoTo RunHunterExperianInterfaceExit

End Sub

'MAR184 GHun
'MAR1617 added UPDATELIM flag
Private Sub RunXMLCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMElement, _
                                Optional ByVal blnIsRescore As Boolean = False, _
                                Optional ByVal blnIsReprocess As Boolean = False, _
                                Optional ByVal blnUpdateLIM As Boolean = False)

    Const cstrFunctionName  As String = "RunXMLCreditCheck"
        
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim strResponse         As String
    Dim xmlHeaderNode       As IXMLDOMNode
    Dim xmlTargetReqNode    As IXMLDOMNode
    Dim xmlCCN1Node         As IXMLDOMNode
    Dim bAddressTargeting   As Boolean
    Dim xmlNewRequest       As IXMLDOMNode
    Dim objMsgTmBO          As Object
    Dim objExperianCC       As Object
    Dim blnLIM_Rescore      As Boolean
    Dim strStageId          As String    'MAR630

On Error GoTo RunXMLCreditCheckExit

    'Initialise ======================================================================
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'Validate ========================================================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' PSC 04/11/2005 MAR329
    Set xmlCaseTaskNode = xmlGetNode(vxmlRequestNode, "CASETASK")
    
    ' MAR245 - Check if request includes address targeting data.
    Set xmlTargetReqNode = xmlGetNode(vxmlRequestNode, "TARGETINGDATA")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    
    ' PSC 04/11/2005 MAR329 - Start
    If Not xmlCaseTaskNode Is Nothing Then
        xmlTempRequest.appendChild xmlCaseTaskNode.cloneNode(True)
    End If
    ' PSC 04/11/2005 MAR329 - End
    
    ' ik_MAR452_20051111
    ' xmlTempRequest.appendChild xmlApplicationNode
    xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
    ' ik_MAR452_20051111_ends
    
    ' Add our Address Target Request if/when appropriate
    If Not xmlTargetReqNode Is Nothing Then
        xmlTempRequest.appendChild xmlTargetReqNode
    End If
    
    If blnIsReprocess Then
        xmlTempRequest.setAttribute "REPROCESS", "1"
    End If
    If blnIsRescore Then
        xmlTempRequest.setAttribute "RESCORE", "1"
    End If
    
    
    'MAR841 Set the operation to RunXMLCreditCheck - flags are used for Reprocess and Rescore
    xmlTempRequest.setAttribute "OPERATION", "RunXMLCreditCheck"
    
    'DRC - Check for LIM Rescore task
    'MAR630
    If xmlCaseTaskNode Is Nothing Then
        blnLIM_Rescore = False
    Else
        If UCase(xmlGetAttributeText(xmlCaseTaskNode, "TASKID")) = _
            UCase(GetGlobalParamString("TMLIMRescoreID")) Then
            blnLIM_Rescore = True
        Else
            blnLIM_Rescore = False
        End If
    End If
    
    'MAR1617 add updateLIM flag
    If blnUpdateLIM Then
        xmlTempRequest.setAttribute "LIMUPDATEONLY", "1"
        blnLIM_Rescore = True
    End If
    
    'Do Credit Score =============================================
    Set objExperianCC = gobjContext.CreateInstance("omExperianCC.ExperianCCBO")
    strResponse = objExperianCC.Execute(xmlTempRequest.xml)
    
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
   
    AttachResponseData vxmlResponseNode, xmlTempDoc.documentElement
     
    bAddressTargeting = xmlGetNodeAsBoolean(vxmlResponseNode, ".//ADDRESSTARGETING")
    'If Address Targeting is required don't save at this time
    If Not (bAddressTargeting) Then
    
        'If we've got this far, need to Update the Case Task status
        'Build the Request and update the case task
        ' PSC 04/11/2005 MAR329 - Start
        If Not xmlCaseTaskNode Is Nothing Then
            Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
            xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(xmlTempRequest, "USERID")
            xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(xmlTempRequest, "UNITID")
            xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
            xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)
    
            Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
            strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
            errCheckXMLResponse strResponse, True
        End If
        ' PSC 04/11/2005 MAR329 - End

        
        'Call Risk Assessment
        'DRC MAR33 - need to add a context so that RA tasks will be added
        'MAR630
        If Not xmlCaseTaskNode Is Nothing Then                               'MAR630
            xmlSetAttributeValue vxmlRequestNode, "TASKMANAGER", "T"
            
            strStageId = xmlGetAttributeText(xmlCaseTaskNode, "STAGEID")     'MAR630
            xmlSetAttributeValue xmlApplicationNode, "STAGEID", strStageId   'MAR630
        End If                                                               'MAR630
        
        'DRC MAR599 - don't run Case Assessment for  LIM Rescore
        If Not blnLIM_Rescore Then
            RunRiskAssessment vxmlRequestNode, vxmlResponseNode
        End If
        'DRC MAR599
    Else
        'MAR1084 GHun Throw an error if address targeting is required, but not supported
        If xmlGetAttributeAsBoolean(vxmlRequestNode, "NOADDRESSTARGETING") Then
            errThrowError cstrFunctionName, oeTmAddressTargetingNotSupported
        End If
        'MAR1084 End
        
        ' MAR245
        ' Flag if a recore or reprocess was requested.
        If blnIsReprocess Then
           vxmlResponseNode.setAttribute "REPROCESS", "1"
        End If
        If blnIsRescore Then
            vxmlResponseNode.setAttribute "RESCORE", "1"
        End If
        If blnUpdateLIM Then
            vxmlResponseNode.setAttribute "LIMUPDATEONLY", "1"
        End If
        
        'Copy the CaseTask node from the request to the response, when address targeting
        If Not xmlCaseTaskNode Is Nothing Then  'MAR1084 GHun CaseTask node does not always exist
            vxmlResponseNode.appendChild xmlCaseTaskNode.cloneNode(True)
        End If
    End If
    
RunXMLCreditCheckExit:
   
    Set objExperianCC = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlNewRequest = Nothing
    Set objMsgTmBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'MAR184 End

'MAR23 Add KnowYourCustomer
Private Sub KnowYourCustomer(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    KnowYourCustomer -- MAR23
' author:       Tony Liu
' date:         07/09/2005
' description:  Performes a money laundering check against the customer and saves the results
'               in Omiga. If the process is successful, the Run KYC task will be completed.
'
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       vxmlResponseNode
'               XML RESPONSE node
'------------------------------------------------------------------------------------------

On Error GoTo KnowYourCustomerError
    
    Const cstrFunctionName         As String = "KnowYourCustomer"
        
    Dim xmlDoc                     As FreeThreadedDOMDocument40
    Dim xmlApplicationNode         As IXMLDOMNode
    Dim xmlCaseTaskNode            As IXMLDOMNode
    
    Dim xmlTempRequest             As IXMLDOMElement
    Dim xmlElem                    As IXMLDOMElement
    Dim xmlCaseTaskElem            As IXMLDOMElement
    Dim xmlNewRequest              As IXMLDOMNode
    Dim xmlKYCTaskRequest          As IXMLDOMNode
    Dim xmlDGRequestNode           As IXMLDOMNode
    Dim xmlKYCResponseNode         As IXMLDOMNode
    Dim xmlCustomerNode            As IXMLDOMNode
    Dim xmlTempRootElem            As IXMLDOMNode
    Dim strAppNo                   As String
    Dim strAFFNo                   As String
    Dim strCustomerNo              As String
    Dim strCustomerVersionNo       As String
    Dim strOtherSystemCustomerNo   As String
    Dim strFirstForename           As String
    Dim strSurname                 As String
    Dim strCustomerName            As String
    Dim strTaskName                As String
    Dim strUserId                  As String

    Dim objCustomerBO              As Object
    Dim objMsgTmBO                 As Object
    
    Dim strCustomerNumberResponse  As String
    Dim strResponse                As String
    Dim strCustomerKYCStatus       As String
    Dim strCustomerKYCAddressFlag  As String
    Dim strCustomerKYCIDFlag       As String
    Dim strTaskId                  As String
    Dim strKYCResponse             As String
    
    Dim strCaseId                  As String
    Dim strActivityId              As String
    Dim strActivityInstance        As String
    Dim strCaseStageSequenceNumber As String
    Dim strStageId                 As String
    
    Dim strStatus                  As String

    Dim objNoTxBO                  As omTmNoTxBO    'EP974/MAR1848
    
    'Initialise ======================================================================
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    strUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
    strCaseId = xmlGetAttributeText(xmlCaseTaskNode, "CASEID")
    strActivityId = xmlGetAttributeText(xmlCaseTaskNode, "ACTIVITYID")
    strActivityInstance = xmlGetAttributeText(xmlCaseTaskNode, "ACTIVITYINSTANCE")
    strCaseStageSequenceNumber = xmlGetAttributeText(xmlCaseTaskNode, "CASESTAGESEQUENCENO")
    strStageId = xmlGetAttributeText(xmlCaseTaskNode, "STAGEID")
    
    'Get and Synchronise Customer Details
    'Build Request and Get Customer data
    strAppNo = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strAFFNo = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    strCustomerNo = xmlGetAttributeText(xmlCaseTaskNode, "CUSTOMERIDENTIFIER")
    
    'Set objCustomerBO = gobjContext.CreateInstance("omCust.CustomerBO")
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strAppNo
    xmlNewRequest.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElem.Text = strCustomerNo
    xmlNewRequest.appendChild xmlElem
    
    'EP974/MAR1848 Call GetCustomerNumbers in a separate transaction
    'Set objCustomerBO = gobjContext.CreateInstance("omCust.CustomerBO")
    'strCustomerNumberResponse = objCustomerBO.GetCustomerNumbers(xmlNewRequest.xml)
    Set objNoTxBO = gobjContext.CreateInstance(gstrTASKMANAGEMENT_COMPONENT & ".OmTmNoTxBO")
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETCUSTOMERNUMBERS"
    strCustomerNumberResponse = objNoTxBO.OmTmNoTxRequest(xmlNewRequest.xml)
    'EP974/MAR1848 End

    errCheckXMLResponse strCustomerNumberResponse, True
    
    xmlDoc.loadXML (strCustomerNumberResponse)
    
    xmlDoc.selectSingleNode ("//RESPONSE/CUSTOMER")
    
    strCustomerVersionNo = xmlDoc.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    strOtherSystemCustomerNo = xmlDoc.selectSingleNode(".//OTHERSYSTEMCUSTOMERNUMBER").Text
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "DISTRIBUTIONCHANNELID", xmlGetAttributeText(xmlNewRequest, "CHANNELID")
    
    Set xmlTempRootElem = xmlDoc.createElement("SEARCH")
    xmlNewRequest.appendChild xmlTempRootElem
    
    Set xmlTempRequest = xmlDoc.createElement("CUSTOMER")
    ' MAR1477 Getandsync was locking the customer and then not unlocking it.
    xmlTempRequest.setAttribute "NOLOCK", "1"
    xmlTempRequest.appendChild(xmlDoc.createElement("CUSTOMERNUMBER")).Text = strCustomerNo
    xmlTempRequest.appendChild(xmlDoc.createElement("CUSTOMERVERSIONNUMBER")).Text = strCustomerVersionNo
    xmlTempRequest.appendChild(xmlDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")).Text = strOtherSystemCustomerNo
    xmlTempRootElem.appendChild xmlTempRequest
    
    'EP974/MAR1848 Call GetAndSynchroniseCustomerDetails in a separate transaction
    'strResponse = objCustomerBO.GetAndSynchroniseCustomerDetails(xmlNewRequest.xml)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETANDSYNCHRONISECUSTOMERDETAILS"
    strResponse = objNoTxBO.OmTmNoTxRequest(xmlNewRequest.xml)
    'EP974/MAR1848 End

    errCheckXMLResponse strResponse, True
    
    xmlDoc.loadXML (strResponse)

    Dim objKnowYourCustomer As Object   'MAR231 GHun changed to late binding as per omTM standard
    Set objKnowYourCustomer = gobjContext.CreateInstance("omKnowYourCustomer.KnowYourCustomerBO")
    
    strFirstForename = xmlDoc.selectSingleNode(".//FIRSTFORENAME").Text
    strSurname = xmlDoc.selectSingleNode(".//SURNAME").Text
    strCustomerName = strFirstForename & " " & strSurname
    strCustomerKYCStatus = xmlDoc.selectSingleNode(".//CUSTOMERKYCSTATUS").Text
    strCustomerKYCAddressFlag = xmlDoc.selectSingleNode(".//CUSTOMERKYCADDRESSFLAG").Text
    strCustomerKYCIDFlag = xmlDoc.selectSingleNode(".//CUSTOMERKYCIDFLAG").Text
    
    If strCustomerKYCStatus = "Satisfied" Then
        ' Do nothing
    Else
        ' Create Case Task node for ad hoc tasks
        Set xmlCaseTaskElem = xmlDoc.createElement("CASETASK")
        xmlSetAttributeValue xmlCaseTaskElem, "SOURCEAPPLICATION", "Omiga"
        xmlSetAttributeValue xmlCaseTaskElem, "CASEID", strCaseId
        xmlSetAttributeValue xmlCaseTaskElem, "ACTIVITYID", strActivityId
        xmlSetAttributeValue xmlCaseTaskElem, "ACTIVITYINSTANCE", strActivityInstance
        xmlSetAttributeValue xmlCaseTaskElem, "CASESTAGESEQUENCENO", strCaseStageSequenceNumber
        xmlSetAttributeValue xmlCaseTaskElem, "STAGEID", strStageId
        
        ' Generate Task Request
        Set xmlKYCTaskRequest = xmlGetRequestNode(vxmlRequestNode)
        Set xmlElem = xmlDoc.createElement("TASK")
        xmlKYCTaskRequest.appendChild xmlElem

        Set xmlElem = xmlDoc.createElement("CUSTOMER")
        xmlSetAttributeValue xmlElem, "CUSTOMERNAME", strCustomerName
        xmlKYCTaskRequest.appendChild xmlElem
        
        xmlKYCTaskRequest.appendChild xmlCaseTaskElem
        xmlKYCTaskRequest.appendChild xmlApplicationNode
        
        If strCustomerKYCStatus = "Pending" Then
            'Generate ID and Address tasks if the flags are set to Outstanding
            
            If (strCustomerKYCAddressFlag = "Outstanding") Then
            
                'Generate KYC_AddressFlag task
                strTaskId = GetMandatoryGlobalParamString("TMKYCAddressFlag")
                xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
                CreateKYCTask xmlKYCTaskRequest
      
            End If
        
            If (strCustomerKYCIDFlag = "Outstanding") Then
            
                'Generate KYC_IDFlag task
                strTaskId = GetMandatoryGlobalParamString("TMKYCIdFlag")
                xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
                CreateKYCTask xmlKYCTaskRequest
            
            End If
        
        ElseIf strCustomerKYCStatus = "Not Done" Then
                
            'Generate KYC_AddressFlag task
            strTaskId = GetMandatoryGlobalParamString("TMKYCAddressFlag")
            xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask xmlKYCTaskRequest
                
            'Generate KYC_IDFlag task
            strTaskId = GetMandatoryGlobalParamString("TMKYCIdFlag")
            xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask xmlKYCTaskRequest
              
        ElseIf strCustomerKYCStatus = "Under Review" Then
        
            'Generate KYC_ProcReferral task
            strTaskId = GetMandatoryGlobalParamString("TMKYCProcessorReferral")
            xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask xmlKYCTaskRequest
                
        ElseIf strCustomerKYCStatus = "Unsatisfied" Then
            
            'Call Experian KYC Interface (C# COM omKnowYourCustomer)
            'Build up the XML to send to the Direct Gateway
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    
            Set xmlCustomerNode = xmlDoc.createElement("CUSTOMER")
            xmlSetAttributeValue xmlCustomerNode, "CUSTOMERNUMBER", strCustomerNo
            xmlSetAttributeValue xmlCustomerNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNo
            xmlSetAttributeValue xmlCustomerNode, "OTHERSYSTEMCUSTOMERNUMBER", strOtherSystemCustomerNo 'MAR467 GHun
            xmlNewRequest.appendChild xmlCustomerNode
            
            'Send request
            strKYCResponse = objKnowYourCustomer.RUNKNOWYOURCUSTOMER(xmlNewRequest.xml)
            
            '***************************************************
            'TESTING
            
'            strKYCResponse = "<RESPONSE><KycDetails>"
'            strKYCResponse = strKYCResponse & "<Status>Satisfied</Status>"
'            strKYCResponse = strKYCResponse & "<IdFlag>IDText</IdFlag>"
'            strKYCResponse = strKYCResponse & "<AddressFlag>AddressText</AddressFlag>"
'            strKYCResponse = strKYCResponse & "</KycDetails></RESPONSE>"
            
            '*************************************************
            
            xmlDoc.loadXML strKYCResponse
            
            errCheckXMLResponse strKYCResponse, True
            
            'Process the response
            Set xmlKYCResponseNode = xmlDoc.selectSingleNode(".//KycDetails")

            'If a response has been returned, save the details and create Ad Hoc tasks as required
            If Not xmlKYCResponseNode Is Nothing Then

                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlNewRequest.appendChild xmlCustomerNode
                xmlNewRequest.appendChild xmlKYCResponseNode

                DecodeKYCCheckResponse xmlNewRequest, xmlKYCTaskRequest
                
                strStatus = "1"
                
            Else
                'An error has been returned
                strStatus = "0"
                
            End If
            
            'Create an entry in the KnowYourCustomerCheck table
            
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
               
            Set xmlCustomerNode = xmlDoc.createElement("KNOWYOURCUSTOMERCHECK")
            xmlNewRequest.appendChild xmlCustomerNode
    
            Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlElem.Text = strCustomerNo
            xmlCustomerNode.appendChild xmlElem
   
            Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlElem.Text = strCustomerVersionNo
            xmlCustomerNode.appendChild xmlElem

            Set xmlElem = xmlDoc.createElement("USERID")
            xmlElem.Text = strUserId
            xmlCustomerNode.appendChild xmlElem
   
            Set xmlElem = xmlDoc.createElement("SUCCESSINDICATOR")
            xmlElem.Text = strStatus
            xmlCustomerNode.appendChild xmlElem
            
            Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")  'EP974/MAR1848
            strResponse = objCustomerBO.CreateKYCCheck(xmlNewRequest.xml)
            errCheckXMLResponse strResponse, True
                                    
            'Generate KYC_UW_Referral task
            strTaskId = GetMandatoryGlobalParamString("TMKYCUnderwriterReferral")
            xmlSetAttributeValue xmlKYCTaskRequest.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask xmlKYCTaskRequest
            
        Else ' CustomerKYCStatus = "Expired" or "Null"
        
            ' Call Experian KYC Interface (C# COM omKnowYourCustomer)
            'Build up the XML to send to the Direct Gateway
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            
            Set xmlCustomerNode = xmlDoc.createElement("CUSTOMER")
            xmlSetAttributeValue xmlCustomerNode, "CUSTOMERNUMBER", strCustomerNo
            xmlSetAttributeValue xmlCustomerNode, "CUSTOMERVERSIONNUMBER", strCustomerVersionNo
            xmlSetAttributeValue xmlCustomerNode, "OTHERSYSTEMCUSTOMERNUMBER", strOtherSystemCustomerNo 'MAR467
            xmlNewRequest.appendChild xmlCustomerNode
            
            'Send request
            strKYCResponse = objKnowYourCustomer.RUNKNOWYOURCUSTOMER(xmlNewRequest.xml)
            xmlDoc.loadXML strKYCResponse
            
            errCheckXMLResponse strKYCResponse, True
            
            'Process the response
            Set xmlKYCResponseNode = xmlDoc.selectSingleNode(".//KycDetails")

            'If a response has been returned, save the details and create Ad Hoc tasks as required
            If Not xmlKYCResponseNode Is Nothing Then

                Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlNewRequest.appendChild xmlCustomerNode
                xmlNewRequest.appendChild xmlKYCResponseNode

                DecodeKYCCheckResponse xmlNewRequest, xmlKYCTaskRequest
                
                strStatus = "1"
                
            Else
                'An error has been returned
                strStatus = "0"
                
            End If
            
            'Create an entry in the KnowYourCustomerCheck table
            
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
               
            Set xmlCustomerNode = xmlDoc.createElement("KNOWYOURCUSTOMERCHECK")
            xmlNewRequest.appendChild xmlCustomerNode
    
            Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
            xmlElem.Text = strCustomerNo
            xmlCustomerNode.appendChild xmlElem
   
            Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
            xmlElem.Text = strCustomerVersionNo
            xmlCustomerNode.appendChild xmlElem

            Set xmlElem = xmlDoc.createElement("USERID")
            xmlElem.Text = strUserId
            xmlCustomerNode.appendChild xmlElem
   
            Set xmlElem = xmlDoc.createElement("SUCCESSINDICATOR")
            xmlElem.Text = strStatus                                        ' MAR467
            xmlCustomerNode.appendChild xmlElem
   
            Set objCustomerBO = gobjContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")  'EP974/MAR1848
            strResponse = objCustomerBO.CreateKYCCheck(xmlNewRequest.xml)
            errCheckXMLResponse strResponse, True
           
        End If

    End If
    
    'Update the task as completed.
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
    xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")

    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    gobjContext.SetComplete ' AS 31/03/2006 MAR1507 PERFORMANCE - SetComplete is done in the wrong place for KnowYourCustomer.
    
KnowYourCustomerExit:
   
    Set xmlDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlElem = Nothing
    Set xmlCaseTaskElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlKYCTaskRequest = Nothing
    Set xmlDGRequestNode = Nothing
    Set xmlKYCResponseNode = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    Set objNoTxBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
KnowYourCustomerError:

    gobjContext.SetAbort
    
    GoTo KnowYourCustomerExit
    
End Sub

Private Sub CreateKYCTask(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo CreateKYCTaskError

    Const cstrFunctionName As String = "CreateKYCTask"
    
    Dim xmlDoc             As FreeThreadedDOMDocument40
    Dim xmlNewRequest      As IXMLDOMNode
    Dim xmlTaskOut         As IXMLDOMNode
    Dim xmlCaseTaskNode    As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlElem            As IXMLDOMElement
    Dim strResponse        As String
    Dim strTaskId          As String
    Dim strCustomerName    As String
    Dim strTaskName        As String
    Dim objMsgTmBO         As Object
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    
    ' PSC 02/12/2005 MAR770 - Start
    Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("CASETASK")
    
    If Not xmlCaseTaskNode Is Nothing Then
        Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True)
    End If
    
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("APPLICATION")
    
    If Not xmlApplicationNode Is Nothing Then
        Set xmlApplicationNode = xmlApplicationNode.cloneNode(True)
    End If
    ' PSC 02/12/2005 MAR770 - End
    
    strTaskId = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("TASK"), "TASKID")
    strCustomerName = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("CUSTOMER"), "CUSTOMERNAME")
    
    'Get details for this task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "GETTASKDETAIL"
    Set xmlElem = xmlDoc.createElement("TASK")
    xmlSetAttributeValue xmlElem, "TASKID", strTaskId
    xmlNewRequest.appendChild xmlElem
                
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
                
    xmlDoc.loadXML (strResponse)

    Set xmlTaskOut = xmlDoc.selectSingleNode("//TASK")
          
    'Create Adhoc Case Task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEADHOCCASETASK"
    xmlSetAttributeValue xmlCaseTaskNode, "TASKID", strTaskId
                
    strTaskName = xmlGetAttributeText(xmlTaskOut, "TASKNAME") & " for " & strCustomerName
    xmlSetAttributeValue xmlCaseTaskNode, "CASETASKNAME", strTaskName
                
    ' PSC 02/12/2005 MAR770 - End
    ' PSC 01/12/2005 MAR770 - Start
    xmlNewRequest.appendChild xmlCaseTaskNode
    xmlNewRequest.appendChild xmlApplicationNode
    ' PSC 01/12/2005 MAR770 - End
    ' PSC 02/12/2005 MAR770 - End
        
    CreateAdhocCaseTask xmlNewRequest

CreateKYCTaskExit:
    
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTaskOut = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlElem = Nothing
    Set objMsgTmBO = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)

    Exit Sub

CreateKYCTaskError:

    gobjContext.SetAbort
    
    GoTo CreateKYCTaskExit

End Sub
  
Private Sub DecodeKYCCheckResponse(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlTaskRequestNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' procedure:    DecodeKYCCheckResponse
' description:  Creates necessary Ad Hoc tasks.
'               Updates CustomerVersion and KYCCheckTable with data received from a successful
'               call to Experian.
' pass:         vxmlRequestNode
'               <KycDetails>
'                   <Status>
'                   <IdFlag>
'                   <AddressFlag>
'               </KycDetails>
'
'------------------------------------------------------------------------------------------

    On Error GoTo DecodeKYCCheckResponseError

    Const cstrFunctionName As String = "DecodeKYCCheckResponse"
        
    Dim xmlDoc                      As FreeThreadedDOMDocument40
    Dim xmlElem                     As IXMLDOMElement
    Dim xmlCustomerVersionElem      As IXMLDOMElement
    Dim xmlNewRequest               As IXMLDOMNode
    Dim strResponse                 As String
    Dim strCustomerNumber           As String
    Dim strCustomerVersionNumber    As String
    Dim strKYCStatus                As String
    Dim strKYCAddressFlag           As String
    Dim strKYCIDFlag                As String
    Dim strTaskId                   As String

    Dim objCustomerBO               As Object

    'Initialise ======================================================================
    
    Set objCustomerBO = gobjContext.CreateInstance("omCust.CustomerBO")

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    strCustomerNumber = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("CUSTOMER"), "CUSTOMERNUMBER")
    strCustomerVersionNumber = xmlGetAttributeText(vxmlRequestNode.selectSingleNode("CUSTOMER"), "CUSTOMERVERSIONNUMBER")
    strKYCStatus = vxmlRequestNode.selectSingleNode(".//Status").Text
    strKYCAddressFlag = vxmlRequestNode.selectSingleNode(".//AddressFlag").Text
    strKYCIDFlag = vxmlRequestNode.selectSingleNode(".//IdFlag").Text
        
    'Update the Customer Version table
                
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
               
    Set xmlCustomerVersionElem = xmlDoc.createElement("CUSTOMERVERSION")
    xmlNewRequest.appendChild xmlCustomerVersionElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElem.Text = strCustomerNumber
    xmlCustomerVersionElem.appendChild xmlElem
   
    Set xmlElem = xmlDoc.createElement("CUSTOMERVERSIONNUMBER")
    xmlElem.Text = strCustomerVersionNumber
    xmlCustomerVersionElem.appendChild xmlElem

    Set xmlElem = xmlDoc.createElement("CUSTOMERKYCSTATUS")
    xmlElem.Text = strKYCStatus
    xmlCustomerVersionElem.appendChild xmlElem
   
    Set xmlElem = xmlDoc.createElement("CUSTOMERKYCADDRESSFLAG")
    xmlElem.Text = strKYCAddressFlag
    xmlCustomerVersionElem.appendChild xmlElem
   
    Set xmlElem = xmlDoc.createElement("CUSTOMERKYCIDFLAG")
    xmlElem.Text = strKYCIDFlag
    xmlCustomerVersionElem.appendChild xmlElem

    strResponse = objCustomerBO.UpdateCustomerVersion(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    'Create Ad Hoc tasks as required
    If strKYCStatus = "Satisfied" Then
        ' Do nothing
    ElseIf strKYCStatus = "Pending" Then
        'Generate ID and Address tasks if the flags are set to Outstanding
            
        If (strKYCAddressFlag = "Outstanding") Then
            
            'Generate KYC_AddressFlag task
            strTaskId = GetMandatoryGlobalParamString("TMKYCAddressFlag")
            xmlSetAttributeValue vxmlTaskRequestNode.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask vxmlTaskRequestNode
      
        End If
        
        If (strKYCIDFlag = "Outstanding") Then
            
            'Generate KYC_IDFlag task
            strTaskId = GetMandatoryGlobalParamString("TMKYCIdFlag")
            xmlSetAttributeValue vxmlTaskRequestNode.selectSingleNode("TASK"), "TASKID", strTaskId
            CreateKYCTask vxmlTaskRequestNode
            
        End If
        
    ElseIf strKYCStatus = "Not Done" Then
                
        'Generate KYC_AddressFlag task
        strTaskId = GetMandatoryGlobalParamString("TMKYCAddressFlag")
        xmlSetAttributeValue vxmlTaskRequestNode.selectSingleNode("TASK"), "TASKID", strTaskId
        CreateKYCTask vxmlTaskRequestNode
                
        'Generate KYC_IDFlag task
        strTaskId = GetMandatoryGlobalParamString("TMKYCIdFlag")
        xmlSetAttributeValue vxmlTaskRequestNode.selectSingleNode("TASK"), "TASKID", strTaskId
        CreateKYCTask vxmlTaskRequestNode
              
    ElseIf strKYCStatus = "Under Review" Then
        
        'Generate KYC_ProcReferral task
        strTaskId = GetMandatoryGlobalParamString("TMKYCProcessorReferral")
        xmlSetAttributeValue vxmlTaskRequestNode.selectSingleNode("TASK"), "TASKID", strTaskId
        CreateKYCTask vxmlTaskRequestNode
    End If

    gobjContext.SetComplete ' AS 31/03/2006 MAR1507 PERFORMANCE - SetComplete is done in the wrong place for KnowYourCustomer.

DecodeKYCCheckResponseExit:
   
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlCustomerVersionElem = Nothing
    Set xmlNewRequest = Nothing
    Set objCustomerBO = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)

    Exit Sub

DecodeKYCCheckResponseError:
    
    gobjContext.SetAbort
    
    GoTo DecodeKYCCheckResponseExit

End Sub

'SR 04/11/2005 : MAR329
'MAR750  Changes for interface call
Private Sub RunLegalAddressTriggerInterface(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo RunLegalAddressTriggerInterfaceExit

    Const cstrFunctionName As String = "RunLegalAddressTriggerInterface"
    
    Dim xmlApplication As IXMLDOMElement
    Dim xmlCaseTask As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlApp As IXMLDOMElement
    Dim objAppBO As Object
    Dim objInterfaceBO As Object
    Dim strResponse As String
    Dim lngError As Long
    Dim strApplNo As String
    Dim strAFFNo As String
    Dim strOtherSysAccountNumber As String
    
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlDoc = New FreeThreadedDOMDocument40
   
    'Build Request and Get Application data
    strApplNo = xmlGetAttributeText(xmlApplication, "APPLICATIONNUMBER")
    strAFFNo = xmlGetAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApp = xmlDoc.createElement("APPLICATION")
    xmlNewRequest.appendChild xmlApp
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strApplNo
    xmlApp.appendChild xmlElem
    
    strResponse = objAppBO.GetApplicationData(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
    strOtherSysAccountNumber = xmlDoc.selectSingleNode("//APPLICATION/OTHERSYSTEMACCOUNTNUMBER").Text
   
    'MAR1122 Error if no account number is present
    If (Len(strOtherSysAccountNumber) = 0) Then
        errThrowError cstrFunctionName, oeMissingElementValue, "No OtherSystemAccountNumber for APPLICATIONNUMBER " & strApplNo
    End If
   
    ' Make call to the Interface
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "UPDATELEGALADDRESSINDICATOR", "1"
    Set xmlElem = xmlApplication.cloneNode(False)
    xmlNewRequest.appendChild xmlElem
    xmlSetAttributeValue xmlElem, "OTHERSYSTEMACCOUNTNUMBER", strOtherSysAccountNumber
        
    'Call the interface here
    Set objInterfaceBO = gobjContext.CreateInstance("omLAU.LegalAddressBO")
    strResponse = objInterfaceBO.RunLegalAddressTriggerInterface(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
        
    'MAR853  If no error returned, Update the task status as complete
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)

    strResponse = gobjTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
RunLegalAddressTriggerInterfaceExit:

    Set xmlApplication = Nothing
    Set xmlCaseTask = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElem = Nothing
    Set xmlApp = Nothing
    Set xmlDoc = Nothing
    Set objAppBO = Nothing
    Set objInterfaceBO = Nothing
    
    gobjContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub

End Sub
'SR 04/11/2005 : MAR329 - End

'MAR750  New RunPinNumberTriggerInterface
Private Sub RunPinNumberTriggerInterface(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo RunPinNumberTriggerInterfaceExit

    Const cstrFunctionName As String = "RunPinNumberTriggerInterface"
    
    Dim xmlApplication As IXMLDOMElement
    Dim xmlCaseTask As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlTable As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim xmlElem2 As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objInterfaceBO As Object
    Dim strResponse As String
    
    Dim objAppManagerBO     As Object               ' MAR1244
    Dim xmlNumberRequest    As IXMLDOMNode          ' MAR1244
    Dim xmlCustomerList     As IXMLDOMNodeList      ' MAR1244
    
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    Set xmlDoc = New FreeThreadedDOMDocument40
    
     'MAR1244 Get customer data if necessary
    If (xmlApplication.selectNodes("//CUSTOMER").length = 0) Then
    
        Set xmlNumberRequest = vxmlRequestNode.cloneNode(False)
        Set xmlNode = xmlDoc.createElement("APPLICATIONNUMBER")
        xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONNUMBER")
        xmlNumberRequest.appendChild xmlNode
        Set xmlNode = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNode.Text = xmlGetMandatoryAttributeText(xmlApplication, "APPLICATIONFACTFINDNUMBER")
        xmlNumberRequest.appendChild xmlNode
    
        Set objAppManagerBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    
        strResponse = objAppManagerBO.FindCustomersForApplication(xmlNumberRequest.xml)
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlDoc.documentElement, , True

       'Build request to call the interface
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlNewRequest.appendChild xmlApplication.cloneNode(False)

        Set xmlElem2 = xmlDoc.createElement("CUSTOMERLIST")
        xmlNewRequest.appendChild xmlElem2
        For Each xmlNode In xmlDoc.selectNodes("RESPONSE/CUSTOMER/OTHERSYSTEMCUSTOMERNUMBER")

            Set xmlTable = xmlDoc.createElement("CUSTOMER")
            xmlTable.appendChild xmlNode
        
            xmlElem2.appendChild xmlTable
        Next xmlNode

    Else
        'Use Customer data passed in.
        'Build request to call the interface
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlNewRequest.appendChild xmlApplication.cloneNode(False)
    
        Set xmlElem2 = xmlDoc.createElement("CUSTOMERLIST")
        xmlNewRequest.appendChild xmlElem2
        For Each xmlNode In xmlApplication.selectNodes("//CUSTOMER")

            Set xmlTable = xmlDoc.createElement("CUSTOMER")
           
            Set xmlElem = xmlDoc.createElement("OTHERSYSTEMCUSTOMERNUMBER")
            xmlElem.Text = xmlGetAttributeText(xmlNode, "OTHERSYSTEMCUSTOMERNUMBER")
            xmlTable.appendChild xmlElem
        
            xmlElem2.appendChild xmlTable
        Next xmlNode
    End If
    
    'Call the interface here
    Set objInterfaceBO = gobjContext.CreateInstance("omPinMailer.PinMailerBO")
    strResponse = objInterfaceBO.RunPinNumberTriggerInterface(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
      
    'If the call to Interface is successful, update the casetask
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
    xmlSetAttributeValue xmlCaseTask, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTask, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTask, "TASKSTATUS", 40
    
    xmlNewRequest.appendChild xmlCaseTask.cloneNode(True)
    
    strResponse = gobjTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
RunPinNumberTriggerInterfaceExit:

    Set xmlCaseTask = Nothing
    Set xmlApplication = Nothing
    Set xmlNewRequest = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlElem2 = Nothing
    Set xmlTable = Nothing
    Set xmlDoc = Nothing
    Set objInterfaceBO = Nothing
    
    gobjContext.SetComplete
    
    errCheckError cstrFunctionName     ' MAR750
    
    Exit Sub

End Sub
Private Sub RunGetMaxCaseTaskInstance(ByVal vxmlRequestNode As IXMLDOMNode, _
                                        ByVal vxmlResponseNode As IXMLDOMNode)
'Input : <CASETASK STAGEID="" TASKID="" CASEACTIVITYGUID="" CASESTAGESEQUENCENO="">
'Output: <TASKINSTANCE>num</TASKINSTANCE>

    Const cstrFunctionName As String = "RunGetMaxCaseTaskInstance"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim nMaxTaskInst As Integer
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    
    Set xmlCaseTaskNode = vxmlRequestNode.selectSingleNode("CASETASK")
    
    nMaxTaskInst = GetMaxCaseTaskInstance(xmlGetMandatoryAttributeText(xmlCaseTaskNode, "STAGEID"), _
                            xmlGetMandatoryAttributeText(xmlCaseTaskNode, "TASKID"), _
                            xmlGetMandatoryAttributeText(xmlCaseTaskNode, "CASEACTIVITYGUID"), _
                            xmlGetMandatoryAttributeText(xmlCaseTaskNode, "CASESTAGESEQUENCENO"))
    Set xmlResponseElem = xmlDoc.createElement("TASKINSTANCE")
    xmlResponseElem.Text = CStr(nMaxTaskInst)
    vxmlResponseNode.appendChild xmlResponseElem
    
End Sub

' MAR630  Add new function
Private Function GetMaxCaseTaskInstance(ByVal vstrStageId As String, ByVal vstrTaskId As String, _
                                        ByVal vstrCaseActivityGuid As String, ByVal vstrCaseStageSeqNo As String) As Integer

    Const cstrFunctionName As String = "GetMaxCaseTaskInstance"

    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT MAX(TaskInstance) FROM CaseTask"
    strSQL = strSQL & " WHERE CaseActivityGuid = 0x" & vstrCaseActivityGuid
    strSQL = strSQL & " AND StageId = '" & vstrStageId & "'"
    strSQL = strSQL & " AND CaseStageSequenceNo = " & vstrCaseStageSeqNo
    strSQL = strSQL & " AND TaskId = '" & vstrTaskId & "'"
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdText
    cmd.CommandText = strSQL
    Set rst = executeGetRecordSet(cmd)
    
    'Get the task instance from the result (default to 1 if not found)
    If Not rst Is Nothing Then
        If IsNull(rst.Fields(0).Value) Then
            GetMaxCaseTaskInstance = 1
        Else
            GetMaxCaseTaskInstance = rst.Fields(0).Value
        End If
    Else
        GetMaxCaseTaskInstance = 1
    End If
    
GetMaxCaseTaskInstanceExit:
    
    Set cmd = Nothing
    Set rst = Nothing
    
    errCheckError cstrFunctionName
    
End Function

'MAR1048  Add new task to create Funds Release Failed task
'MAR1110  Add creation of task note
Private Sub CreateFundsReleaseFailedTask(ByVal vxmlRequestNode As IXMLDOMNode, _
                                              ByVal vstrTaskNote As String)
        
    On Error GoTo CreateTaskErr

    Const cstrFunctionName As String = "omTMBO.CreateFundsReleaseFailedTask"
        
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim objNoTxBO  As Object
    Dim strResponse  As String
    Dim strNoteType  As String       ' MAR1110
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "CREATEFUNDSRELEASEFAILEDTASK"

    strNoteType = GetFirstComboValueId("TaskNoteType", "M")
        
    Set xmlTempRequest = xmlTempDoc.createElement("TASKNOTE")
    xmlSetAttributeValue xmlTempRequest, "NOTEENTRY", vstrTaskNote
    xmlSetAttributeValue xmlTempRequest, "NOTETYPE", strNoteType
    vxmlRequestNode.appendChild xmlTempRequest

    Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNoTxBO")
    
    strResponse = objNoTxBO.OmTmNoTxRequest(vxmlRequestNode.xml)
    
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, , True
    
    gobjContext.SetComplete
     
CreateTaskExit:
    
    Set xmlTempRequest = Nothing
    Set xmlTempDoc = Nothing
    Set objNoTxBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
        
    Exit Sub
    
CreateTaskErr:

    gobjContext.SetAbort
    
    GoTo CreateTaskExit

End Sub

'MAR1088 GHun
Private Sub RunAutoApplicationDecline(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo RunAutoApplicationDeclineErr
    Const cstrFunctionName As String = "RunAutoApplicationDecline"

    Dim xmlCaseTask As IXMLDOMNode
    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseStage As IXMLDOMNode
    Dim xmlTempResponse As FreeThreadedDOMDocument40
    Dim objNoTxBO As omTmNoTxBO
    Dim objBO As Object
    
    Dim strExceptionReason As String
    Dim strCompleteValue As String
    Dim strResponse As String
    Dim intDeclinedStage As Integer
    
    Set xmlCaseTaskIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    Set xmlTempResponse = New FreeThreadedDOMDocument40
    xmlTempResponse.validateOnParse = False
    xmlTempResponse.setProperty "NewParser", True
        
    strExceptionReason = GetFirstComboValueId("ExceptionReason", "AD")
    intDeclinedStage = GetMandatoryGlobalParamAmount("DeclinedStageValue")
    
    ' Move to Stage
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "MoveToStage"
    Set xmlCaseStage = xmlTempResponse.createElement("CASESTAGE")
    xmlTempRequest.appendChild xmlCaseStage
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "CASEID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskIn, xmlCaseStage, "ACTIVITYINSTANCE"
    xmlSetAttributeValue xmlCaseStage, "STAGEID", CStr(intDeclinedStage)
    xmlSetAttributeValue xmlCaseStage, "EXCEPTIONREASON", strExceptionReason
    xmlTempRequest.appendChild xmlApplication.cloneNode(True)
    Set objNoTxBO = gobjContext.CreateInstance(App.Title & ".omTmNoTxBo")
    strResponse = objNoTxBO.OmTmNoTxRequest(xmlTempRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
        
RunAutoApplicationDeclineExit:

    Set xmlCaseTask = Nothing
    Set xmlCaseTaskIn = Nothing
    Set xmlApplication = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseStage = Nothing
    Set xmlTempResponse = Nothing
    Set objNoTxBO = Nothing
    Set objBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RunAutoApplicationDeclineErr:

    gobjContext.SetAbort
    GoTo RunAutoApplicationDeclineExit

End Sub
'MAR1088 End

'MAR1717 GHun Changed to return an xml node
'MAR1172
Private Function InitialCreditCheckDone(ByVal xmlRequestNode As IXMLDOMNode) As IXMLDOMNode

    On Error GoTo InitialCreditCheckDoneVbErr

    Const strFunctionName As String = "InitialCreditCheckDone"
    
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempNode As IXMLDOMNode
    
    Dim strData As String
    Dim strAppNumber As String
    Dim strAppFFNumber As String
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    
    strAppNumber = xmlGetMandatoryAttributeText(xmlRequestNode, "APPLICATIONNUMBER")
    Set adoParameter = adoCommand.CreateParameter("@pApplicationNumber", adBSTR, adParamInput, Len(strAppNumber), strAppNumber)
    adoCommand.Parameters.Append adoParameter
    
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlRequestNode, "APPLICATIONFACTFINDNUMBER")
    Set adoParameter = adoCommand.CreateParameter("@pApplicationFactFindNumber", adInteger, adParamInput, , CInt(strAppFFNumber))
    adoCommand.Parameters.Append adoParameter
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.Open
    
    With adoCommand
        .CommandText = "USP_GETCREDITCHECKSTATUS"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Properties("Output Stream") = adoStream
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    If Len(strData) > 0 Then    'MAR1717 GHun
        Set xmlOut = xmlLoad(strData, strFunctionName)
    
        'MAR1717 GHun
        If xmlOut Is Nothing Then
            Set InitialCreditCheckDone = Nothing
        Else
            Set InitialCreditCheckDone = xmlOut.firstChild
        End If
        'MAR717 End
    End If
 
InitialCreditCheckDoneExit:
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Function
    
InitialCreditCheckDoneVbErr:
        
    gobjContext.SetComplete
    
    GoTo InitialCreditCheckDoneExit
    

End Function
'MAR1172

' PSC 03/02/2006 MAR1156 - Start
Private Sub SendPack(ByVal vxmlRequest As IXMLDOMNode, ByVal strPackNumber As String)
On Error GoTo SendPackExit:

    Const cstrFunctionName As String = "SendPack"

    Dim xmlPackRequestDoc As FreeThreadedDOMDocument40
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlPackRequestNode As IXMLDOMNode
    Dim objPack As Object
    
    Dim strResponse As String
        
    Set xmlPackRequestDoc = New FreeThreadedDOMDocument40
    xmlPackRequestDoc.async = False
    xmlPackRequestDoc.setProperty "NewParser", True
        
    Set xmlApplicationNode = vxmlRequest.selectSingleNode("APPLICATION")
    
    Set xmlPackRequestNode = xmlGetRequestNode(vxmlRequest)
    xmlPackRequestDoc.appendChild xmlPackRequestNode
    xmlSetAttributeValue xmlPackRequestNode, "OPERATION", "SENDPACK"
    xmlCopyMandatoryAttribute xmlApplicationNode, xmlPackRequestNode, "APPLICATIONNUMBER"
    xmlCopyMandatoryAttribute xmlApplicationNode, xmlPackRequestNode, "APPLICATIONFACTFINDNUMBER"
    xmlSetAttributeValue xmlPackRequestNode, "PACKCONTROLNUMBER", strPackNumber
   
    On Error Resume Next
    Set objPack = gobjContext.CreateInstance("omPack.PackManagerBO")
    On Error GoTo SendPackExit

    If Not objPack Is Nothing Then
        strResponse = objPack.omRequest(xmlPackRequestNode.xml)
        errCheckXMLResponse strResponse, True
    End If
        
SendPackExit:

    Set xmlPackRequestDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlPackRequestNode = Nothing
    Set objPack = Nothing
    
    errCheckError cstrFunctionName
    
    Exit Sub
End Sub
' PSC 03/02/2006 MAR1156 - End

'MAR1300 GHun
Private Sub SetChangeOfProperty(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    Const cstrFunctionName  As String = "SetChangeOfProperty"
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlApp              As IXMLDOMElement
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlNewProperty      As IXMLDOMElement
    Dim xmlElem             As IXMLDOMElement
    Dim objAppBO            As Object
    Dim strResponse         As String
    Dim lngErrNo            As Long     'MAR1607 GHun
    
On Error GoTo SetChangeOfPropertyExit

    Set xmlDoc = xmlCreateDOMObject
    Set xmlApp = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlNewProperty = xmlDoc.createElement("NEWPROPERTY")
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONNUMBER")
    xmlNewProperty.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElem.Text = xmlGetMandatoryAttributeText(xmlApp, "APPLICATIONFACTFINDNUMBER")
    xmlNewProperty.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CHANGEOFPROPERTY")
    xmlElem.Text = "1"
    xmlNewProperty.appendChild xmlElem
    
    xmlRequest.appendChild xmlNewProperty

    Set objAppBO = gobjContext.CreateInstance("omApp.NewPropertyBO")
    strResponse = objAppBO.UpdateNewProperty(xmlRequest.xml)
    
    'MAR1607 GHun
    lngErrNo = errCheckXMLResponse(strResponse, False)
    If lngErrNo <> 0 Then
        If errGetOmigaErrorNumber(lngErrNo) = oeNoRowsAffected Then
            errThrowError cstrFunctionName, 8549
        Else
            errCheckXMLResponse strResponse, True, vxmlResponseNode
        End If
    End If
    'MAR1607 End
        
    'Update Case Task
    Set xmlRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlRequest, "OPERATION", "UPDATECASETASK"
    Set xmlElem = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK").cloneNode(True)
    
    xmlElem.setAttribute "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlElem.setAttribute "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlElem.setAttribute "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlElem.setAttribute "TASKSTATUS", cstrTaskComplete
    
    xmlRequest.appendChild xmlElem
    
    strResponse = gobjTmBO.TmRequest(xmlRequest.xml)
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
SetChangeOfPropertyExit:
    'MAR1607 GHun
    If Err.Number <> 0 Then
        gobjContext.SetAbort
    End If
    'MAR1607 End
    Set xmlDoc = Nothing
    Set xmlApp = Nothing
    Set xmlRequest = Nothing
    Set xmlNewProperty = Nothing
    Set xmlElem = Nothing
    Set objAppBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
End Sub
'MAR1300 End

' PSC 05/05/2006 MAR1593 - Start
Private Sub RunEmailSMSInterface(ByVal vxmlRequest As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo RunEmailSMSInterfaceExit:

Const strFunctionName As String = "RunEmailSMSInterface"
    
    Dim xmlTempDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlTempRequest As IXMLDOMNode
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlFullfillmentRequest As IXMLDOMNode
    
    Set xmlFullfillmentRequest = vxmlRequest.cloneNode(True)
    Set xmlCaseTaskNode = xmlGetMandatoryNode(xmlFullfillmentRequest, "CASETASK")
    xmlSetAttributeValue xmlCaseTaskNode, "BYPASSTASKSTATUSUPDATE", "1"
    
    SMSFulfilmentExtract xmlFullfillmentRequest, vxmlResponseNode
    EMailFulfilmentExtract xmlFullfillmentRequest, vxmlResponseNode

    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
        
    Set xmlTempRequest = vxmlRequest.cloneNode(False)
    xmlTempDoc.appendChild xmlTempRequest
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UpdateCaseTask"
    Set xmlCaseTaskNode = xmlCaseTaskNode.cloneNode(True)
    xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", "40"
    xmlTempRequest.appendChild xmlCaseTaskNode
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.async = False
    xmlOut.setProperty "NewParser", True
    
    ' run Tm UpdateCaseTask request
    RunTmRequest xmlTempDoc, xmlOut

    Debug.Print "UpdateCaseTask response:"
    Debug.Print xmlOut.xml
    ' parse response
    errCheckXMLResponseNode xmlOut.selectSingleNode("RESPONSE"), , True

    
RunEmailSMSInterfaceExit:

    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlFullfillmentRequest = Nothing
    Set xmlOut = Nothing

    errCheckError strFunctionName
End Sub
' PSC 05/05/2006 MAR1593 - End

'MAR1715 GHun
Private Sub GetApplicationFirstTitleFromId(ByVal vxmlRequestNode As IXMLDOMNode, ByRef vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName As String = "GetApplicationFirstTitleFromId"
    
    Dim adoConn As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoStream As ADODB.Stream
    Dim adoParameter As ADODB.Parameter
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlAppNode As IXMLDOMNode
    Dim xmlAppFTNode As IXMLDOMNode
    
    Dim strData As String
    Dim strAppNumber As String
    Dim strAppFactFindNo As String
    Dim strId As String
    
On Error GoTo GetApplicationFirstTitleFromIdVbErr
    
    Set adoConn = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoStream = New ADODB.Stream
    
    Set xmlAppNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    strAppNumber = xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONNUMBER")
    strAppFactFindNo = xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONFACTFINDNUMBER")
    strId = xmlGetMandatoryAttributeText(xmlAppNode, "APPLICATIONFIRSTITLEID")
    
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
    
    ' Use stream object to retrieve XML from the database
    adoStream.Open
    
    With adoCommand
        .CommandText = "USP_GETAPPLICATIONFIRSTTITLE"
        .CommandType = adCmdStoredProc
        .ActiveConnection = adoConn
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, Len(strAppNumber), strAppNumber)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , strAppFactFindNo)
        .Parameters.Append .CreateParameter("@p_ApplicationFirstTitleId", adInteger, adParamInput, , strId)
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
    End With
    
    strData = adoStream.ReadText()
    
    adoConn.Close
    adoStream.Close
    
    Set xmlOut = xmlLoad(strData, cstrFunctionName)
        
    Set xmlAppFTNode = xmlGetNode(xmlOut.documentElement, "APPLICATIONFIRSTTITLE")
    
    If (xmlAppFTNode Is Nothing) Then
        errThrowError cstrFunctionName, oeRecordNotFound
    Else
        xmlSetAttributeValue xmlOut.documentElement, "TYPE", "SUCCESS"
    End If
    
    Set vxmlResponseNode = xmlOut.documentElement
    
GetApplicationFirstTitleFromIdExit:
    
    Set adoConn = Nothing
    Set adoCommand = Nothing
    Set adoStream = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me), vxmlResponseNode
    
    Exit Sub
    
GetApplicationFirstTitleFromIdVbErr:
    
    gobjContext.SetComplete
    Set vxmlResponseNode = CreateErrorResponseNode
    GoTo GetApplicationFirstTitleFromIdExit
    
End Sub
'MAR1715 End

#If EPSOM Then
'EP17 SAB - Start
Private Sub RunRealHunterExperianInterface(ByVal vxmlRequestNode As IXMLDOMNode, _
                                       ByVal vxmlResponseNode As IXMLDOMNode)
    
    Const cstrFunctionName  As String = "RunRealHunterExperianInterface"
    
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlTemp             As IXMLDOMElement
    Dim objMsgTmBO          As Object
    
    Dim objExperianHI As Object
    Dim strApplNumber As String
    Dim strAFFNumber As String
    Dim strTaskStatus As String
    Dim strResponse As String
    Dim lngErrNo As Long

    'MAR339
    Dim xmlHUNONode         As IXMLDOMNode
    Dim strErrorCode        As String
    Dim strErrorMessage     As String
    Dim strMessage          As String
    Dim strReturnCode       As String
    
On Error GoTo RunRealHunterExperianInterfaceError

    'Initialise
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'MAR339
    'Initialise the task status to Incomplete.
    'Status will be set to Complete when a successful Experian Hunter call is made.
    strTaskStatus = GetFirstComboValueId("TaskStatus", "I") 'Task Incomplete
    
    'Validate
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    strApplNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
    
    Set objExperianHI = gobjContext.CreateInstance("omExperianHI.ExperianHIBO")
    strResponse = objExperianHI.RunExperianHI(xmlTempRequest.xml)
    
    'Check for an ErrorCode in the response and raise an error if found.
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
    'A call to Experian will result in a HUNO node or an ERR1 node.
    'Check for the error block and raise an error if found.
    
    xmlTempDoc.loadXML strResponse

    ' NOTE: Any Experian errors will be handled by omExperianHI
    
    Set xmlHUNONode = xmlTempDoc.selectSingleNode(".//HUNO")
    
    'Check the return code in the HUNO block and raise an error if necessary.
    If Not xmlHUNONode Is Nothing Then
        
        strReturnCode = xmlHUNONode.selectSingleNode(".//RETURNCODE").Text
        
        If (strReturnCode = "9999") Then
            'Raise error and do not complete the task (status will stay at Incomplete)
            errThrowError cstrFunctionName, 7001
        
        ElseIf (strReturnCode = "0000") Then
        
            'Success - set the task to Complete
            strTaskStatus = GetFirstComboValueId("TaskStatus", "CP")  'Task complete
        
        End If
   
    End If
    
    'Now Update the CaseTask
    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlTempRequest, "OPERATION", "UPDATECASETASK"
    
    Set xmlTemp = xmlCaseTaskNode.cloneNode(True)
    xmlSetAttributeValue xmlTemp, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlTemp, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlTemp, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlTemp, "TASKSTATUS", strTaskStatus
    xmlTempRequest.appendChild xmlTemp
    
    Set objMsgTmBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlTempRequest.xml)
    
    errCheckXMLResponse strResponse, True, vxmlResponseNode
    
RunRealHunterExperianInterfaceExit:
    Set xmlTempDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTemp = Nothing
    Set xmlHUNONode = Nothing
    
    Set objExperianHI = Nothing
    Set objMsgTmBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
RunRealHunterExperianInterfaceError:

    gobjContext.SetAbort
    
    GoTo RunRealHunterExperianInterfaceExit

End Sub
'EP17 SAB - End

'EP2 IK
Private Sub RunEpsomCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMElement, _
                                Optional ByVal blnIsRescore As Boolean = False, _
                                Optional ByVal blnIsReprocess As Boolean = False)

    Const cstrFunctionName  As String = "RunEpsomCreditCheck"
        
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim strResponse         As String
    Dim xmlHeaderNode       As IXMLDOMNode
    Dim xmlTargetReqNode    As IXMLDOMNode
    Dim xmlCCN1Node         As IXMLDOMNode
    Dim bAddressTargeting   As Boolean
    Dim xmlNewRequest       As IXMLDOMNode
    Dim objMsgTmBO          As Object
    Dim objExperianCC       As Object
    Dim blnLIM_Rescore      As Boolean
    Dim strStageId          As String    'MAR630

On Error GoTo RunEpsomCreditCheckExit

    'Initialise ======================================================================
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    
    'Validate ========================================================================
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    ' PSC 04/11/2005 MAR329
    Set xmlCaseTaskNode = xmlGetNode(vxmlRequestNode, "CASETASK")
    
    ' MAR245 - Check if request includes address targeting data.
    Set xmlTargetReqNode = xmlGetNode(vxmlRequestNode, "TARGETINGDATA")
    
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    
    ' PSC 04/11/2005 MAR329 - Start
    If Not xmlCaseTaskNode Is Nothing Then
        xmlTempRequest.appendChild xmlCaseTaskNode.cloneNode(True)
    End If
    ' PSC 04/11/2005 MAR329 - End
    
    ' ik_MAR452_20051111
    ' xmlTempRequest.appendChild xmlApplicationNode
    xmlTempRequest.appendChild xmlApplicationNode.cloneNode(True)
    ' ik_MAR452_20051111_ends
    
    ' Add our Address Target Request if/when appropriate
    If Not xmlTargetReqNode Is Nothing Then
        xmlTempRequest.appendChild xmlTargetReqNode
    End If
    
    If blnIsReprocess Then
        xmlTempRequest.setAttribute "REPROCESS", "1"
    End If
    If blnIsRescore Then
        xmlTempRequest.setAttribute "RESCORE", "1"
    End If
    
    'MAR841 Set the operation to RunXMLCreditCheck - flags are used for Reprocess and Rescore
    xmlTempRequest.setAttribute "OPERATION", "RunXMLCreditCheck"
    
    'DRC - Check for LIM Rescore task
    'MAR630
    If xmlCaseTaskNode Is Nothing Then
        blnLIM_Rescore = False
    Else
        If UCase(xmlGetAttributeText(xmlCaseTaskNode, "TASKID")) = _
            UCase(GetGlobalParamString("TMLIMRescoreID")) Then
            blnLIM_Rescore = True
        Else
            blnLIM_Rescore = False
        End If
    End If
    
    'Do Credit Score =============================================
    Set objExperianCC = gobjContext.CreateInstance("omExperianCC.ExperianCCBO")
    strResponse = objExperianCC.Execute(xmlTempRequest.xml)
    
    xmlTempDoc.loadXML strResponse
    errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
   
    AttachResponseData vxmlResponseNode, xmlTempDoc.documentElement
     
    bAddressTargeting = xmlGetNodeAsBoolean(vxmlResponseNode, ".//ADDRESSTARGETING")
    'If Address Targeting is required don't save at this time
    If Not (bAddressTargeting) Then
    
        'If we've got this far, need to Update the Case Task status
        'Build the Request and update the case task
        ' PSC 04/11/2005 MAR329 - Start
        If Not xmlCaseTaskNode Is Nothing Then
            Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    
            xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(xmlTempRequest, "USERID")
            xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(xmlTempRequest, "UNITID")
            xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
            xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
            xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)
    
            Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
            strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
            errCheckXMLResponse strResponse, True
        End If
        ' PSC 04/11/2005 MAR329 - End

        
        'Call Risk Assessment
        'DRC MAR33 - need to add a context so that RA tasks will be added
        'MAR630
        If Not xmlCaseTaskNode Is Nothing Then                               'MAR630
            xmlSetAttributeValue vxmlRequestNode, "TASKMANAGER", "T"
            
            strStageId = xmlGetAttributeText(xmlCaseTaskNode, "STAGEID")     'MAR630
            xmlSetAttributeValue xmlApplicationNode, "STAGEID", strStageId   'MAR630
        End If                                                               'MAR630
        
        'DRC MAR599 - don't run Case Assessment for  LIM Rescore
        'ik_EP2_17/03/2006
'        If Not blnLIM_Rescore Then
'            RunRiskAssessment vxmlRequestNode, vxmlResponseNode
'        End If
        'ik_EP2_17/03/2006_ends
        'DRC MAR599
    Else
        'MAR1084 GHun Throw an error if address targeting is required, but not supported
        If xmlGetAttributeAsBoolean(vxmlRequestNode, "NOADDRESSTARGETING") Then
            errThrowError cstrFunctionName, oeTmAddressTargetingNotSupported
        End If
        'MAR1084 End
        
        ' MAR245
        ' Flag if a recore or reprocess was requested.
        If blnIsReprocess Then
           vxmlResponseNode.setAttribute "REPROCESS", "1"
        End If
        If blnIsRescore Then
            vxmlResponseNode.setAttribute "RESCORE", "1"
        End If
        
        'Copy the CaseTask node from the request to the response, when address targeting
        If Not xmlCaseTaskNode Is Nothing Then  'MAR1084 GHun CaseTask node does not always exist
            vxmlResponseNode.appendChild xmlCaseTaskNode.cloneNode(True)
        End If
    End If
    
RunEpsomCreditCheckExit:
   
    Set objExperianCC = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlNewRequest = Nothing
    Set objMsgTmBO = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub RunEpsomCaseAssess(ByVal vxmlRequestNode As IXMLDOMNode, Optional ByVal vxmlResponseNode = Nothing)

    On Error GoTo RunEpsomCaseAssessExit
    
    Const cstrFunctionName As String = "RunEpsomCaseAssess"
    
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlCaseTaskElem As IXMLDOMElement
    
    'SR 04/11/2005 : MAR269
    Dim xmlCaseActivity As IXMLDOMElement
    Dim xmlCurrentCaseStage As IXMLDOMElement
    Dim xmlRiskResponseDoc As FreeThreadedDOMDocument40 'MAR1173 GHun
    Dim xmlTempResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlTempElem2 As IXMLDOMElement
    Dim xmlInterfaceElem As IXMLDOMElement     ' MAR630
    Dim intFailLevel As Integer
    Dim intUWCount As Integer
    Dim intProcCount As Integer
    Dim intFraudCount As Integer
    'SR 04/11/2005 : MAR269 - End

    Dim objBO As Object
    
    Dim strApplicationNumber As String
    Dim strAFFNumber As String
    Dim strStageId As String 'SR 04/11/2005 : MAR269
    Dim strResponse As String
    Dim strInterfaceType As String          ' MAR630
    Dim strTaskId As String                 ' MAR630
    Dim strCaseAssessment As String         ' MAR630
    
    ' PSC 21/03/2007 EP2_1619 - Start
    Dim xmlScores As IXMLDOMNodeList
    Dim xmlScore As IXMLDOMNode
    Dim xmlValidationType As IXMLDOMNode
    
    Dim xmlCrudRequest As IXMLDOMNode
    Dim xmlRoot As IXMLDOMNode
    Dim xmlComboValues As FreeThreadedDOMDocument40
    Dim xmlMsgSubType As IXMLDOMNode
    Dim objCrudBO As Object

    Dim blnUseCaseAssessmentReferral As Boolean
    Dim strMsgSubType As String
    Dim intMsgSubTypePos As Integer
    ' PSC 21/03/2007 EP2_1619 - End

    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
    
    'SR 04/11/2005 : MAR269 - check whether StageId was passed in (as an attrib to Application node)
    '                if no, call GetCurrentStage
    Set xmlTempResponseDoc = New FreeThreadedDOMDocument40
    
    ' PSC 21/03/2007 EP2_1619 - Start
    xmlTempResponseDoc.validateOnParse = False
    xmlTempResponseDoc.setProperty "NewParser", True
    xmlTempResponseDoc.async = False
    ' PSC 21/03/2007 EP2_1619 - End

    'MAR630
    strTaskId = ""
    If Not vxmlRequestNode.selectSingleNode("CASETASK[@TASKID]") Is Nothing Then
        strTaskId = vxmlRequestNode.selectSingleNode("CASETASK/@TASKID").Text
    End If
      
    'Create Request and run RiskAssessment
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)

    If xmlGetAttributeText(xmlApplicationElem, "STAGEID") = "" Then
        xmlSetAttributeValue xmlApplicationElem, "STAGEID", strStageId
    End If
    'SR 04/11/2005 : MAR269 - End
    
    ' PSC 20/11/01 SYS2867/SYS2988
    xmlNewRequest.appendChild xmlMakeNodeElementBased(xmlApplicationElem, False, "RISKASSESSMENT")
    
    Set objBO = gobjContext.CreateInstance("omRA.RiskAssessmentBO")
    strResponse = objBO.RunRiskAssessment(xmlNewRequest.xml)
    
    ' ik_MAR452_20051111
    ' xmlLoad strResponse, xmlTempResponseDoc
    Set xmlRiskResponseDoc = xmlLoad(strResponse, cstrFunctionName) 'MAR1173 GHun
    ' ik_MAR452_20051111_ends
    
    Dim lngErrNo As Long
    lngErrNo = errCheckXMLResponseNode(xmlRiskResponseDoc.documentElement, , False) 'MAR1173 GHun
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, 4813
    Else   'SR 04/11/2005 : MAR269
        'MAR1173 GHun
        intFailLevel = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "FAILLEVEL")
        intUWCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "UNDERWRITERCOUNT")
        intProcCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "PROCESSORCOUNT")
        intFraudCount = xmlGetAttributeAsInteger(xmlRiskResponseDoc.documentElement, "PROCESSORCOUNT")
        'MAR1173 End
        'MAR1142 DRC Allow web apps to RA create tasks
        ' If xmlGetAttributeText(vxmlRequestNode, "TASKMANAGER") <> "" And intFailLevel > 1 Then
        'MAR1142 DRC
        If intFailLevel > 1 Then
        
            ' PSC 21/03/2007 EP2_1619
            blnUseCaseAssessmentReferral = GetGlobalParamBoolean("UseCaseAssessmentReferral")
            'Call HandleInterfaceResponse here
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "HANDLEINTERFACERESPONSE"
            xmlNewRequest.appendChild xmlApplicationElem.cloneNode(True)
            'MAR 1142 DRC - Case Activity can be created in Handle Interface
            If Not xmlCaseActivity Is Nothing Then
                xmlNewRequest.appendChild xmlCaseActivity.cloneNode(True)
            End If
            'MAR 1142 DRC
            'MAR630  Create the MESSAGESUBTYPELIST node under INTERFACE
            Set xmlInterfaceElem = xmlTempResponseDoc.createElement("INTERFACE")
            
            'MAR630 Get the value ID for interface type Case Assessement "CA"
            strInterfaceType = GetFirstComboValueId("InterfaceType", "CA")
            xmlSetAttributeValue xmlInterfaceElem, "INTERFACETYPE", strInterfaceType
             
            xmlSetAttributeValue xmlInterfaceElem, "MESSAGETYPE", intFailLevel
            
            xmlSetAttributeValue xmlInterfaceElem, "CREATETASKFLAG", "1"
            
            'MAR 1142 DRC - No messagesubtypes for decline tasks
            
            ' PSC 21/03/2007 EP2_1619 - Start
            If blnUseCaseAssessmentReferral = False Then
                If intFailLevel < 3 Then
                    If intUWCount > 0 Or intProcCount > 0 Or intFraudCount > 0 Then
                        Set xmlTempElem2 = xmlTempResponseDoc.createElement("MESSAGESUBTYPELIST")
                        xmlSetAttributeValue xmlTempElem2, "CREATETASKFLAG", 1
                        xmlInterfaceElem.appendChild xmlTempElem2
                    End If
                             
                    If intUWCount > 0 Then
                        Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                        xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "10"
                        xmlTempElem2.appendChild xmlTempElem
                    End If
                    
                    If intProcCount > 0 Then
                        Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                        xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "20"
                        xmlTempElem2.appendChild xmlTempElem
                    End If
                    
                    If intFraudCount > 0 Then
                        Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                        xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", "30"
                        xmlTempElem2.appendChild xmlTempElem
                    End If
                End If
            Else
                ' Get the combo group in order to determine the message sub types
                Set objCrudBO = CreateObject(gstrCRUD_COMPONENT & ".omCrudBO")
                Set xmlCrudRequest = xmlGetRequestNode(vxmlRequestNode)
                xmlSetAttributeValue xmlCrudRequest, "CRUD_OP", "READ"
                xmlSetAttributeValue xmlCrudRequest, "ENTITY", "COMBOGROUP"
                Set xmlRoot = xmlTempResponseDoc.createElement("COMBOGROUP")
                xmlSetAttributeValue xmlRoot, "GROUPNAME", "CaseAssessmentReferralType"
                xmlCrudRequest.appendChild xmlRoot
                    
                strResponse = objCrudBO.omRequest(xmlCrudRequest.xml)
                Set xmlComboValues = xmlLoad(strResponse, cstrFunctionName)
                errCheckXMLResponseNode xmlComboValues.documentElement, vxmlResponseNode, True
                
                strStageId = xmlGetNodeText(vxmlRequestNode, "CASETASK/@STAGEID")
                    
                ' Get unique failed scores
                Set xmlScores = xmlRiskResponseDoc.selectNodes("RESPONSE/RULE[@SCORE != '0' and not(@SCORE = preceding-sibling::RULE/@SCORE)]/@SCORE")
                        
                If xmlScores.length > 0 Then
                    Set xmlTempElem2 = xmlTempResponseDoc.createElement("MESSAGESUBTYPELIST")
                    xmlSetAttributeValue xmlTempElem2, "CREATETASKFLAG", 1
                    xmlInterfaceElem.appendChild xmlTempElem2
                End If
                    
                ' Get the message sub types
                For Each xmlScore In xmlScores
                    ' Get the corresponding validation type
                    Set xmlValidationType = xmlComboValues.selectSingleNode("RESPONSE/COMBOGROUP/COMBOVALUE[@VALUEID='" & strStageId & "']/COMBOVALIDATION[starts-with(@VALIDATIONTYPE, '$" & xmlScore.Text & "')]/@VALIDATIONTYPE")
                    
                    ' If validation type is found get the message sub type
                    If Not xmlValidationType Is Nothing Then
                        intMsgSubTypePos = InStrRev(xmlValidationType.Text, "%")
                        strMsgSubType = Mid$(xmlValidationType.Text, intMsgSubTypePos + 1)
                        
                        ' Only create one if we haven't got one for this type already
                        Set xmlMsgSubType = xmlTempElem2.selectSingleNode("MESSAGESUBTYPE[@MESSAGESUBTYPE='" & strMsgSubType & "']")
                    
                        If xmlMsgSubType Is Nothing Then
                            Set xmlTempElem = xmlTempResponseDoc.createElement("MESSAGESUBTYPE")
                            xmlSetAttributeValue xmlTempElem, "MESSAGESUBTYPE", strMsgSubType
                            xmlTempElem2.appendChild xmlTempElem
                        End If
                    End If
                Next
            End If 'MAR 1142 DRC
            ' PSC 21/03/2007 EP2_1619 - End
            
            ' IK 11/04/2007 EP2_2264
            ' HandleInterfaceResponse only when tasks identified
            
            If xmlTempElem2.hasChildNodes Then
            
               xmlNewRequest.appendChild xmlInterfaceElem
    
               strResponse = OmTmRequest(xmlNewRequest.xml)
               Set xmlTempResponseDoc = xmlLoad(strResponse, cstrFunctionName)
               errCheckXMLResponseNode xmlTempResponseDoc.documentElement, , True 'MAR1143 GHun Raise the error
               
            End If
            ' IK 11/04/2007 EP2_2264 ends
            
        End If

        ' Now update Case Task if input
        If Not vxmlRequestNode.selectSingleNode("CASETASK[@TASKID]") Is Nothing Then
            Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
            xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
            
            Set xmlCaseTaskElem = vxmlRequestNode.selectSingleNode("CASETASK").cloneNode(True)
            xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
            xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
            xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
            xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", 40
            
            xmlNewRequest.appendChild xmlCaseTaskElem
            
            Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
            strResponse = objBO.TmRequest(xmlNewRequest.xml)
            
            errCheckXMLResponse strResponse, True
        End If
        
        ' PSC 04/11/2005 MAR329- Start
        ' Send the response back if required
        If Not vxmlResponseNode Is Nothing Then
        
            Dim xmlRAChild As IXMLDOMNode
            
            For Each xmlRAChild In xmlRiskResponseDoc.documentElement.childNodes    'MAR1173 GHun
                ' ik_MAR452_20051111
                vxmlResponseNode.appendChild xmlRAChild.cloneNode(True)
            Next
            
            ' ik_MAR452_20051111
            Set xmlRAChild = Nothing
        End If
        ' PSC 04/11/2005 MAR329 - End

    End If 'SR 04/11/2005 : MAR269 - End

RunEpsomCaseAssessExit:
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    Set objBO = Nothing
    'SR 04/11/2005 : MAR269
    Set xmlCaseActivity = Nothing
    Set xmlCurrentCaseStage = Nothing
    Set xmlTempResponseDoc = Nothing
    Set xmlRiskResponseDoc = Nothing    'MAR1173 GHun
    Set xmlTempElem = Nothing
    Set xmlTempElem2 = Nothing
    'SR 04/11/2005 : MAR269 - End
    
    ' PSC 21/03/2007 EP2_1619 - Start
    Set xmlScores = Nothing
    Set xmlScore = Nothing
    Set xmlValidationType = Nothing
    Set xmlCrudRequest = Nothing
    Set xmlRoot = Nothing
    Set xmlComboValues = Nothing
    Set objCrudBO = Nothing
    Set xmlMsgSubType = Nothing
    ' PSC 21/03/2007 EP2_1619 - End

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
Private Sub RunEpsomKYC( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    RunEpsomKYC -- EP16
' author:       LDM
' date:         04/04/2006
' description:  Performes a money laundering check against the customer and saves the results
'               in Omiga. If the process is successful, the Run KYC task will be completed.
'
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       vxmlResponseNode
'               XML RESPONSE node
'------------------------------------------------------------------------------------------

On Error GoTo RunEpsomKYCError
    
    Const cstrFunctionName         As String = "RunEpsomKYC"
        
    Dim xmlDoc                     As FreeThreadedDOMDocument40
    Dim xmlApplicationNode         As IXMLDOMNode
    Dim xmlCaseTaskNode            As IXMLDOMNode
    Dim xmlElem                    As IXMLDOMElement
    Dim xmlNewRequest              As IXMLDOMNode
    Dim objCustomerBO              As Object
    Dim objMsgTmBO                 As Object
    Dim strAppNo                   As String
    Dim strCustomerNo              As String
    Dim strCustomerNumberResponse  As String
    Dim strResponse                As String
    Dim strKYCResponse             As String

    'Initialise
    gobjContext.SetComplete
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
    'Get Customer Details: Build Request and Get Customer data
    strAppNo = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strCustomerNo = xmlGetAttributeText(xmlCaseTaskNode, "CUSTOMERIDENTIFIER")
    
    Set objCustomerBO = gobjContext.CreateInstance("omCust.CustomerBO")
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlElem = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElem.Text = strAppNo
    xmlNewRequest.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CUSTOMERNUMBER")
    xmlElem.Text = strCustomerNo
    xmlNewRequest.appendChild xmlElem
    
    'get the customer info
    strCustomerNumberResponse = objCustomerBO.GetCustomerNumbers(xmlNewRequest.xml)
    errCheckXMLResponse strCustomerNumberResponse, True
    
    xmlDoc.loadXML (strCustomerNumberResponse)
    Set xmlNewRequest = xmlGetMandatoryNode(xmlDoc, "//RESPONSE")
    Set xmlElem = xmlDoc.createElement("USERID")
    xmlElem.Text = xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlNewRequest.appendChild xmlElem
      
    Dim objKnowYourCustomer As Object   'MAR231 GHun changed to late binding as per omTM standard
    Set objKnowYourCustomer = gobjContext.CreateInstance("omKnowYourCustomer.KnowYourCustomerBO")
    
    'Call Experian KYC Interface (C# COM omKnowYourCustomer)
    strKYCResponse = objKnowYourCustomer.RUNKNOWYOURCUSTOMER(xmlDoc.documentElement.xml)
    errCheckXMLResponse strKYCResponse, True
    
                                  
    '==========================================================
    'Update the task as completed.
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"

    xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
    xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")

    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    
RunEpsomKYCExit:
   
    Set xmlDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlElem = Nothing
    Set xmlNewRequest = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    Exit Sub
    
RunEpsomKYCError:

    gobjContext.SetAbort
    Set xmlDoc = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlElem = Nothing
    Set xmlNewRequest = Nothing
    Set objCustomerBO = Nothing
    Set objMsgTmBO = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

'EP2 IK ends

'LDM 23/03/2006  EP6
'Upgrade an enquiry footprint creditcheck to a full application footprint
Private Sub RunXMLExpQuotationConversionCC(ByVal vxmlRequestNode As IXMLDOMNode, _
                                ByVal vxmlResponseNode As IXMLDOMElement)

    Const cstrFunctionName  As String = "RunXMLExpQuotationConversionCC"
        
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim xmlTempRequest      As IXMLDOMElement
    Dim xmlNewRequest       As IXMLDOMElement
    Dim xmlApplicationNode  As IXMLDOMNode
    Dim xmlCaseTaskNode     As IXMLDOMNode
    Dim objMsgTmBO          As Object
    Dim objExperianCC       As Object
    
    Dim strResponse         As String
    Dim strAppNo            As String
    Dim strTaskId           As String
    Dim DoUpgrade           As Boolean
    
On Error GoTo RunXMLExpQuotationConversionCCError

    'Initialise
    gobjContext.SetComplete
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    xmlTempDoc.async = False
    DoUpgrade = False
    
    'Validate
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
    
    'Call the stored procedure to see if it is valid to upgrade the experian footprint
    DoUpgrade = CheckUpdateConditions(xmlApplicationNode)
    
    If DoUpgrade = True Then
        
        'Do the Upgrade Credit Score Experian call
        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
        Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
        xmlNewRequest.appendChild xmlApplicationNode.cloneNode(False)
        xmlNewRequest.setAttribute "OPERATION", "RUNXMLUPGRADECREDITCHECK"
        
        Set objExperianCC = gobjContext.CreateInstance("omExperianCC.ExperianCCBO")
        strResponse = objExperianCC.Execute(xmlNewRequest.xml)
         
        xmlTempDoc.loadXML strResponse
        AttachResponseData vxmlResponseNode, xmlTempDoc.documentElement
        errCheckXMLResponseNode xmlTempDoc.documentElement, vxmlResponseNode, True
        
    End If
    
    '==================================================
    'Update the current task.
    Set xmlNewRequest = xmlGetRequestNode(xmlTempRequest)
    xmlNewRequest.setAttribute "OPERATION", "UPDATECASETASK"
    xmlSetAttributeValue xmlCaseTaskNode, "USERID", xmlGetAttributeText(xmlTempRequest, "USERID")
    xmlSetAttributeValue xmlCaseTaskNode, "UNITID", xmlGetAttributeText(xmlTempRequest, "UNITID")
    xmlSetAttributeValue xmlCaseTaskNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlTempRequest, "USERAUTHORITYLEVEL")
     
    If DoUpgrade = True Then
       xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)
    Else
       xmlSetAttributeValue xmlCaseTaskNode, "TASKSTATUS", CStr(omiga4TASKSTATUSNotApplicable)
    End If
    
    xmlNewRequest.appendChild xmlCaseTaskNode.cloneNode(True)
    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    'if we are not upgrading then create a new action for a full credit check
    If DoUpgrade = False Then
    
        Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
        Dim xmlTMNewCaseTaskNode As IXMLDOMElement
        
        'create a new credit check task
        strTaskId = GetMandatoryGlobalParamString("TMExpNewCreditCheck")
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40

        xmlThisRequestDoc.validateOnParse = False
        xmlThisRequestDoc.setProperty "NewParser", True
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        xmlThisRequestDoc.appendChild xmlNewRequest

        Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
        xmlNewRequest.appendChild xmlTMNewCaseTaskNode
        xmlNewRequest.appendChild xmlApplicationNode.cloneNode(True)
        xmlNewRequest.setAttribute "OPERATION", "CREATEADHOCCASETASK"

        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "STAGEID"
        xmlCopyAttribute xmlCaseTaskNode, xmlTMNewCaseTaskNode, "CASESTAGESEQUENCENO"
        xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTaskId
        'xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
        'xmlSetAttributeValue xmlTMNewCaseTaskNode, "MANDATORYINDICATOR", "1"

        CreateAdhocCaseTask xmlNewRequest
         
    End If
    
RunXMLExpQuotationConversionCCExit:
   
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNewRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set objMsgTmBO = Nothing
    Set objExperianCC = Nothing
    Exit Sub
    
RunXMLExpQuotationConversionCCError:

    gobjContext.SetAbort
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    Set xmlNewRequest = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set objMsgTmBO = Nothing
    Set objExperianCC = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub

'LDM 23/03/2006  EP6
'This calls the Usp_CheckUpgradeTimeLimit stored procedure via the CRUD.
'To find out if the last creditcheck is within a timelimit
Private Function CheckUpdateConditions(xmlApplicationNode As IXMLDOMNode) As Boolean

    Const cstrFunctionName As String = "CheckUpdateConditions"
    On Error GoTo checkUpdateConditionsError
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlReturnDoc As FreeThreadedDOMDocument40
    Dim objCRUD As Object
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strWithinLimit  As String
    
    'Initialise
    gobjContext.SetComplete
    CheckUpdateConditions = False 'default
    
    Set xmlReturnDoc = New FreeThreadedDOMDocument40
    xmlReturnDoc.validateOnParse = False
    xmlReturnDoc.setProperty "NewParser", True
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    'xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True

    strApplicationNumber = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
        
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "CREDITCHECKUPGRADEDECISION"
    xmlElem.setAttribute "SCHEMA_NAME", "WebServiceSchema"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CREDITCHECKUPGRADEDECISION")
    xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNumber
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
    xmlNode.appendChild xmlElem
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        
    'if the DAYSSINCECHECK <= POSTARCHIVEDAY then WITHINLIMIT is 1; else WITHINLIMIT is 0
    'The result should come back something like underneath
    '<RESPONSE TYPE="SUCCESS">
    '    <CHECKUPGRADETIMELIMIT DAYSSINCECHECK="49" POSTARCHIVEDAY="90" WITHINLIMIT="1"/>
    '</RESPONSE>
    If Not xmlReturnDoc.loadXML(objCRUD.omRequest(xmlDoc.xml)) Then
        errThrowError cstrFunctionName, oeXMLParserError, xmlReturnDoc.parseError.reason
    End If
    
    If xmlReturnDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        errThrowError cstrFunctionName, oeUnspecifiedError, xmlReturnDoc.xml
    End If
    
    If xmlReturnDoc.selectSingleNode("RESPONSE/CHECKUPGRADETIMELIMIT") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving CHECKUPGRADETIMELIMIT details"
    End If
    
    Set xmlNode = xmlReturnDoc.selectSingleNode("RESPONSE/CHECKUPGRADETIMELIMIT")
    
    strWithinLimit = xmlGetAttributeText(xmlNode, "WITHINLIMIT")
    
    If strWithinLimit = "1" Then
        CheckUpdateConditions = True
    End If
 
        
checkUpdateConditionsExit:
    
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlReturnDoc = Nothing
    Set objCRUD = Nothing
    Exit Function

checkUpdateConditionsError:
    
    gobjContext.SetAbort
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlReturnDoc = Nothing
    Set objCRUD = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function

'LDM 23/03/2006  EP6 end

' ik_26/04/2006_EP465
Private Function GetNewLoan(ByVal vxmlRequest As IXMLDOMElement) As String
    
    Const cstrFunctionName As String = "GetNewLoan"
    On Error GoTo GetNewLoanExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim objCRUD As Object
    
    If Not vxmlRequest.selectSingleNode("APPLICATION[@APPLICATIONNUMBER][@APPLICATIONFACTFINDNUMBER]") Is Nothing Then
    
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
            
        Set xmlElem = xmlDoc.createElement("REQUEST")
        xmlElem.setAttribute "CRUD_OP", "READ"
        Set xmlNode = xmlDoc.appendChild(xmlElem)
        
        Set xmlElem = xmlDoc.createElement("NEWLOAN")
        xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequest.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", vxmlRequest.selectSingleNode("APPLICATION/@APPLICATIONFACTFINDNUMBER").Text
        xmlNode.appendChild xmlElem
    
        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        xmlDoc.loadXML objCRUD.omRequest(xmlDoc.xml)
        Set objCRUD = Nothing
    
        If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
            errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving NEWLOAN details"
        End If
        
        GetNewLoan = xmlDoc.xml
    
    End If
    
GetNewLoanExit:
        
    Set objCRUD = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function
' ik_26/04/2006_EP465_ends

' IK_18/05/2006_EP574
Private Sub BuildHandoverDoc( _
    ByVal vxmlRequestNode As IXMLDOMElement, _
    ByVal vxmlResponseNode As IXMLDOMElement)
    
    Const cstrFunctionName As String = "BuildHandoverDoc"
    On Error GoTo BuildHandoverDocExit
    
    Dim xmlDoc1 As DOMDocument40
    Dim xmlDoc2 As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlPi As IXMLDOMProcessingInstruction
    Dim comObj As Object
    Dim objMsgTmBO As Object
    Dim strResponse As String
    Dim strPath As String
    'IK_18/05/2006_EP672
    Dim strTransformer As String
    
    If vxmlRequestNode.selectSingleNode("APPLICATION[@APPLICATIONNUMBER]") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "no APPLICATIONNUMBER on REQUEST"
    End If
    
    If vxmlRequestNode.selectSingleNode("CASETASK[@TASKID]") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "no TASKID on REQUEST"
    End If
    
    'IK_18/05/2006_EP672
    GetHandoverParams strPath, strTransformer
    If Not IsHandoverLocationValid(strPath) Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "Handover File Location does not exist (" & strPath & ")"
    End If
        
    Set xmlDoc1 = New FreeThreadedDOMDocument40
    xmlDoc1.async = False
    xmlDoc1.setProperty "NewParser", True
        
    Set xmlElem = xmlDoc1.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "SCHEMA_NAME", "COMPLETIONS"
    xmlElem.setAttribute "ENTITY_REF", "APPLICATION"
    xmlElem.setAttribute "COMBOLOOKUP", "Y"
    xmlElem.setAttribute "COMBOTYPELOOKUP", "EX"
    Set xmlNode = xmlDoc1.appendChild(xmlElem)
        
    Set xmlElem = xmlDoc1.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
        
    Set comObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc1.loadXML comObj.omRequest(xmlDoc1.xml)
    Set comObj = Nothing
    
    If xmlDoc1.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving APPLICATION details"
    End If
    
    Set xmlDoc2 = New FreeThreadedDOMDocument40
    xmlDoc2.async = False
    xmlDoc2.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc2.createElement("REQUEST")
    xmlElem.setAttribute "OPERATION", "RUNCOMPLETIONRULES"
    Set xmlNode = xmlDoc2.appendChild(xmlElem)
    xmlNode.appendChild xmlDoc1.documentElement.firstChild.cloneNode(True)
    
    Set comObj = GetObjectContext.CreateInstance("omadminrules.adminrulesbo")
    xmlDoc1.loadXML comObj.omRequest(xmlDoc2.xml)
    Set comObj = Nothing
    
    If xmlDoc1.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error transforming APPLICATION details"
    End If
        
    If Len(strTransformer) <> 0 Then
        Set xmlPi = xmlDoc1.createProcessingInstruction("xml-stylesheet", "type='text/xsl' href='" & strTransformer & "'")
        xmlDoc1.insertBefore xmlPi, xmlDoc1.firstChild
    End If
    
    xmlDoc1.Save strPath & vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text & "_HandOver.xml"
                                  
    '==========================================================
    'Update the task as completed.
    xmlDoc1.loadXML ""
    Set xmlElem = xmlDoc1.appendChild(vxmlRequestNode.cloneNode(False))
    xmlElem.setAttribute "OPERATION", "UPDATECASETASK"
    Set xmlElem = xmlElem.appendChild(vxmlRequestNode.selectSingleNode("CASETASK").cloneNode(False))
    xmlElem.setAttribute "USERID", vxmlRequestNode.getAttribute("USERID")
    xmlElem.setAttribute "UNITID", vxmlRequestNode.getAttribute("UNITID")
    xmlElem.setAttribute "USERAUTHORITYLEVEL", vxmlRequestNode.getAttribute("USERAUTHORITYLEVEL")
    xmlElem.setAttribute "TASKSTATUS", CStr(omiga4TASKSTATUSComplete)

    Set objMsgTmBO = gobjContext.CreateInstance("msgTm.MsgTmBO")
    strResponse = objMsgTmBO.TmRequest(xmlDoc1.xml)
    Set objMsgTmBO = Nothing
    errCheckXMLResponse strResponse, True
    
BuildHandoverDocExit:
        
    Set comObj = Nothing
    Set objMsgTmBO = Nothing
    Set xmlPi = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc1 = Nothing
    Set xmlDoc2 = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'IK_18/05/2006_EP672
Private Sub GetHandoverParams(ByRef rstrPath As String, ByRef rstrTransformer As String)
    Const cstrFunctionName As String = "GetHandoverParams"
    On Error GoTo GetHandoverParamsExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim objCRUD As Object
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
        
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlDoc.appendChild xmlElem
        
    Set xmlElem = xmlDoc.createElement("OPERATION")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.documentElement.appendChild(xmlElem)
        
    Set xmlElem = xmlDoc.createElement("GLOBALPARAMETER")
    xmlElem.setAttribute "NAME", "HandoverDocLocation"
    xmlNode.appendChild xmlElem
        
    Set xmlElem = xmlDoc.createElement("OPERATION")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.documentElement.appendChild(xmlElem)
        
    Set xmlElem = xmlDoc.createElement("GLOBALPARAMETER")
    xmlElem.setAttribute "NAME", "HandoverDocTransformer"
    xmlNode.appendChild xmlElem
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc.loadXML objCRUD.omRequest(xmlDoc.xml)
    Set objCRUD = Nothing
    
    If Not xmlDoc.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='HandoverDocLocation'][@STRING]") Is Nothing Then
        rstrPath = xmlDoc.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='HandoverDocLocation']/@STRING").Text
        If Right(rstrPath, 1) <> "\" Then
            rstrPath = rstrPath & "\"
        End If
    End If
    
    If Not xmlDoc.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='HandoverDocTransformer'][@STRING]") Is Nothing Then
        rstrTransformer = xmlDoc.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='HandoverDocTransformer']/@STRING").Text
    End If
    
GetHandoverParamsExit:
        
    Set objCRUD = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Function IsHandoverLocationValid(ByVal vstrPath As String) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    If fso.FolderExists(vstrPath) Then
        IsHandoverLocationValid = True
    End If
    Set fso = Nothing
End Function
' IK_18/05/2006_EP574_ends

#End If

' PE 09/11/2006 - EP2_67
Private Sub RunVexValuationInstruction( _
    ByVal vxmlRequestXml As FreeThreadedDOMDocument40, _
    ByVal vxmlResponseXml As FreeThreadedDOMDocument40)
    
    Const cstrFunctionName As String = "RunVexValuationInstruction"
    On Error GoTo RunVexValuationInstructionExit

    Dim objVex As Object
    Dim strResponse As String
    
    ' Note that omVex is a .Net component
    Set objVex = CreateObject("omVex.omVexBO")
    strResponse = ""
    
    objVex.RunVexValuationInstruction vxmlRequestXml.xml, strResponse
    
    If strResponse <> "" Then
        vxmlResponseXml.loadXML strResponse
    End If

RunVexValuationInstructionExit:
            
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

' PE 09/11/2006 - EP2_67
Private Sub RunVexStatusUpdate( _
    ByVal vxmlRequestXml As FreeThreadedDOMDocument40, _
    ByVal vxmlResponseXml As FreeThreadedDOMDocument40)
    
    Const cstrFunctionName As String = "RunVexStatusUpdate"
    On Error GoTo RunVexStatusUpdateExit

    Dim objVex As Object
    Dim strResponse As String
    
    ' Note that omVex is a .Net component
    Set objVex = CreateObject("omVex.omVexBO")
    strResponse = ""
    
    objVex.RunVexStatusUpdate vxmlRequestXml.xml, strResponse
    
    If strResponse <> "" Then
        vxmlResponseXml.loadXML strResponse
    End If

RunVexStatusUpdateExit:
            
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

'EP2_1903 GHun
Private Function HandleInboundDocument(ByVal vxmlRequest As IXMLDOMElement)
    Const cstrFunctionName  As String = "HandleInboundDocument"
    Dim objCRUD             As Object
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlEntity           As IXMLDOMElement
    Dim xmlElem             As IXMLDOMElement
    Dim xmlResponse         As IXMLDOMElement

On Error GoTo HandleInboundDocumentErr

    Set xmlDoc = xmlCreateDOMObject()

    Set xmlRequest = xmlDoc.createElement("REQUEST")
    xmlRequest.setAttribute "CRUD_OP", "READ"
    xmlRequest.setAttribute "ENTITY_REF", "GETFULFILMENTEMAILMESSAGETYPE"
    xmlRequest.setAttribute "SCHEMA_NAME", "omTMSchema"
        
    Set xmlEntity = xmlDoc.createElement("GETFULFILMENTEMAILMESSAGETYPE")
    Set xmlElem = vxmlRequest.selectSingleNode("DOCUMENTLIST")
    If Not xmlElem Is Nothing Then
        xmlEntity.setAttribute "XML", xmlElem.xml
        xmlRequest.appendChild xmlEntity
        
        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        
        strResponse = objCRUD.omRequest(xmlRequest.xml)
        Set xmlDoc = xmlLoad(strResponse, cstrFunctionName)
        
        errCheckXMLResponseNode xmlDoc.documentElement, , True
    
        For Each xmlEntity In xmlDoc.selectNodes("RESPONSE/APPLICATION")
            Set xmlElem = xmlEntity.selectSingleNode("INTERFACE")
            
            If Not xmlElem Is Nothing Then
                Set xmlRequest = vxmlRequest.cloneNode(False)
                xmlRequest.setAttribute "OPERATION", "HANDLEINTERFACERESPONSE"
                xmlRequest.appendChild xmlEntity
                xmlRequest.appendChild xmlElem
                
                Set xmlResponse = xmlDoc.createElement("RESPONSE")
                HandleInterfaceResponse xmlRequest, xmlResponse
                'Do not check errors in xmlResponse as we want to continue processing all the message regardless
            End If
        Next
    End If

HandleInboundDocumentErr:

    Set objCRUD = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlEntity = Nothing
    Set xmlResponse = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
End Function
'EP2_1903 End

