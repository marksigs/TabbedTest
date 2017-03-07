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
'Workfile:      OmTmBOBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:   Task Manager Business Object.
'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    22/02/01    SYS1971 Add ImportAccountsIntoApplication
'APS    03/03/01    SYS1993 Add APPLICATIONPRIORITYVALUE tag to xml passed to CreateApplication method
'APS    04/03/01    SYS1920 Added IsApplicationAtOffer method
'JLD    02/03/01    SYS1879 Add DeclineApplication
'SR     08/03/01    New methods CreateDefaultApplicatinConditions, UpdateApplicationConditions
'APS    08/03/01    SYS1920 Amened IsApplicationAtOffer method
'APS    08/03/01    SYS1923 Issue Offer Processing
'PSC    09/03/01    SYS1879 Approve Recommend Application Processing
'SR     09/03/01    New Methods RunCreditCheck, RunRiskAssessment
'DJP    09/03/01    SYS1839 Added CreateValuationReport
'PSC    12/03/01    SYS1833 Added CompleteValuerInstructions
'JR     13/03/02    SYS1878 Added CreateReportOnTitle
'APS    15/03/01    SYS2076
'SR     15/03/01    SYS1924 - RunRiskAssessment
'APS    20/03/01    SYS2107
'SR     21/03/01    SYS2078 - Modified method 'CompleteValuerInstructions'
'DRC    26/03/01    SYS1932 Added GetTaskContactDetails
'PSC    27/03/01    SYS2053 Amend IssueOffer to remodel the quotation
'TJ     28/03/01    SYS1950 Removed mandatory check of TypeOfBuyer
'DRC    05/04/01    SYS1925 Modified AddDefaultValuesToCaseTask to ensure DueDate is a working day
'                           and set the OWINGUSERID and OWNINGUNITID
'PSC    09/04/01    SYS2053 Amend IssueOffer to get the mortgage sub quote number for the
'                           accepted quotation
'GD     12/04/01    SYS2092 Changes to CompleteValuerInstructions & AddDefaultValuesToCaseTask
'JR     20/04/01    SYS2048 Added ValidateReportOnTitle
'SR     24/04/01    SYS2169 - Do not delete the special coditions
'SR     25/04/01    SYS2194 - Add appropriate task while adding 'ValuerInstructions'
'GD     09/05/01    SYS2092 Added late binding to omApp
'SA     16/05/01    SYS2256 - Plugged in Validation methods for running rules.
'DRC    23/05/01    SYS2093 Changed interface for ValidateValuationReport to pass back a Response arguement
'SR     05/06/01    SYS2325 New method 'GetNewNumbers'
'DRC    19/07/01    SYS2308 Changed interface for CriticalDataCheck & ForwardBORequest to pass back a Response
'DRC    10/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'IK     30/10/01    SYS2865 fix to ProcessCarryForwardTasks
'JLD    14/11/01    SYS2839 added tags to pass thru' for internet core
'DRC    14/11/01    SYS2266 moved ProcessAutomaticTasks into MovetoStage, MoveToNextStage, CreateActivity
'  "      "          "      from AddDefaultsToCaseStage. Also fixes to ProcessAutomaticTasks and PrintDocumentForTasks
'PSC    20/11/01    SYS2867/SYS2988 Amend APPLICATION tag to RISKASSESSMENT
'PSC    26/11/01    SYS3195 Make calls to printing late bound
'PSC    27/11/01    SYS3221 Comment out call to NTTxBO to stop locking problems
'SA     29/11/01    SYS3270 Changed calls to GetApplicationData to use method in NtTxBo
'JR     03/12/01    SYS2459 (Omiplus24) Modified GetTaskContactDetails
'PSC    04/12/01    SYS3289 Amend MoveToNextStage and MoveToStage so that the call to omTmRequest sends
'                           down a string rather than an XML node
'BG     04/12/01    SYS3287 added [@TRIGGERTASKIND=""1""] condition to select the correct stagetask
'                           element in CreateCaseStageTrigger method.
'PSC    06/12/01    SYS3290 Correct SetComplete/SetAbort combinations
'BG     07/11/01    SYS3333 wrong global used in chaseuptask.
'DRC    10/12/01    SYS3295 Add Printing parameters & Printer Info to Issue Offer
'                           and Application & Customer Nodes to Interface of ProcessAutomaticTasks
'BG     11/11/01    SYS3458 Add CASEPRIORITY attribute to request for call in GetNextStageInSequence method
'BG     21/12/01    SYS3268 Added ReceivedOffer method.
'BG     21/12/01    SYS3268 Added extra functionality to IssueOffer method.
'DB     03/01/02    SYS3520 Added required primary key to ValidateReportOnTitle.
'DM     17/01/02    SYS3185 Made GetCurrentStage accessible through omTMRequest
'PSC    07/02/02    SYS3995 Amend IssueOffer to work correctly
'PSC    15/02/02    SYS3997 Amend CreateActivity to call GetApplicationData in omTmCommonBO
'DRC    18/04/02    SYS2675 Ammend GetTaskContactDetails to  use the TaskContactType combo group
'DR     25/04/02    SYS4467 Needed to add more print attributes for DMS to run from auto print tasks.
'CL     15/05/02    SYS4565 Changes to 'OmTmRequest' and addition of function 'DoClientRequest'
'DRC&CL 31/05/02    SYS4550 Added error handling for failed Automatic Tasks in MoveTo..Stage methods
'SR     20/02/2004  BBG54 - Modified method 'CreateApplication'
'------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'GHun   24/05/2002  BMIDS00005 CRWP2 CM065 Added UpdateCRSContactLog
'GHun   06/06/2002  BMIDS00026 IWP2 BM060 Added UpdateCRSCustomer
'MDC    10/06/2002  BMIDS00025 IWP1 BM059 Added RunHunterInterface
'GHun   14/06/2002  BMIDS00058 XML Node misnamed in UpdateCRSContacLog
'GHun   18/06/2002  BMIDS00026, BMIDS00005 Minor changes following code review
'MDC    28/06/2002  BMIDS00142 IWP1 BM059 Code review change
'MDC    17/07/2002  BMIDS00006 CAWP1 BM054 Customer Account Download
'MO     01/08/2002  BMIDS00218 IDWP1, Modified omTMRequest to expose CreateActivity
'MV     14/08/2002  BMIDS00333  Core Ref : SYS4919 ; Modified MoveToStage(); MoveToNextStage(); CreateActivity();
'                   CriticalDataCheck();RunUserRules()
'MDC    23/08/2002  BMIDS00336 CCWP1 BM062 - Credit Check and Full Bureau Download
'DPF    21/08/2002  BMIDS00344  APWP3 - new method called CreateValuationReportNoInst
'DPF    27/08/2002  BMIDS00344  APWP3 - add code to CompleteValuerInstruction to create
'                               Valuation Report (if possible, based upon existing report
'GD     03/10/2002  BMIDS00313  APWP2 - Amended CreateDefaultApplicationConditions
'BG     10/10/2002  BMIDS00612 add new indicator to ensure archiving happens for these prints.
'PSC    21/10/2002  BMIDS00676  Add new method RunGenericInterface
'PSC    24/10/2002  BMIDS00710  Amend RunGenericInterface to pass Customer Version Number
'PSC    28/10/2002  BMIDS00736  Amend RunGenericInterface to pass case task info
'DPF    06/11/2002  BMIDS00852  Applied a fix to CreateValuationReportNoInst to use Validation Type in a
'                               combo lookup and not the value id.
'GD     06/11/2002  BMIDS00743  Delete Standard AND Product conditions in CreateDefaultApplicationConditions
'MO     08/11/2002  BMIDS00752  When an error occurs in Hunter the error description is not displayed
'MO     11/11/2002  BMIDS00812  Made changes to ProcessCarryForwardTasks to carry forward tasks correctly!
'DB     13/11/2002  BMIDS00862  Removed any solicitors ROT validation as not needed for bmids.
'MO     15/11/2002  BMIDS00814  Created new method TransferApplicationOwnership
'DPF    19/11/2002  BMIDS00994  Amended ReinstatePreviousStage to check if all case tasks are complete
'                               before adding in UpdateCaseTask operation
'GD     20/11/2002  BMIDS00037  Changes to IsApplicationAtOfferStage.
'MO     25/11/2002  BMIDS01076  Change to MoveStage and MoveNextStage to raise proper warnings when auto tasks
'                               fail. Also had to change omTMRequest to pass the response node to the methods.
'MDC    29/11/2002  BMIDS01112  Ensure mandatory tasks are marked mandatory following Rules processing on MoveToStage
'PSC    05/12/2002  BM0167      Add UnapproveApplication Interface Task
'MDC    20/12/2002  BM0177      Check CaseTasks exist before attempting update in TransferApplicationOwnership
'AW     20/12/2002  BM0190      Use new node name PANELVALUERLISTLIST
'PSC    06/01/2003  BM0089      Amend MoveToNextStage and MoveToStage so that the completion of previous
'                               stage and creation of new stage are in a differnt transaction to performing
'                               Automatic tasks
'PSC    08/01/2003  BM0138      Amend ProcessAutomaticTasks to not check for the printer details on an interface task as
'                               this should be done within the task processing itself
'BS     07/02/2003  BM0136  Amend MoveToNextStageNewTx to move a case to an exception stage as determined
'                               by the user defined omTmRules
'MV     27/02/2003  BM0362      Amended ValidateReportOnTitle()
'GHun   10/03/2003  BM0376      Amended ImportBureauData to call RunIncomeCalcs on completion
'DB     31/03/2003  BM0248      CreateAdHocTask altered to add mandatory flag.
'INR    08/05/2003  BM0553      Assigned userid to a task not displayed correctly.
'MDC    08/05/2003  BM0340      Performance enhancements in MoveToNextStage and MoveToStage
'GD     15/05/2003  BM0501      Fix to error handling in AddDefaultValuesToCaseTask
'GD     19/05/2003              Reinstated latest version
'INR    09/07/2003  BMIDS593    When we raise a oeTmCCOKBureauImportFailed or oeTmBureauDataImportFailed error,
'                               The original error description is written to the event log.
'                               Add to the error message to notify the user.
'                               This version done to exclude BMIDS585 changes for promotion.
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BBG History
'
'Prog   Date        Description
'MV     17/03/2004  BBG107 - Created New method CreateQuestReportData() and Amended omTmRequest()
'TK     07/04/2004  BBG168 - Populate LTDCompanyIndicator in Application Table.
'TK     01/06/2004  BBG406 - Added RunTransactCreditSearch() method
'TK     21/06/2004  BBG609 - Amended IssueOffer()
'TK     30/06/2004  BBG895 - Amend CreateQuestReportData() and CreateValuationReportNoInst()
'MSla   30/06/2004  BBG882  Amended validation of ROT Legal Rep
'TK     27/07/2004  E2EM00000113 - Added RunAutomaticTransactDecision() method
'PSC    29/07/2004  BBG1144 - Amend IssueOffer and PrintDocumentForTask to send down additional
'                             TEMPLATEDATA if required
'TK     13/08/2004  BBG406 - Amend RunTransactCreditSearch()
'PSC    20/08/2004  BBG1233 - Add OverturnKFI(). Amend PrintDocumentForTask to determine whether to
'                             send the document to the web
'TK     26/08/2004  E2EM00001205 - Amended CreateDefaultApplicationConditions()
'MV     10/09/2004  BBG1379 Amended PrintDocumentForTask()
'TK     14/09/2004  BBG1415 Amend IssueOffer()
'TK     15/09/2004  BBG1207 Amend CriticalDataCheck() and add TransactSummaryCheckRequired()
'TK     16/09/2004  BBG1434 Amend IssueOffer()
'MV     24/09/2004  E2EM00001614    Amended RunHunterInterface()
'MV     27/09/2004  E2EM00001614    Amended RunHunterInterface() to Include TaskDueDate Column
'MV     29/09/2004  E2EM00001614    Amended RunHunterInterface() to RoleBack all the changes made
'SJha   01/10/2004  BBG1176         Amended PrintDocumentForTask()
'PSC    06/10/2004  BBG1572         Add attribute to determine what kind of document is being printed
'MV     11/10/2004  BBG1207 - Amended CriticalDataCheck(); added IsTaskCompleted() ; IsTaskPresent() ; GetPackaged()
'MV     20/10/2004  BBG2246 - Amended ValidateReportOnTitle() and others tidy up
'TK     25/10/2004  BBG1697 - Amend IssueOffer()
'BT     26/10/2004  BBG1707 - Remove the COPIES field from the KFI data
'PSC    05/11/2004  BBG1772 - Amend PrintDocumentForTask to take into account if application is packaged
'TK     11/11/2004  E2EM00002757 - Amend IsApplicationAtOfferStage()
'TK     15/11/2004  E2EM00002757 - Amend IsApplicationAtOfferStage()
'PSC    23/11/2004  E2EM00002747 - Amend IsApplicationAtOfferStage() to use sequence numbers
'TK     08/12/2004  E2EM00003094 - Amend CriticalDataCheck() and performance fixes
'PSC    09/12/2004  E2EM00003099 - Amend ProcessCarryForwardTasks to copy correct attributes and
'                                  Amend IsApplicationAtOfferStage vto cater for no archive tasks
'SDS    10/12/2004  BBG1741 - Task Management - Move Stage is taking a long time when performed against a large DB
'SDS    15/12/2004  E2EM00003099 - Amend CriticalDataCheck to create task, if it doesn't exist in Incomplete state
'TK     18/01/2005  E2EM00003094 - Changes to CriticalDataCheck()
'MV     21/04/2005  CORE106 - Amended OverTurnKFI(),IssueOffer(),ProcessAutomaticTasks()
'MV     25/04/2005  CORE107 - Amended IssueOffer(),ProcessAutomaticTasks(), OverturnKFI()
'AS     06/05/2005  CORE117 - Undo CORE106/CODE107, and implement solution that uses combo box text (prevents problems elsewhere).
'AS     25/05/2005  CORE137 omTM must specify compression method when calling omPrint to generate a document
'AS     25/07/2005  CORE174 omTM cannot create an ad hoc task for anyone except the logged in user.
'                           AddDefaultValuesToCaseTask moved to omTMFunctions.bas.
'                           SetCaseTaskDueDateAndTime moved to omTMFunctions.bas.
'------------------------------------------------------------------------------------------
'CORE History
'Prog   Date        Defect     Description
'SR     25/01/2006  BBGRb74    new methods - SubmitProduct, TransferApplicationOwnershipOnSubmit
'                              AddDefaultProductTasks, UpdateCaseTaskReferenceNumber.
'                              modified method - CreateAdhocCaseTask
'SR     24/03/2006  CORE258    BBGRb604 modified methods AddDefaultProductTasks, UpdateCaseTaskReferenceNumber
'SR     26/03/2006  CORE258    BBGRb604 modified method SubmitProduct
'SR     29/03/2006  CORE258    BBGRb604/799
'AS     21/07/2006  CORE285    BBGRb1298 INR Add default stagetask info when creating adhoc tasks.
'AS     09/08/2006  CORE285    BBGRb1298 INR Add default stagetask info for mortgage tasks as well.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Private gobjContext As ObjectContext
Private gobjTmBO As Object
Private Const cstrTaskComplete = 40

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub
Private Sub CreateQuestReportData(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo CreateQuestReportDataExit
    Const cstrFunctionName As String = "CreateQuestReportData"
    
    Dim strResponse  As String
    Dim lngErrNo As Long
    Dim strErrDescription As String
    Dim strValidationType As String
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlNewRequestValuationNode As IXMLDOMNode 'BBG895
    Dim objBO As Object
    
    Set objBO = gobjContext.CreateInstance("omQuest.QuestBO")
    strResponse = objBO.CreateQuestReportData(vxmlRequestNode.xml)
    
    'On Error Resume Next 'BBG895
    
    errCheckXMLResponse strResponse, True
    lngErrNo = Err.Number
    strErrDescription = Err.Description
    
    On Error GoTo CreateQuestReportDataExit
    
    ' Get the task status for validation type "CP"
    If lngErrNo <> 0 Then
        strValidationType = GetFirstComboValueId("TaskStatus", "I")
    Else
        'TK 30/06/2004 BBG895 Start
        strResponse = ""
        Set xmlNewRequestValuationNode = vxmlRequestNode.cloneNode(True)
        Call CreateValuationReportNoInst(xmlNewRequestValuationNode)

        'errCheckXMLResponse strResponse, True
        lngErrNo = Err.Number
        strErrDescription = Err.Description
        On Error GoTo CreateQuestReportDataExit

        If lngErrNo = 0 Then
            strValidationType = GetFirstComboValueId("TaskStatus", "CP")
        End If
        'TK 30/06/2004 BBG895 End
    End If
    
    ' Now update Case Task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    Set xmlCaseTaskElem = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
    xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", strValidationType
    xmlNewRequest.appendChild xmlCaseTaskElem
    
    Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    Exit Sub
    
CreateQuestReportDataExit:

    Set objBO = Nothing
    Set xmlNewRequest = Nothing 'BBG895
    Set xmlCaseTaskElem = Nothing 'BBG895
    Set xmlNewRequestValuationNode = Nothing 'BBG895
    errCheckError cstrFunctionName

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
    
    Dim strOperation As String
'    Dim strAttribValue As String
   
    Set gobjTmBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    
    If gobjTmBO Is Nothing Then
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
    
    ' parse request node
    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    xmlCheckMandatoryAttribute xmlRequestNode, "OPERATION"
    strOperation = xmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    strOperation = UCase$(strOperation)
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
        'SDS  BBG1741  10/12/2004__START
        'Case "MOVETONEXTSTAGE"
        '    MoveToNextStage xmlRequestNode
        'Case "MOVETOSTAGE"
        '    MoveToStage xmlRequestNode
        'SDS  BBG1741  10/12/2004__END
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
            PrintDocumentForTask xmlRequestNode, xmlResponseNode
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
        'SDS  BBG1741  10/12/2004__START
        'Case "DECLINEAPPLICATION"
        '    DeclineApplication xmlRequestNode
        'SDS  BBG1741  10/12/2004__END
        Case "CREATEDEFAULTAPPLICATIONCONDITIONS"
            CreateDefaultApplicationConditions xmlRequestNode
        Case "UPDATEAPPLICATIONCONDITIONS"
            UpdateApplicationConditions xmlRequestNode
                        
        ' APS SYS1923
        Case "ISSUEOFFER"
            IssueOffer xmlRequestNode, xmlResponseNode
        ' PSC SYS1879
        Case "APPROVERECOMMENDAPPLICATION"
            ApproveRecommendApplication xmlRequestNode, xmlResponseNode
        Case "RUNCREDITCHECK"
            RunCreditCheck xmlRequestNode
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
        Case "PROCESSAUTOMATICTASKS"
             ProcessAutomaticTasks xmlRequestNode
                     
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
        'SDS  BBG1741  10/12/2004__START
        'PSC 06/01/2003 BM0089 - Start
        ' This is a kludge to get around the transactioning errors
        'Case "MOVETONEXTSTAGENEWTX"
        '    MoveToNextStageNewTx xmlRequestNode, xmlResponseNode
        'Case "MOVETOSTAGENEWTX"
        '    MoveToStageNewTx xmlRequestNode, xmlResponseNode
        'PSC 06/01/2003 BM0089 - End
        'SDS  BBG1741  10/12/2004__END
        Case "CREATEQUESTREPORTDATA"
            CreateQuestReportData xmlRequestNode
        'TK 01/06/2004 BBG406 - Start
        Case "RUNTRANSACTCREDITSEARCH"
            RunTransactCreditSearch xmlRequestNode
        'TK 27/07/2004 E2EM00000113 - start
        Case "RUNTRANSACTFULLDECISION"
            RunTransactFullDecision xmlRequestNode
        'PSC 20/08/2004 BBG1233 - Start
        Case "OVERTURNKFI"
            OverturnKFI xmlRequestNode, xmlResponseNode
        'PSC 20/08/2004 BBG1233 - End
        Case "UPDATECASETASKREFERENCENUMBER"
            UpdateCaseTaskReferenceNumber xmlRequestNode, xmlResponseNode
        Case "ADDDEFAULTPRODUCTTASKS"
            AddDefaultProductTasks xmlRequestNode, xmlResponseNode
        Case "TRANSFERAPPLICATIONOWNERSHIPONSUBMIT"
            TransferApplicationOwnershipOnSubmit xmlRequestNode, xmlResponseNode
        Case "SUBMITPRODUCT"
            SubmitProduct xmlRequestNode, xmlResponseNode
        Case "UPDATEAPPSTATUSANDSUBMITPRODUCT"
            UpdateAppStatusAndSubmitProduct xmlRequestNode, xmlResponseNode
        Case "UPDATEAPPLICATIONSTATUS"
            UpdateApplicationStatus xmlRequestNode
        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, "OPERATION = " & strOperation
        ' PSC SYS1971 - End
    End Select
        
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
        
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
    GoTo OmTmRequestExit
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
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objApplManagerBO As Object, objApplicationBO As Object, objAdminInterfaceBO As Object
    Dim objCustomerBO As Object, objMsgTmBO As Object
    'Build Reuqest and Get Application data
    Set xmlDoc = xmlCreateDOMObject
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
    Set xmlNewRequest = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
GetNewNumbersErr:
    gobjContext.SetComplete
    GoTo GetNewNumbersExit
End Sub
Private Sub CreateDefaultApplicationConditions(ByVal vxmlRequestNode As IXMLDOMNode)
'GD BMIDS00313
On Error GoTo CreateDefaultApplicationConditionsErr
    Const cstrFunctionName As String = "CreateDefaultApplicationConditions"
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlNewRequest As IXMLDOMNode, xmlTableNode As IXMLDOMElement
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    ' validate REQUEST
    Set xmlTableNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    strApplicationNumber = xmlGetAttributeText(xmlTableNode, "APPLICATIONNUMBER")
    strApplicationFFNumber = xmlGetAttributeText(xmlTableNode, "APPLICATIONFACTFINDNUMBER")
    If strApplicationNumber = "" Then
        errThrowError cstrFunctionName, oeMissingElementValue, "Application Number"
    End If
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    Dim objBO As Object
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    ' Build the request and call 'Conditions.DeleteAllApplicationConditions'
    Set xmlDoc = xmlCreateDOMObject
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "DELETEALLAPPLICATIONCONDITIONS"
    Set xmlTableNode = xmlDoc.createElement("APPLICATIONCONDITIONS")
    xmlSetAttributeValue xmlTableNode, "APPLICATIONNUMBER", strApplicationNumber
    'GD BMIDS00313 added AppFFNum.
    xmlSetAttributeValue xmlTableNode, "APPLICATIONFACTFINDNUMBER", strApplicationFFNumber
    'GD BMIDS00743  START - Make 2 calls to DeleteAllApplicationConditions
    ' because adoDelete doesn't allow OR's, or any strSQL to be passed to it (like the Get methods do)
'    ' SR 24/04/01 : SYS2169 - Do not delete the special coditions
'    xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", 10
'    ' SR 24/04/01 : End
    'TK 26/08/2004 E2EM00001205 Start - Delete all application condition records for each condition types
'    Dim colValueIDs As Collection
'    Dim strValueID As String
'    'Set strStandardCondType = Get ValueID for ValidationType 'STD' (Standard Condition)
'    Set colValueIDs = New Collection
'    GetValueIdsForValidationType "ConditionType", "STD", colValueIDs
'    'Extract first value id
'    strValueID = colValueIDs.Item(1)
'    xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", strValueID
    'TK 26/08/2004 End
    xmlNewRequest.appendChild xmlTableNode
    'Call Delete for the first time
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML (objBO.omAppProcRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    'Set strProductCondType = Get ValueID for ValidationType 'P' (Product Condition)
    'TK 26/08/2004 E2EM00001205 - Start
'    Set colValueIDs = New Collection
'    GetValueIdsForValidationType "ConditionType", "P", colValueIDs
'    'Extract first value id
'    strValueID = colValueIDs.Item(1)
'    xmlSetAttributeValue xmlTableNode, "CONDITIONTYPE", strValueID
'    'Call Delete for the second time
'    xmlThisResponseDoc.setProperty "NewParser", True
'    xmlThisResponseDoc.validateOnParse = False
'    xmlThisResponseDoc.loadXML (objBO.omAppProcRequest(xmlNewRequest.xml))
'    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    'TK 26/08/2004 End
    'GD BMIDS00743  END
    'Build the request and call 'ConditionsBO.CreateDefaultApplicationConditions'
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEDEFAULTAPPLICATIONCONDITIONS"
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML (objBO.omAppProcRequest(xmlNewRequest.xml))
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    'Update conditions task as pending
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    xmlNewRequest.appendChild vxmlRequestNode.selectSingleNode(".//CASETASK")
    Set xmlTableNode = xmlNewRequest.selectSingleNode(".//CASETASK")
    xmlSetAttributeValue xmlTableNode, "TASKSTATUS", 20
    Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
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
    'TK 14/09/2004 BBG1415 Dim xmlThisResponseNode As IXMLDOMNode
    Dim xmlApplicationRequestNode As IXMLDOMNode
    Dim xmlCustomerRequestNode As IXMLDOMNode
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
    Set xmlThisRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisRequestDoc.async = False
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
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "CHANNELID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "CHANNELID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "USERID"
    xmlElemFromAttrib xmlThisRequestNode, vxmlRequestNode, "UNITID"
    ' APS SYS1993
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "APPLICATIONPRIORITYVALUE"
    
    ' TK 07/04/2004 : BBG168 - New column on Application LTDCompanyIndicator
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LTDCOMPANYINDICATOR"
    
    ' SR 20/02/2004 : BBG54 - New column on AFF LevelOfAdvice
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LEVELOFADVICE"
    
    ' JLD SYS2839 required for internet system
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOCKING"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "RESIDENCYOKINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "AGEOKINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOANAMOUNTINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "LOANTERMINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "PRIMARYRESIDENCEINDICATOR"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "MARKETINGSOURCE"
    xmlElemFromAttrib xmlThisRequestNode, xmlApplicationRequestNode, "INTENDEDAPPLICANTCOUNT"
    ' run CreateApplicationAndCustomer ==============================================
    Debug.Print "CreateApplicationAndCustomer request:"
    Debug.Print xmlThisRequestDoc.xml
    Set xmlThisResponseDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
    Set objBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
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
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlApplicationRequestNode = Nothing
    Set xmlCreateApplicationAndCustomerResponseNode = Nothing
    Set xmlElem = Nothing
    Set xmlSrceNode = Nothing
    Set xmlNode = Nothing
    Set objBO = Nothing
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
    Set xmlNextStageDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlNextStageDoc.async = False
    GetNextStage vxmlRequestNode, xmlCaseActivityRequestNode, xmlNextStageDoc
    Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
        
    If AnyRulesForStage(xmlNextStageNode) = True Then
        'Get APPLICATION data
        Set xmlApplicationDataDoc = xmlCreateDOMObject
        'TK 14/09/2004 BBG1415 xmlApplicationDataDoc.async = False
        Set objCommonBO = gobjContext.CreateInstance(App.Title & ".OmTmCommonBO")
        objCommonBO.GetApplicationData xmlApplicationRequestNode, xmlApplicationDataDoc, "TMRBTemplate"
        Set objCommonBO = Nothing
        Set xmlRbResponseNode = xmlApplicationDataDoc.selectSingleNode("RESPONSE")
        'Run rules
        If xmlNextCaseStageDoc Is Nothing Then
            Set xmlNextCaseStageDoc = xmlCreateDOMObject
            'TK 14/09/2004 BBG1415 xmlNextCaseStageDoc.async = False
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
            Set xmlNextCaseStageDoc = xmlCreateDOMObject
            'TK 14/09/2004 BBG1415 xmlNextCaseStageDoc.async = False
        Else
            xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
        End If
        CreateCaseStageFromStage _
            xmlNextStageNode, _
            xmlNextCaseStageDoc
        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
    End If
    'Run CreateActivity request
    AddDefaultValuesToCaseStage vxmlRequestNode, xmlNextCaseStageNode, xmlNextStageNode
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
    Set objCommonBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

'SDS  BBG1741  10/12/2004__START
'Private Sub MoveToStage(ByVal vxmlRequestNode As IXMLDOMNode)
'
'    On Error GoTo MoveToStageExit
'    Const cstrFunctionName As String = "MoveToStage"
'    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
'    'TK 14/09/2004 BBG1415 Dim xmlCreateCaseStageDoc As FreeThreadedDOMDocument40
'    Dim xmlApplicationRequestNode As IXMLDOMNode
'    Dim xmlCurrentStageNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlNextStageNode As IXMLDOMNode
'    Dim xmlRequestCaseStageNode As IXMLDOMNode
'    Dim xmlNodeList As IXMLDOMNodeList
'    Dim xmlApplicationNode As IXMLDOMNode
'    Dim xmlPrinterNode As IXMLDOMNode
'    Dim xmlAutoCaseTaskNode As IXMLDOMNode
'    Dim xmlAutoRequestNode As IXMLDOMNode
'    Dim xmlAutoRequestDoc As FreeThreadedDOMDocument40
'    Dim objNTxBO As IOmTmNTxBO
'    Dim strXSL As String
'    Dim sResponse As String
'    ' validate request
'    Set xmlRequestCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
'    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "SOURCEAPPLICATION"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "CASEID"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "ACTIVITYID"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "STAGEID"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
'    'PSC 06/01/2003 BM0089 - Start
'    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
'    sResponse = objNTxBO.MoveToStage(vxmlRequestNode)
'    ' get New current (i.e. updated) CASESTAGE detail =====================================================
'    Set xmlCurrentStageDoc = xmlLoad(sResponse, cstrFunctionName)
'    'PSC 06/01/2003 BM0089 - End
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    Debug.Print xmlCurrentStageNode.xml
'    strXSL = "CASETASK[@AUTOMATICTASKIND='1']"
'    Set xmlNodeList = xmlCurrentStageNode.selectNodes(strXSL)
'    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
'    Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("./PRINTER")
''
'    'AQR SYS4550 - DRC & CPL    ' Error Check
'    '
'    Dim sTempTaskName As String
''    Dim xmlTempNode As IXMLDOMNode
'    Dim iAutoErrorCount As Integer
'    On Error Resume Next
'    sTempTaskName = ""
'    For Each xmlAutoCaseTaskNode In xmlNodeList
'        Set xmlAutoRequestDoc = New FreeThreadedDOMDocument40
'        xmlAutoRequestDoc.async = False
'        Set xmlAutoRequestNode = xmlAutoRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
'        xmlAutoRequestNode.Attributes.getNamedItem("OPERATION").Text = "ProcessAutomaticTasks"
'        xmlAutoRequestNode.appendChild xmlAutoCaseTaskNode
'        xmlAutoRequestNode.appendChild xmlApplicationNode
'        xmlAutoRequestNode.appendChild xmlPrinterNode
'
'        ' PSC 04/12/01 SYS3289
'        'PSC 06/01/2003 BM0089 - Start
'        objNTxBO.ProcessAutomaticTasks xmlAutoRequestNode
'        If Err.Number <> 0 Then
'            Err.Clear
'            iAutoErrorCount = iAutoErrorCount + 1
'            sTempTaskName = sTempTaskName & xmlAutoCaseTaskNode.Attributes.getNamedItem("TASKID").xml
'        End If
'        'PSC 06/01/2003 BM0089 - End
'    Next
'    On Error GoTo MoveToStageExit
'    'END AQR SYS2266 DRC
''
''AQR 4550 check the error count in the process automatic task loop and re-raise an error
''
'    'BMIDS01076 MO 25/11/2002 - Change to raise a proper error number and as a warning
'
'    'If iAutoErrorCount > 0 Then
'    '    Err.Raise _
'    '    oeUnspecifiedError, _
'    '    TypeName(Me) & "." & cstrFunctionName, _
'    '    "WARNING " & iAutoErrorCount & " Automatic  Task(s) were not completed :-  " & sTempTaskName
'    'End If
'    If iAutoErrorCount > 0 Then
'        errThrowError cstrFunctionName, oeTmAutomaticTaskDidntComplete, "", iAutoErrorCount, sTempTaskName
'    End If
'MoveToStageExit:
'    Set xmlCurrentStageDoc = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlCreateCaseStageDoc = Nothing
'    Set xmlApplicationRequestNode = Nothing
'    Set xmlCurrentStageNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlNextStageNode = Nothing
'    Set xmlRequestCaseStageNode = Nothing
'    Set xmlNodeList = Nothing
'    Set xmlApplicationNode = Nothing
'    Set xmlPrinterNode = Nothing
'    Set xmlAutoCaseTaskNode = Nothing
'    Set xmlAutoRequestNode = Nothing
'    Set xmlAutoRequestDoc = Nothing
'    Set objNTxBO = Nothing
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub
'Private Sub MoveToNextStage(ByVal vxmlRequestNode As IXMLDOMNode)
'    On Error GoTo MoveToNextStageExit
'    Const cstrFunctionName As String = "MoveToNextStage"
'    Dim xmlCurrentStageRequestNode As IXMLDOMNode
'    Dim xmlApplicationRequestNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
'    'TK 14/09/2004 BBG1415 Dim xmlCurrentStageNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlNodeList As IXMLDOMNodeList
'    'TK 14/09/2004 BBG1415 Dim xmlApplicationNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlPrinterNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlAutoCaseTaskNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlAutoRequestNode As IXMLDOMNode
'    'TK 14/09/2004 BBG1415 Dim xmlAutoRequestDoc As FreeThreadedDOMDocument40
'    Dim objNTxBO As IOmTmNTxBO
'    'TK 14/09/2004 BBG1415 Dim strXSL As String
'    Dim sResponse As String
'    ' validate request
'    Set xmlCurrentStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTSTAGE")
'    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "SOURCEAPPLICATION"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "CASEID"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "ACTIVITYID"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "STAGEID"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
'
'    ' PSC 06/01/2003 BM0089 - Start
'    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
'    sResponse = objNTxBO.MoveToNextStage(vxmlRequestNode)
'
'    'BM0340 MDC 22/04/2003
''    ' get New current (i.e. updated) CASESTAGE detail =====================================================
''    Set xmlCurrentStageDoc = xmlLoad(sResponse, cstrFunctionName)
''    ' PSC 06/01/2003 BM0089 - End
''
''    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
''    Debug.Print xmlCurrentStageNode.xml
''    strXSL = "CASETASK[@AUTOMATICTASKIND='1']"
''    Set xmlNodeList = xmlCurrentStageNode.selectNodes(strXSL)
''    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
''    Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("./PRINTER")
''    '
''    'AQR SYS4550 - DRC & CPL    ' Error Check
''    '
''    Dim sTempTaskName As String
''    Dim xmlTempNode As IXMLDOMNode
''    Dim iAutoErrorCount As Integer
''
''    On Error Resume Next
''    sTempTaskName = ""
''    For Each xmlAutoCaseTaskNode In xmlNodeList
''        Set xmlAutoRequestDoc = New FreeThreadedDOMDocument40
''        xmlAutoRequestDoc.async = False
''        Set xmlAutoRequestNode = xmlAutoRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
''        xmlAutoRequestNode.Attributes.getNamedItem("OPERATION").Text = "ProcessAutomaticTasks"
''
''        xmlAutoRequestNode.appendChild xmlAutoCaseTaskNode
''        xmlAutoRequestNode.appendChild xmlApplicationNode
''        xmlAutoRequestNode.appendChild xmlPrinterNode
''
''        ' PSC 04/12/01 SYS3289
''        'PSC 06/01/2003 BM0089 - Start
''        objNTxBO.ProcessAutomaticTasks xmlAutoRequestNode
''
''        If Err.Number <> 0 Then
''            Err.Clear
''            iAutoErrorCount = iAutoErrorCount + 1
''            sTempTaskName = sTempTaskName & xmlAutoCaseTaskNode.Attributes.getNamedItem("TASKID").xml
''        End If
''        'PSC 06/01/2003 BM0089 - End
''    Next
''
''    On Error GoTo MoveToNextStageExit
''    'END AQR SYS2266 DRC
''        '
'''AQR 4550 check the error count in the process automatic task loop and re-raise an error
'''
''    'BMIDS01076 MO 25/11/2002 - Change to raise a proper error number and as a warning
''
''        'If iAutoErrorCount > 0 Then
''        '    Err.Raise _
''        '    oeUnspecifiedError, _
''        '    TypeName(Me) & "." & cstrFunctionName, _
''        '    "WARNING " & iAutoErrorCount & " Automatic  Task(s) were not completed :-  " & sTempTaskName
''        'End If
''
''    If iAutoErrorCount > 0 Then
''        errThrowError cstrFunctionName, oeTmAutomaticTaskDidntComplete, "", iAutoErrorCount, sTempTaskName
''    End If
'    'BM0340 MDC 22/04/2003 - End
'MoveToNextStageExit:
'    Set xmlCurrentStageRequestNode = Nothing
'    Set xmlApplicationRequestNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlCurrentStageDoc = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlCurrentStageNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlNodeList = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlApplicationNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlPrinterNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlAutoCaseTaskNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlAutoRequestNode = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlAutoRequestDoc = Nothing
'    Set objNTxBO = Nothing
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub
'SDS  BBG1741  10/12/2004__END

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
    xmlThisRequestDoc.async = False
    Set xmlPreviousStageDoc = New FreeThreadedDOMDocument40
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
    Dim strUserId As String, strUnitId As String ', strResponse As String
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
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.loadXML objBO.omAppProcRequest(xmlNewRequest.xml)
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
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
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
    xmlCurrentStageDoc.async = False
    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    ' create GetStageDetail request ====================================================
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
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
    xmlStageTaskDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlStageTaskDoc
    Set xmlStageNode = xmlStageTaskDoc.selectSingleNode("RESPONSE/STAGE")
    If xmlStageNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoTaskDetail
    End If
    Set xmlOutstandingCaseTaskList = _
        xmlCurrentStageNode.selectNodes("CASETASK[@TASKSTATUS='10'  or  @TASKSTATUS='20']")
    If xmlOutstandingCaseTaskList.length <> 0 Then
        strPriority = xmlGetAttributeText(xmlPriorityRequestNode, "APPLICATIONPRIORITY")
        strXSLPattern = "[@CASEPRIORITY='" & strPriority & "']"
        ' create UpdateCaseTask request(s) =============================================
        ' drop existing (GetStageTaskList) request
        xmlThisRequestNode.removeChild xmlThisRequestNode.firstChild
        ' set OPERATION attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
        For Each xmlOutstandingCaseTaskNode In xmlOutstandingCaseTaskList
            strXSLPattern = _
                "STAGETASK[@TASKID='" & _
                xmlGetAttributeText(xmlOutstandingCaseTaskNode, "TASKID") & _
                "']" & _
                "[@CASEPRIORITY='" & strPriority & "']" & _
                "[@ADJUSTMENTDAYS  !=  '0'  or  @ADJUSTMENTHOURS  !=  '0']"
            Set xmlStageTaskNode = xmlStageNode.selectSingleNode(strXSLPattern)
            If Not xmlStageTaskNode Is Nothing Then
                dtDueDateTime = _
                    xmlGetAttributeAsDate(xmlOutstandingCaseTaskNode, "TASKDUEDATEANDTIME")
                lngOffset = _
                    (xmlGetAttributeAsInteger(xmlStageTaskNode, "ADJUSTMENTDAYS") * 24) + _
                    xmlGetAttributeAsInteger(xmlStageTaskNode, "ADJUSTMENTHOURS")
                If lngOffset <> 0 Then
                    dtDueDateTime = DateAdd("h", lngOffset, dtDueDateTime)
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
    Set gobjEnquiry = Nothing
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
    'TK 14/09/2004 BBG1415 Dim xmlChaseTaskDetailNode As IXMLDOMNode
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
    Set xmlThisRequestDoc = xmlCreateDOMObject
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
    Set xmlThisCaseTaskDetailDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisCaseTaskDetailDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlThisCaseTaskDetailDoc
    Set xmlThisCaseTaskDetailNode = _
        xmlThisCaseTaskDetailDoc.selectSingleNode("RESPONSE/CASETASK")
    If xmlThisCaseTaskDetailNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoTaskDetail
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
    Set xmlChaseTaskTemplateDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlChaseTaskTemplateDoc.async = False
    ' run Tm GetCaseTasks request
    RunTmRequest xmlThisRequestDoc, xmlChaseTaskTemplateDoc
    Set xmlChaseTaskTemplateNode = _
        xmlChaseTaskTemplateDoc.selectSingleNode("RESPONSE/TASK")
    If xmlChaseTaskTemplateNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoTaskDetail
    End If
    ' create UpdateCaseTask request ====================================================
    ' create chase-up task from new CaseTask details plus
    ' attributes from existing casetask
    Set xmlUpdateCaseTaskDoc = xmlCreateDOMObject
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
    ' create CreateCaseTask request ====================================================
    ' create chase-up task from new CaseTask details plus
    ' attributes from existing casetask
    Set xmlCreateCaseTaskDoc = xmlCreateDOMObject
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
    Set xmlUpdateCaseTaskDoc = Nothing
    Set xmlCreateCaseTaskDoc = Nothing
    Set xmlChaseTaskTemplateDoc = Nothing
    Set xmlChaseTaskTemplateNode = Nothing
    Set xmlUpdateCaseTaskRequestNode = Nothing
    Set xmlCreateCaseTaskRequestNode = Nothing
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
    'TK 14/09/2004 BBG1415 Dim xmlNode As IXMLDOMNode
    ' validate request ================================================================
    Set xmlCaseTaskRequestNode = vxmlRequestNode.selectSingleNode("CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "ACTIVITYINSTANCE"
    xmlCheckMandatoryAttribute xmlCaseTaskRequestNode, "STAGEID"
    ' create GetStageTaskList request ==================================================
    Set xmlThisRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisRequestDoc.async = False
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
    Set xmlTriggerTaskDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlTriggerTaskDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlTriggerTaskDoc
    'BG 04/11/01 SYS3287 added [@TRIGGERTASKIND=""1""] condition to select the correct stagetask element
    Set xmlTriggerTaskNode = xmlTriggerTaskDoc.selectSingleNode("RESPONSE/STAGETASK[@TRIGGERTASKIND=""1""]")
    If xmlTriggerTaskNode Is Nothing Then
        errThrowError cstrFunctionName, oeTmNoTaskDetail
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
    xmlThisRequestNode.appendChild xmlElem
    ' run TmRequest
    RunTmRequest xmlThisRequestDoc, xmlTriggerTaskDoc
CreateCaseStageTriggerExit:
    
    Set xmlElem = Nothing
    'TK 14/09/2004 BBG1415 Set xmlNode = Nothing
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
    ' validate request ================================================================
    Set xmlCaseTaskInputList = vxmlRequestNode.selectNodes("CASETASK")
    xmlCheckMandatoryNode vxmlRequestNode, "APPLICATION/@APPLICATIONPRIORITY"
    ' create CreateCaseTask request ====================================================
    ' attributes from original casetask request
    Set xmlCreateCaseTaskDoc = xmlCreateDOMObject
    Set xmlCreateCaseTaskRequestNode = _
        xmlCreateCaseTaskDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set operation attribute
    xmlCreateCaseTaskRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateCaseTask"
    ' create GetTaskDetail request =====================================================
    Set xmlGetTaskDetailDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlGetTaskDetailDoc.async = False
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
    Set xmlTaskDetailDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlTaskDetailDoc.async = False
    'BM0340 MDC 17/04/2003
    Dim strUserId As String
    Dim strUnitId As String
    Call GetApplicationOwners(vxmlRequestNode, strUserId, strUnitId)
    'BM0340 MDC 17/04/2003 - End
    ' create CaseTask elements
    For Each xmlNode In xmlCaseTaskInputList
        
        xmlCheckMandatoryAttribute xmlNode, "SOURCEAPPLICATION"
        xmlCheckMandatoryAttribute xmlNode, "CASEID"
        xmlCheckMandatoryAttribute xmlNode, "ACTIVITYID"
        xmlCheckMandatoryAttribute xmlNode, "ACTIVITYINSTANCE"
        xmlCheckMandatoryAttribute xmlNode, "STAGEID"
        xmlCheckMandatoryAttribute xmlNode, "TASKID"
        Set xmlCreateCaseTaskNode = _
            xmlCreateCaseTaskRequestNode.appendChild(xmlNode.cloneNode(True))
        xmlCopyAttribute xmlNode, xmlGetTaskNode, "TASKID"
        'BM0553
        xmlCopyAttribute xmlNode, xmlGetTaskNode, "STAGEID"
            
        ' run GetTaskDetail request
        RunTmRequest xmlGetTaskDetailDoc, xmlTaskDetailDoc
        Set xmlTaskDetailNode = xmlTaskDetailDoc.selectSingleNode("RESPONSE/TASK")
        If Not xmlTaskDetailNode Is Nothing Then
            'BM0340 MDC 17/04/2003
            'CORE174 AS 25/07/2005 AddDefaultValuesToCaseTask moved to omTMFunctions.bas; pass gobjContext.
            AddDefaultValuesToCaseTask gobjContext, vxmlRequestNode, xmlCreateCaseTaskNode, xmlTaskDetailNode, strUserId, strUnitId
            'BM0340 MDC 17/04/2003 - End
        End If
        ' AQR SYS1791
        ' add ORIGINATINGSTAGEID
        xmlCopyAttributeValue _
            xmlNode, xmlCreateCaseTaskNode, "STAGEID", "ORIGINATINGSTAGEID"
        'DB BM0248 - Add mandatory flag.
        xmlCopyAttributeValue xmlNode, xmlCreateCaseTaskNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
        'DB End
        
    Next
            
    ' run CreateCaseTask request
    RunTmRequest xmlCreateCaseTaskDoc
CreateAdhocCaseTaskExit:
    
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
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CriticalDataCheck(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CriticalDataCheckExit
    Const cstrFunctionName As String = "CriticalDataCheck"
    
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
'    Dim blnAfterCreditCheckStage As Boolean
'    Dim blnCompletedCreditCheck As Boolean
'    Dim blnOutstandingCreditCheck As Boolean
    Dim blnOutstandingCostModelTask As Boolean
    Dim blnCreditCheckTestRequired As Boolean
    Dim blnCostModelTestRequired As Boolean
    
    Dim objNTxBO As IOmTmNTxBO

    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strCurrentStageID As String
    Dim strTaskId As String 'TK 08/12/2004 E2EM00003094
    Dim xmlCaseTask As IXMLDOMNode
    
    gobjContext.SetComplete
    
    ' validate request
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
    
    ' get current CASESTAGE detail
    Set xmlThisRequestDoc = xmlCreateDOMObject
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
    xmlCopyAttribute xmlContextNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlContextNode, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlContextNode, xmlElem, "STAGEID"
    xmlCopyAttributeValue xmlContextNode, xmlElem, "APPLICATIONNUMBER", "CASEID"
    xmlThisRequestNode.appendChild xmlElem
    Set xmlCurrentStageDoc = xmlCreateDOMObject
    
    GetCurrentStage xmlThisRequestNode, xmlCurrentStageDoc
    
    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
    
    
    strApplicationNumber = xmlGetAttributeText(xmlContextNode, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = xmlGetAttributeText(xmlContextNode, "APPLICATIONFACTFINDNUMBER")
    strCurrentStageID = xmlGetAttributeText(xmlContextNode, "STAGEID")
    
    blnCreditCheckTestRequired = False
    
    If strApplicationNumber <> "" Or strApplicationFactFindNumber <> "" Then
        If TransactSummaryCheckRequired(strApplicationNumber, strApplicationFactFindNumber) Then
            blnCreditCheckTestRequired = True
        End If
    End If
    
    ' cost modelling test required
    If Not blnOutstandingCostModelTask Then
        If CostModelCheckRequired(xmlCurrentStageNode, xmlContextNode) Then
            blnCostModelTestRequired = True
        End If
    End If
    
    If blnCreditCheckTestRequired Or blnCostModelTestRequired Then
        ' get Application before image
        Set xmlApplicationDataDoc = xmlCreateDOMObject
        'TK 14/09/2004 BBG1415 xmlApplicationDataDoc.async = False
        Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
        objNTxBO.GetApplicationData xmlContextNode, xmlApplicationDataDoc, "CDRBTemplate"
        Set objNTxBO = Nothing
        Set xmlCriticalDataRequestDoc = xmlCreateDOMObject
        'TK 14/09/2004 BBG1415 xmlCriticalDataRequestDoc.async = False
        Set xmlCriticalDataRequestNode = _
            xmlCriticalDataRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        Set xmlElem = xmlCriticalDataRequestDoc.createElement("BEFORE")
        Set xmlNode = xmlCriticalDataRequestNode.appendChild(xmlElem)
        xmlNode.appendChild _
            xmlApplicationDataDoc.selectSingleNode("RESPONSE/APPLICATION").cloneNode(True)
    End If
    ' call Data Update Function
    ForwardBORequest vxmlRequestNode, vxmlResponseNode
    If blnCreditCheckTestRequired Or blnCostModelTestRequired Then
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
        ' get Application after image
        xmlApplicationDataDoc.removeChild xmlApplicationDataDoc.firstChild
        Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
        objNTxBO.GetApplicationData xmlContextNode, xmlApplicationDataDoc, "CDRBTemplate"
        Set objNTxBO = Nothing
        Set xmlElem = xmlCriticalDataRequestDoc.createElement("AFTER")
        Set xmlNode = xmlCriticalDataRequestNode.appendChild(xmlElem)
        xmlNode.appendChild _
            xmlApplicationDataDoc.selectSingleNode("RESPONSE/APPLICATION").cloneNode(True)
        Set xmlCriticalDataRulesResponseDoc = xmlCreateDOMObject
        'TK 14/09/2004 BBG1415 xmlCriticalDataRulesResponseDoc.async = False
            
        RunCriticalDataRules xmlCriticalDataRequestNode, xmlCriticalDataRulesResponseDoc
        Set xmlCriticalDataRulesResponseNode = _
            xmlCriticalDataRulesResponseDoc.selectSingleNode("RESPONSE")
        errCheckXMLResponseNode xmlCriticalDataRulesResponseNode, , True
        
        ' any tasks to be created?
        If xmlCriticalDataRulesResponseNode.hasChildNodes Then
            'TK 08/12/2004 E2EM00003094
            Set xmlCaseTask = xmlCriticalDataRulesResponseNode.selectSingleNode("CASETASK")
            If Not xmlCaseTask Is Nothing Then
                strTaskId = xmlGetAttributeText(xmlCaseTask, "TASKID")
            End If
            'If the task is not present or if the task is completed add the task
            If strTaskId <> "" Then
            'SDS 15/12/2004 E2EM00003094__START
                'If Not IsTaskPresent(xmlCurrentStageDoc, strTaskId) Or _
                        IsTaskCompleted(xmlCurrentStageDoc, strTaskId) Then
                'SDS : If the task specified does NOT exist in Incomplete status, then add the task
                'TK 18/01/2005 E2EM00003094 Include NotApplicable task status
                Dim blnIncomplete As Boolean
                Dim blnNotApplicable As Boolean
            
                blnIncomplete = IsTaskPresent(xmlCurrentStageDoc, strTaskId, , , omiga4TASKSTATUSIncomplete)
                blnNotApplicable = IsTaskPresent(xmlCurrentStageDoc, strTaskId, , , omiga4TASKSTATUSNotApplicable)
                
                If Not blnIncomplete Then
                    CreateCriticalDataTasks vxmlRequestNode, _
                                                xmlCurrentStageNode, _
                                                xmlContextNode, _
                                                xmlCriticalDataRulesResponseNode
                Else
                    If Not blnIncomplete And Not blnNotApplicable Then
                        CreateCriticalDataTasks vxmlRequestNode, _
                                                xmlCurrentStageNode, _
                                                xmlContextNode, _
                                                xmlCriticalDataRulesResponseNode
                    End If
                End If
                'TK 18/01/2005 E2EM00003094 End
            End If
            'SDS 15/12/2004 E2EM00003094__END
        End If
    End If
    'TK 14/09/2004 BBG1415
    Set xmlContextNode = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlCurrentStageNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationDataDoc = Nothing
    Set xmlCriticalDataRequestDoc = Nothing
    Set xmlCriticalDataRequestNode = Nothing
    Set xmlCriticalDataRulesResponseDoc = Nothing
    Set xmlCriticalDataRulesResponseNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set objNTxBO = Nothing
    Set xmlCaseTask = Nothing
CriticalDataCheckExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub


Private Function IsTaskCompleted(ByVal vxmlApplicationDataNode As IXMLDOMNode, _
                                ByVal strTaskId As String, _
                                Optional ByVal strTaskName As String) As Boolean
    
    ' Optional argument strTaskName is for the cases where the task has been repeated for
    ' more than one applicant so we want to see if the correct task has been completed
    
    IsTaskCompleted = False
    
    Dim xmlTaskList As IXMLDOMNodeList
    Set xmlTaskList = vxmlApplicationDataNode.selectNodes("//CASETASK") ' All current tasks
    
    Dim xmlTaskArchiveList As IXMLDOMNodeList
    Set xmlTaskArchiveList = vxmlApplicationDataNode.selectNodes("CASEACTIVITY/CASETASKARCHIVE") ' All completed tasks
    
    Dim blnCompleted As Boolean
    blnCompleted = False
    
    Dim xmlTaskNode As IXMLDOMNode
    For Each xmlTaskNode In xmlTaskList
        ' If its our task
        If strTaskId = xmlGetAttributeText(xmlTaskNode, "TASKID") Then
            If Len(strTaskName) = 0 Then
                If xmlGetAttributeAsInteger(xmlTaskNode, "TASKSTATUS") = 40 Then  ' 40 = Complete
                    blnCompleted = True
                End If
            ElseIf strTaskName = xmlGetAttributeText(xmlTaskNode, "CASETASKNAME") Then
                If xmlGetAttributeAsInteger(xmlTaskNode, "TASKSTATUS") = 40 Then  ' 40 = Complete
                    blnCompleted = True
                End If
            End If
        End If
    Next
    
    Dim xmlTaskArchive As IXMLDOMNode
    For Each xmlTaskArchive In xmlTaskArchiveList
        ' If its our task
        If strTaskId = xmlGetAttributeText(xmlTaskArchive, "TASKID") Then
            If Len(strTaskName) = 0 Then
                If xmlGetAttributeAsInteger(xmlTaskArchive, "TASKSTATUS") = 40 Then  ' 40 = Complete
                    blnCompleted = True
                End If
            ElseIf strTaskName = xmlGetAttributeText(xmlTaskArchive, "CASETASKNAME") Then
                If xmlGetAttributeAsInteger(xmlTaskArchive, "TASKSTATUS") = 40 Then  ' 40 = Complete
                    blnCompleted = True
                End If
            End If
        End If
    Next
    
    If blnCompleted Then IsTaskCompleted = True
    
    Set xmlTaskList = Nothing
    Set xmlTaskArchiveList = Nothing
    Set xmlTaskNode = Nothing
    Set xmlTaskArchive = Nothing
    
End Function

Private Function IsTaskPresent(ByVal vxmlApplicationCurrentStageDataNode As IXMLDOMNode, _
                                ByVal strTaskId As String, _
                                Optional ByVal strTaskName As String, _
                                Optional ByVal strStage As String, _
                                Optional ByVal intStatus As Integer) As Boolean
    
    ' Optional argument strTaskName is for the cases where the task has been repeated for
    ' more than one applicant so we want to see if the correct task has been completed
    ' Note not implemented in this routine but see isTaskComplete
    
    ' Optional argument strStage is for when it is important that a task is present at a
    ' particular stage
    
    ' OptionalArgument intStatus for when we are concerned about the state of the task.
    
    IsTaskPresent = False
    
    Dim xmlTask As IXMLDOMNode
    
    ' No point going further if we have no base data
    If vxmlApplicationCurrentStageDataNode Is Nothing Then
        Return
    End If
    
    ' Build up the XML search string according to the input parameters
    ' This will search the tasks in the current stage
    ' BBG1419 - Allow for packaged tasks in the search string
    Dim strXSL As String
    strXSL = ".//CASESTAGE/CASETASK[@TASKID='" & strTaskId & "'" & _
             " or @TASKID='" & GetPackagedTask(strTaskId) & "']"
    If Len(strStage) <> 0 Then
        strXSL = strXSL & "[@STAGEID='" & strStage & "']"
    End If
    If intStatus <> 0 Then
        strXSL = strXSL & "[@TASKSTATUS='" & intStatus & "']"
    End If
    
    Set xmlTask = vxmlApplicationCurrentStageDataNode.selectSingleNode(strXSL)
    
    If Not xmlTask Is Nothing Then
        IsTaskPresent = True
        GoTo isTaskPresentExit
    End If
    
    ' Not in the current stage - try in the archive
    strXSL = "CASEACTIVITY/CASETASKARCHIVE[@TASKID='" & strTaskId & "'" & _
             " or @TASKID='" & GetPackagedTask(strTaskId) & "']"
    If Len(strStage) <> 0 Then
        strXSL = strXSL & "[@STAGEID='" & strStage & "']"
    End If
    If intStatus <> 0 Then
        strXSL = strXSL & "[@TASKSTATUS='" & intStatus & "']"
    End If

    Set xmlTask = vxmlApplicationCurrentStageDataNode.selectSingleNode(strXSL)

    If Not xmlTask Is Nothing Then
        IsTaskPresent = True
    End If
    
isTaskPresentExit:
    Set xmlTask = Nothing
    
End Function

Private Function GetPackagedTask(ByVal strTaskName As String) As String
'
' Function which returns the related Packaged task name. Used in isTaskPresent etc
' If not one of the known ones where we had field length restrictions put an "_P" on
' the end of the task name
'
    Dim strResp As String
    Select Case strTaskName
        Case "Address_Undisclosed"
            strResp = "Address_Undisclose_P"
        Case "Bank_Stats_Business"
            strResp = "Bank_Stats_Busines_P"
        Case "Credit_Agreemt_Undis"
            strResp = "Credit_Agree_Undis_P"
        Case "Declaration_SignDate"
            strResp = "Declaration_Signed_P"
        Case "Decree_Absolute_Req"
            strResp = "Decree_Absolut_Req_P"
        Case "Emp_Contracts(2)_Req"
            strResp = "Emp_Contracts(2)_P"
        Case "Future_Resi_Address"
            strResp = "Future_Resi_Addres_P"
        Case "Incentive_MaxLoan_OK"
            strResp = "Incentive_MaxLn_OK_P"
        Case "Income_Proof_>65_Req"
            strResp = "Income_Proof_>65_P"
        Case "Lender_Fee_Rtrn_Chq"
            strResp = "Lender_Fee_Rtn_Chq_P"
        Case "Lenders_Ref_Request"
            strResp = "Lenders_Ref_Req_P"
        Case "Let_Management_Lead"
            strResp = "Let_Management_Led_P"
        Case "Low_Rent_MaxLoan_OK"
            strResp = "Low_Rent_MaxLn_OK_P"
        Case "LowIncome_MaxLoan_OK"
            strResp = "LowIncome_MaxLn_OK_P"
        Case "Mortgage_Stat_12Mths"
            strResp = "Mortgage_Stat_12M_P"
        Case "Mortgage_Stats_Orig"
            strResp = "Mortgage_Stat_Orig_P"
        Case "New_Max_Loan_Letter"
            strResp = "New_Max_Loan_Lettr_P"
        Case "Orig_Debt_Loan_Agmt"
            strResp = "Orig_Debt_Loan_Agt_P"
        Case "Pension_Income_Proof"
            strResp = "Pension_Income_Prf_P"
        Case "Redeem_No_MaxLoan_OK"
            strResp = "Redeem_No_MaxLn_OK_P"
        Case "Refer_Compare_Letter"
            strResp = "Refer_Compare_Letr_P"
        Case "Rent_Compare_Review"
            strResp = "Rent_Compar_Review_P"
        Case "Sols_Details_Missing"
            strResp = "Sols_Details_Miss_P"
        Case "Sols_Panel_Instruct"
            strResp = "Sols_Panel_Instrut_P"
        Case "Structure_Report_Req"
            strResp = "Structure_Report_P"
        Case "Tax_Liabilities_O/S"
            strResp = "Tax_Liabilities_OS_P"
        Case "Val_Req_Missing_Info"
            strResp = "Val_Req_Miss_Info_P"
        Case "Val_Structure_Refer"
            strResp = "Val_Structur_Refer_P"
        Case "Verify_Drive_Licence"
            strResp = "Verify_Drive_Licen_P"
        Case Else
            strResp = strTaskName & "_P"
    End Select
    GetPackagedTask = strResp
End Function
Private Sub CheckTasksOnCurrentStage( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByRef vblnHasOutstandingCreditCheckTask As Boolean, _
    ByRef vblnHasCompletedCreditCheckTask As Boolean, _
    ByRef vblnOutstandingCostModelTask As Boolean)
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim strXSL As String, _
        strCreditCheckTask As String, _
        strReCreditCheckTask As String, _
        strReCostModelTask As String
    Dim blnOutstandingCreditCheckTask As Boolean, _
        blnCompletedCreditCheckTask As Boolean, _
        blnOutstandingReDoCreditCheckTask As Boolean, _
        blnCompletedReDoCreditCheckTask As Boolean
    Dim intTaskInstance As Integer, _
        intThisInstance As Integer
    ' needs global param?
    strCreditCheckTask = "Credit_Score"
    strReCreditCheckTask = GetGlobalParamString("TMReprocessCreditCheck")
    strReCostModelTask = GetGlobalParamString("TMRemodelMortgage")
    strXSL = "CASETASK[@TASKID='" & strCreditCheckTask & "']"
    Set xmlNodeList = vxmlCurrentStageNode.selectNodes(strXSL)
    intTaskInstance = -1
    For Each xmlNode In xmlNodeList
        intThisInstance = xmlGetAttributeAsInteger(xmlNode, "TASKINSTANCE")
        If intThisInstance > intTaskInstance Then
            intTaskInstance = intThisInstance
            If xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete Then
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
                If xmlGetAttributeAsInteger(xmlNode, "TASKSTATUS") = omiga4TASKSTATUSComplete Then
                    blnOutstandingReDoCreditCheckTask = False
                    blnCompletedReDoCreditCheckTask = True
                Else
                    blnOutstandingReDoCreditCheckTask = True
                    blnCompletedReDoCreditCheckTask = False
                End If
            End If
        Next
    End If
    If blnOutstandingCreditCheckTask Or blnOutstandingReDoCreditCheckTask Then
        vblnHasOutstandingCreditCheckTask = True
    Else
        If blnCompletedCreditCheckTask Or blnCompletedReDoCreditCheckTask Then
            vblnHasCompletedCreditCheckTask = True
        End If
    End If
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
    Set xmlNodeList = Nothing
    Set xmlNode = Nothing
End Sub
Private Function IsAfterCreditCheckStage(ByVal vstrThisStageId As String) As Boolean
    Const cstrFunctionName As String = "IsAfterCreditCheckStage"
    Dim strCreditCheckStage As String
    strCreditCheckStage = GetMandatoryGlobalParamString("TMCreditCheckStage")
    If Not IsNumeric(vstrThisStageId) Then
        errThrowError _
            cstrFunctionName, _
            oeInvalidParameter, _
            "this Stage Id not numeric: " & vstrThisStageId
    End If
    If Not IsNumeric(strCreditCheckStage) Then
        errThrowError _
            cstrFunctionName, _
            oeInvalidParameter, _
            "CreditCheck Stage Id not numeric: " & strCreditCheckStage
    End If
    IsAfterCreditCheckStage = CInt(vstrThisStageId) > CInt(strCreditCheckStage)
End Function
Private Function CostModelCheckRequired( _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlApplicationNode As IXMLDOMNode) _
    As Boolean
    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
    Dim xmlAttrib As IXMLDOMAttribute
    Dim objNTxBO As IOmTmNTxBO  'SYS3270
    ' add _EXTRACTTYPE_ 'NODE' to application node
    ' this gets copied to omRB request by GetApplicationData,
    ' this causes omRB to return APPLICATION node only (no child nodes)
    Set xmlAttrib = vxmlApplicationNode.ownerDocument.createAttribute("_EXTRACTTYPE_")
    xmlAttrib.Text = "NODE"
    vxmlApplicationNode.Attributes.setNamedItem xmlAttrib
    Set xmlApplicationDataDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlApplicationDataDoc.async = False
    'SYS3270 {
    'GetApplicationData vxmlApplicationNode, xmlApplicationDataDoc
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.GetApplicationData vxmlApplicationNode, xmlApplicationDataDoc
    Set objNTxBO = Nothing
    'SYS3270 }
    vxmlApplicationNode.Attributes.removeNamedItem "_EXTRACTTYPE_"
    CostModelCheckRequired = _
        Len( _
            xmlGetNodeText(xmlApplicationDataDoc, "RESPONSE/APPLICATION/@ACCEPTEDQUOTENUMBER")) _
            > 0
    Set xmlApplicationDataDoc = Nothing
    Set xmlAttrib = Nothing
End Function
'TK 15/09/2004 BBG1207 - check whether any record exists in GetTransactSummary view
Private Function TransactSummaryCheckRequired(ByVal strApplicationNumber As String, _
                            ByVal strApplicationFactFindNumber As String) As Boolean
    Dim objADOAssist As Object
    Dim objSQLAssist As Object
    Dim strCondition  As String
    Dim blnRecordFound As Boolean

    Set objADOAssist = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    Set objSQLAssist = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".SQLAssist")
    
    'Build the condition and Get value from table 'MORGAGESUBQUOTE'
    strCondition = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber) _
                   & " And APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
                   
    blnRecordFound = objADOAssist.CheckRecordExists("GETTRANSACTSUMMARY", strCondition)
    If Not blnRecordFound Then
        TransactSummaryCheckRequired = False
    Else
        TransactSummaryCheckRequired = True
    End If
    
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
End Function

'TK 15/09/2004 BBG1207 End
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
    vxmlCriticalDataRulesResponseDoc.setProperty "NewParser", True
    vxmlCriticalDataRulesResponseDoc.validateOnParse = False
    vxmlCriticalDataRulesResponseDoc.loadXML _
        objCDRulesBO.RunRules(vxmlRequestNode.xml)
    Set objCDRulesBO = Nothing
        
    errCheckXMLResponseNode _
        vxmlCriticalDataRulesResponseDoc.selectSingleNode("RESPONSE"), , True
End Sub
Private Sub CreateCriticalDataTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCurrentStageNode As IXMLDOMNode, _
    ByVal vxmlCriticalDataContextNode As IXMLDOMNode, _
    ByVal vxmlCriticalDataRulesResponseNode As IXMLDOMNode)
    On Error GoTo CreateCriticalDataTasksExit
    Const cstrFunctionName As String = "CreateCriticalDataTasks"
    ' create CreateAdhocCaseTask request
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlThisRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisRequestDoc.async = False
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "CreateAdhocCaseTask"
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    xmlCopyAttribute vxmlCriticalDataContextNode, xmlElem, "APPLICATIONPRIORITY"
    xmlThisRequestNode.appendChild xmlElem
    For Each xmlNode In vxmlCriticalDataRulesResponseNode.childNodes
        Set xmlThisCaseTaskNode = xmlThisRequestNode.appendChild(xmlNode.cloneNode(False))
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "SOURCEAPPLICATION"
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "CASEID"
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "CASESTAGESEQUENCENO"
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "ACTIVITYID"
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "ACTIVITYINSTANCE"
        xmlCopyAttribute vxmlCurrentStageNode, xmlThisCaseTaskNode, "STAGEID"
    Next
    CreateAdhocCaseTask xmlThisRequestNode
CreateCriticalDataTasksExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisRequestDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreateCurrEmployersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlThisRequestDoc = xmlCreateDOMObject
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
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidateCurrEmployersRef
    ' NOT WRITTEN YET
    ' SA SYS2256 Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreatePrevEmployersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    Set xmlThisResponseDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlThisRequestDoc = xmlCreateDOMObject
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
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidatePrevEmployersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreateAccountantsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    'TK 14/09/2004 BBG1415 Set xmlThisRequestDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidateAccountantsRef
    ' NOT YET WRITTEN
    ' SA SYS2256 Code now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreateCurrLendersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidateCurrLendersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreatePrevLendersRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidatePrevLendersRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreateCurrLandlordsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidateCurrLandlordsRef
    ' NOT YET WRITTEN
    'SA SYS2256 Code now availble
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreatePrevLandlordsRef
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlAppStatusNode As IXMLDOMNode       'SA SYS2256
    ' Call referencesBO.ValidatePrevLandlordsRef
    ' NOT YET WRITTEN
    ' SA Methods now available
    Set objBO = gobjContext.CreateInstance("omAppProc.omAppProcBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.omAppProcRequest(vxmlRequestNode.xml)
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
    Set xmlAppStatusNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
ValidatePrevLandlordsRefErr:
    ' PSC 06/12/01 SYS3290
    gobjContext.SetAbort
    GoTo ValidatePrevLandlordsRefExit
End Sub
Private Sub PrintDocumentForTask(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo PrintDocumentForTaskErr
    Const cstrFunctionName As String = "PrintDocumentForTask"
    
    Dim objBO As Object
    
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlRequestPrintAttributesNode As IXMLDOMNode
    Dim xmlRequestApplicationNode As IXMLDOMNode
    Dim objPrintBO As Object
    Dim strValue As String
    Dim strResponseFromPrint As String
    Dim objAppBO As Object
    Dim xmlAppRequest As IXMLDOMNode
    Dim xmlApplication As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlControlData As IXMLDOMElement
    Dim xmlAppResponse As FreeThreadedDOMDocument40
    Dim strTaskId As String
    Dim strResponse As String
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim intSubmissionMethod As Integer
    Dim blnPackagedIndicator As Boolean
    Dim xmlKFIResponseDoc       As FreeThreadedDOMDocument40
    Dim xmlKFIDoc               As DOMDocument
    Dim xmlKFINode              As IXMLDOMNode
    Dim xmlKFICustNumber        As IXMLDOMNode
    Dim xmlKFINodeList          As IXMLDOMNodeList
    Dim xmlKFIRequest           As IXMLDOMElement
    Dim xmlKFICustAddList       As IXMLDOMElement
    Dim xmlKFICustAdd           As IXMLDOMElement
    Dim xmlKFIAppNumber         As IXMLDOMElement
    Dim xmlKFIAppFactNumber     As IXMLDOMElement
    Dim xmlKFICNumber           As IXMLDOMElement
    Dim xmlKFICustVersion       As IXMLDOMElement
    
    Dim xmlPrintResponseDoc  As FreeThreadedDOMDocument40
    Set xmlPrintResponseDoc = New FreeThreadedDOMDocument40
    
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    
    Set xmlRequestPrintAttributesNode = xmlGetMandatoryNode(vxmlRequestNode, "PRINTATTRIBUTES")
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    Set xmlRequestApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "COPIES"
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "PRINTER"
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "DPSDOCUMENTID"
    xmlCheckMandatoryAttribute xmlRequestPrintAttributesNode, "DESTINATIONTYPE"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "TASKID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "STAGEID"
    xmlCheckMandatoryAttribute xmlRequestCaseTaskNode, "OUTPUTDOCUMENT"
    
    'Call the print engine
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    
    ' copy request node
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    
    'set operation attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "PrintDocument"
    
    'add new indicator to ensure archiving happens for these prints.
    xmlSetAttributeValue xmlThisRequestNode, "PRINTINDICATOR", "1"
    
    ' create PRINTDATA element
    Set xmlElem = xmlThisRequestDoc.createElement("PRINTDATA")
    
    ' copy required search attributes
    strAppNo = xmlGetAttributeText(xmlRequestCaseTaskNode, "CASEID")
    xmlElem.setAttribute "APPLICATIONNUMBER", strAppNo
    
    strAppFFNo = xmlGetAttributeText(xmlRequestApplicationNode, "APPLICATIONFACTFINDNUMBER")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNo
    
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CUSTOMERIDENTIFIER")
    xmlElem.setAttribute "CUSTOMERNUMBER", strValue
    
    strValue = xmlGetAttributeText(xmlRequestApplicationNode, "CUSTOMERVERSIONNUMBER")
    xmlElem.setAttribute "CUSTOMERVERSIONNUMBER", strValue
    
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "CONTEXT")
    xmlElem.setAttribute "RECIPIENTKEY", strValue
    
    strTaskId = xmlGetAttributeText(xmlRequestCaseTaskNode, "TASKID")
    xmlElem.setAttribute "REFTYPE", strTaskId
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "METHODNAME")
    xmlElem.setAttribute "METHODNAME", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "RECIPIENTTYPE")
    xmlElem.setAttribute "RECIPIENTTYPE", strValue
    
    xmlThisRequestNode.appendChild xmlElem
    
    ' Create TEMPLATEDATA node if required
    strValue = xmlGetAttributeText(xmlRequestApplicationNode, "PRINTKFI")
    If strValue = "1" Then
        Set xmlElem = xmlThisRequestDoc.createElement("TEMPLATEDATA")
        xmlCopyAttribute xmlRequestApplicationNode, xmlElem, "QUOTATIONNUMBER"
        xmlCopyAttribute xmlRequestApplicationNode, xmlElem, "MORTGAGESUBQUOTENUMBER"
        xmlThisRequestNode.appendChild xmlElem
    End If
    
    ' create CONTROLDATA element
    Set xmlControlData = xmlThisRequestDoc.createElement("CONTROLDATA")
    
    ' clone search attributes
    strValue = xmlGetAttributeText(xmlRequestCaseTaskNode, "OUTPUTDOCUMENT")
    xmlControlData.setAttribute "DOCUMENTID", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "PRINTER")
    xmlControlData.setAttribute "PRINTER", strValue
  
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "COPIES")
    xmlControlData.setAttribute "COPIES", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DELIVERYTYPE")
    xmlControlData.setAttribute "DELIVERYTYPE", strValue
    
    Dim strDestinationType As String
    strDestinationType = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DESTINATIONTYPE")
    xmlControlData.setAttribute "DESTINATIONTYPE", strDestinationType
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "DPSDOCUMENTID")
    xmlControlData.setAttribute "DPSDOCUMENTID", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "HOSTTEMPLATENAME")
    xmlControlData.setAttribute "HOSTTEMPLATENAME", strValue
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "HOSTTEMPLATEDESCRIPTION")
    xmlControlData.setAttribute "HOSTTEMPLATEDESCRIPTION", strValue
    
    xmlControlData.setAttribute "POSTTOWEB", "0"
    
    strValue = xmlGetAttributeText(xmlRequestPrintAttributesNode, "COMPRESSIONMETHOD")
    xmlControlData.setAttribute "COMPRESSIONMETHOD", strValue
    
    xmlCopyAttribute xmlRequestPrintAttributesNode, xmlControlData, "WEBDOCUMENTTYPE"
   
    ' Determine if the document needs to be sent to the web
    If StrComp(strTaskId, "Offer_Broker_RMC", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "Offer_Broker_RLMC", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "Offer_Broker_NonReg", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "Offer_Pack_RMC", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "Offer_Pack_RLMC", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "Offer_Pack_NonReg", vbTextCompare) = 0 Or _
        StrComp(strTaskId, "OVERTURN_KFI", vbTextCompare) = 0 Then
            
        Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
        
        Set xmlAppRequest = xmlGetRequestNode(vxmlRequestNode)
        
        Set xmlApplication = xmlThisRequestDoc.createElement("APPLICATIONFACTFIND")
        xmlAppRequest.appendChild xmlApplication
        
        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONNUMBER")
        xmlNode.Text = strAppNo
        xmlApplication.appendChild xmlNode
        
        Set xmlNode = xmlThisRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")
        xmlNode.Text = strAppFFNo
        xmlApplication.appendChild xmlNode
        
        strResponse = objAppBO.GetApplicationFactFindData(xmlAppRequest.xml)
        
        Set xmlAppResponse = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlAppResponse.documentElement, , True
        
        intSubmissionMethod = xmlGetNodeAsInteger(xmlAppResponse, "RESPONSE/APPLICATIONFACTFIND/METHODOFSUBMISSION")
        blnPackagedIndicator = xmlGetNodeAsBoolean(xmlAppResponse, "RESPONSE/APPLICATIONFACTFIND/APPLICATIONPACKAGEINDICATOR")
        
        If IsValidationType("MethodOfSubmission", intSubmissionMethod, "ONLINE") Then
            
            If blnPackagedIndicator = True Then
                If StrComp(strTaskId, "Offer_Pack_RMC", vbTextCompare) = 0 Or _
                   StrComp(strTaskId, "Offer_Pack_RLMC", vbTextCompare) = 0 Or _
                   StrComp(strTaskId, "Offer_Pack_NonReg", vbTextCompare) = 0 Then
                    xmlControlData.setAttribute "POSTTOWEB", "1"
                End If
            End If
            
            If blnPackagedIndicator = False Then
                If StrComp(strTaskId, "Offer_Broker_RMC", vbTextCompare) = 0 Or _
                   StrComp(strTaskId, "Offer_Broker_RLMC", vbTextCompare) = 0 Or _
                   StrComp(strTaskId, "Offer_Broker_NonReg", vbTextCompare) = 0 Then
                    xmlControlData.setAttribute "POSTTOWEB", "1"
                End If
            End If
            
            If StrComp(strTaskId, "OVERTURN_KFI", vbTextCompare) = 0 Then
                xmlControlData.setAttribute "POSTTOWEB", "1"
            End If
            
        End If
        
    End If
    
    xmlThisRequestNode.appendChild xmlControlData
    
    'Call omPrintBO.PrintDocument
    On Error Resume Next
    Set objPrintBO = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    On Error GoTo PrintDocumentForTaskErr
    
    If Not objPrintBO Is Nothing Then
        
        strResponseFromPrint = objPrintBO.OmRequest(xmlThisRequestDoc.xml)
        
        'CheckResponse for errors
        Set xmlPrintResponseDoc = xmlLoad(strResponseFromPrint, cstrFunctionName)
        errCheckXMLResponseNode xmlPrintResponseDoc.documentElement, , True
        If strDestinationType = "W" Then
            vxmlResponseNode.appendChild xmlPrintResponseDoc.selectSingleNode(".//DOCUMENTCONTENTS")
        End If
    End If
    
'    'Call omTMBO.UpdateCaseTask
'    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
'
'    ' copy request node
'    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
'
'    ' set operation attribute
'    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
'
'    ' create CASETASK element
'    Set xmlElem = xmlThisRequestDoc.createElement("CASETASK")
'
'    ' clone search attributes
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASEID"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "STAGEID"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "CASESTAGESEQUENCENO"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYID"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "ACTIVITYINSTANCE"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKID"
'    xmlCopyAttribute xmlRequestCaseTaskNode, xmlElem, "TASKINSTANCE"
'
'    If xmlGetAttributeText(xmlRequestCaseTaskNode, "INPUTPROCESS") = "" Then
'        xmlElem.setAttribute "TASKSTATUS", "40" 'Complete
'    Else
'        xmlElem.setAttribute "TASKSTATUS", "20" 'Pending
'    End If
'
'    xmlThisRequestNode.appendChild xmlElem
'
'
'    xmlThisResponseDoc.async = False
'
'    ' run Tm UpdateCaseTask request
'    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
    
    ' parse response
    'errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    gobjContext.SetComplete
    
PrintDocumentForTaskExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlRequestPrintAttributesNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlRequestApplicationNode = Nothing
    Set objPrintBO = Nothing
    Set objAppBO = Nothing
    Set xmlAppRequest = Nothing
    Set xmlApplication = Nothing
    Set xmlNode = Nothing
    Set xmlAppResponse = Nothing
    Set xmlControlData = Nothing
    Set xmlKFIResponseDoc = Nothing
    Set xmlKFIDoc = Nothing
    Set xmlKFINode = Nothing
    Set xmlKFICustNumber = Nothing
    Set xmlKFINodeList = Nothing
    Set xmlKFIRequest = Nothing
    Set xmlKFICustAddList = Nothing
    Set xmlKFICustAdd = Nothing
    Set xmlKFIAppNumber = Nothing
    Set xmlKFIAppFactNumber = Nothing
    Set xmlKFICNumber = Nothing
    Set xmlKFICustVersion = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
PrintDocumentForTaskErr:
    
    gobjContext.SetAbort
    GoTo PrintDocumentForTaskExit
    
End Sub
'BBG1465 Rewrite to use USP_TASKSUMMARYTOTALS Stored Proc instead of TM_CASETASKDETAIL View
Private Sub FindUnitTaskSummary( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo FindUnitTaskSummaryExit
    Const cstrFunctionName As String = "FindUnitTaskSummary"
    
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

    ' run the request
    Set xmlCaseTaskDoc = xmlCreateDOMObject
    
    Debug.Print "FindCaseTaskList request:"
    Debug.Print vxmlRequestNode.ownerDocument.xml
    
    RunTmRequest vxmlRequestNode.ownerDocument, xmlCaseTaskDoc
    errCheckXMLResponseNode xmlCaseTaskDoc.selectSingleNode("RESPONSE"), , True
    
    Debug.Print "FindCaseTaskList response:"
    Debug.Print xmlCaseTaskDoc.xml
    
    'Change UnitTotals and UserAvergages to expected node names
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


FindUnitTaskSummaryExit:

    Set xmlCaseTaskDoc = Nothing
    Set xmlUnitRequestNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlTempNode = Nothing
    Set xmlParentNode = Nothing
    Set xmlUserNode = Nothing
    Set xmlUserNodeList = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub GetUserTotals( _
    ByVal vxmlOmigaUserResponseElem As IXMLDOMElement, _
    ByVal vxmlUserCaseTaskList As IXMLDOMNodeList)
    On Error GoTo GetUserTotalsExit
    Const cstrFunctionName As String = "GetUserTotals"
    Dim xmlNode As IXMLDOMNode
    Dim varCaseIds As Collection
    Dim strThisAppNo As String
    Dim intLoop As Integer, _
        intOSTasks As Integer
    Dim blnCaseHit As Boolean
    Dim dtLimit As Date, _
        dtTest As Date
    dtLimit = Now()
        
    Set varCaseIds = New Collection
    For Each xmlNode In vxmlUserCaseTaskList
        
        strThisAppNo = xmlGetAttributeText(xmlNode, "CASEID")
        blnCaseHit = False
        For intLoop = 1 To varCaseIds.Count
            If varCaseIds.Item(intLoop) = strThisAppNo Then
                blnCaseHit = True
                Exit For
            End If
        Next
        If blnCaseHit = False Then
            varCaseIds.Add strThisAppNo
        End If
            
        If xmlAttributeValueExists(xmlNode, "TASKDUEDATEANDTIME") Then
            dtTest = _
                xmlGetAttributeText(xmlNode, "TASKDUEDATEANDTIME")
            If dtTest < dtLimit Then
                intOSTasks = intOSTasks + 1
            End If
        End If
    Next
    vxmlOmigaUserResponseElem.setAttribute "APPLICATIONS", varCaseIds.Count
    vxmlOmigaUserResponseElem.setAttribute "TASKS", vxmlUserCaseTaskList.length
    vxmlOmigaUserResponseElem.setAttribute "OUTSTANDINGTASKS", intOSTasks
GetUserTotalsExit:
    
    Set xmlNode = Nothing
    Set varCaseIds = Nothing
    errCheckError cstrFunctionName
End Sub
Private Sub GetGroupTotals(ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo GetGroupTotalsExit
    Const cstrFunctionName As String = "GetGroupTotals"
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim intApps As Integer, _
        intTasks As Integer, _
        intOSTasks As Integer
    Set xmlNodeList = vxmlResponseNode.selectNodes("OMIGAUSER")
    For Each xmlNode In xmlNodeList
        intApps = intApps + xmlGetAttributeAsInteger(xmlNode, "APPLICATIONS")
        intTasks = intTasks + xmlGetAttributeAsInteger(xmlNode, "TASKS")
        intOSTasks = intOSTasks + xmlGetAttributeAsInteger(xmlNode, "OUTSTANDINGTASKS")
    Next
    Set xmlElem = vxmlResponseNode.ownerDocument.createElement("UNITTOTALS")
    xmlElem.setAttribute "APPLICATIONS", intApps
    xmlElem.setAttribute "TASKS", intTasks
    xmlElem.setAttribute "OUTSTANDINGTASKS", intOSTasks
    vxmlResponseNode.appendChild xmlElem
    If xmlNodeList.length > 0 Then
        Set xmlElem = vxmlResponseNode.ownerDocument.createElement("USERAVERAGES")
        xmlElem.setAttribute "APPLICATIONS", Round(intApps / xmlNodeList.length)
        xmlElem.setAttribute "TASKS", Round(intTasks / xmlNodeList.length)
        xmlElem.setAttribute "OUTSTANDINGTASKS", Round(intOSTasks / xmlNodeList.length)
        vxmlResponseNode.appendChild xmlElem
    End If
GetGroupTotalsExit:
    
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    errCheckError cstrFunctionName
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

    Set xmlRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlRequestDoc.async = False
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
    Set xmlRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlRequestDoc.async = False
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
    'TK 14/09/2004 BBG1415 Dim xmlNode As IXMLDOMNode
    'TK 14/09/2004 BBG1415 Dim xmlAttrib As IXMLDOMAttribute
    ' create GetNextStageTaskList request
    Set xmlRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlRequestDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlNode = Nothing
    'TK 14/09/2004 BBG1415 Set xmlAttrib = Nothing
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
    'TK 14/09/2004 BBG1415 Dim xmlNode As IXMLDOMNode
    'TK 14/09/2004 BBG1415 Dim xmlAttrib As IXMLDOMAttribute
    ' create GetNextStageTaskList request
    Set xmlRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlRequestDoc.async = False
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
    'TK 14/09/2004 BBG1415 Set xmlNode = Nothing
    'TK 14/09/2004 BBG1415 Set xmlAttrib = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlAppPriorityElem = Nothing
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
    'TK 14/09/2004 BBG1415 Dim xmlAttrib As IXMLDOMAttribute
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCaseActivityNode As IXMLDOMNode
    Dim xmlCaseStageNode As IXMLDOMNode
'    Dim intCaseStageSeq As Integer
    'AQR SYS2266
    Dim objNTxBO As IOmTmNTxBO
       
    ' create CreateCaseStage request ==================================================
    Set xmlRequestDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlRequestDoc.async = False
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
    Set xmlResponseDoc = xmlCreateDOMObject
    'AQR SYS2266 DRC
    ' The New Transaction Class is used to ensure that the stage creation is committed
    ' before any automatic tasks are processed
    Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
    objNTxBO.RunTmRequest xmlRequestDoc, xmlResponseDoc
    Set objNTxBO = Nothing
CreateActivityRequestExit:
    
    Set xmlRequestNode = Nothing
    'TK 14/09/2004 BBG1415 Set xmlAttrib = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlCaseActivityNode = Nothing
    Set xmlCaseStageNode = Nothing
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
                        "  and  @CUSTOMERIDENTIFIER='" & _
                        xmlGetAttributeText(xmlSrceNode, "CUSTOMERIDENTIFIER") & "'"
                End If
                                    
                If xmlAttributeValueExists(xmlSrceNode, "CONTEXT") Then
                    
                    strPattern = strPattern & _
                        "  and  @CONTEXT='" & _
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
                        
                        'PSC 09/12/2004 E2EM00003099 - Start
                        xmlCopyAttribute xmlSrceNode, xmlNode, "TASKSTATUS"
                        xmlCopyAttribute xmlSrceNode, xmlNode, "CONTEXT"
                        xmlCopyAttribute xmlSrceNode, xmlNode, "CUSTOMERIDENTIFIER"
                        'PSC 09/12/2004 E2EM00003099 - End
                        
                        
                        
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
'    Dim strPattern As String
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
'    Dim intStatus As Integer
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
    'TK 14/09/2004 BBG1415 Dim xmlAttrib As IXMLDOMAttribute
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
    'TK 14/09/2004 BBG1415 Set xmlAttrib = Nothing
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
    'BM0340 MDC 17/04/2003
    Dim strUserId As String
    Dim strUnitId As String
    Call GetApplicationOwners(vxmlRequestNode, strUserId, strUnitId)
    'BM0340 MDC 17/04/2003 - End
    For Each xmlCaseTaskNode In vxmlNextCaseStageNode.childNodes
        strPattern = _
            "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlCaseTaskNode, "TASKID") & "']"
        Set xmlTaskNode = vxmlNextStageNode.selectSingleNode(strPattern)
        'BM0340 MDC 17/04/2003
        'CORE174 AS 25/07/2005 AddDefaultValuesToCaseTask moved to omTMFunctions.bas; pass gobjContext.
        AddDefaultValuesToCaseTask gobjContext, vxmlRequestNode, xmlCaseTaskNode, xmlTaskNode, strUserId, strUnitId
        'BM0340 MDC 17/04/2003 - End
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

    Set xmlRequestDoc = xmlCreateDOMObject
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
    Set objApplicationBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    Debug.Print xmlRequestDoc.xml
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.validateOnParse = False
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
'    Set xmlThisRequestNode = Nothing
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
    
    errCheckError cstrFunctionName
End Sub
Private Sub Terminate( _
    ByVal vxmlDoc As FreeThreadedDOMDocument40, _
    ByVal vstrFunctionName As String)
    'TK 14/09/2004 BBG1415
'    vxmlDoc.Save ("c:\temp\OmTmBO.xml")
'    Shell _
'        "c:\program files\plus!\microsoft internet\iexplore.exe c:\temp\OmTmBO.xml", _
'        vbMaximizedFocus
'    gobjContext.SetAbort
'    Err.Raise _
'        oeUnspecifiedError, _
'        TypeName(Me) & "." & vstrFunctionName, _
'        "OmTmBO Terminate Request"
End Sub
Private Sub ImportAccountsIntoApplication( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo ImportAccountsIntoApplicationExit
    Const cstrFunctionName As String = "ImportAccountsIntoApplication"
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMElement
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    'TK 14/09/2004 BBG1415 Dim xmlThisResponseNode As IXMLDOMNode
    'TK 14/09/2004 BBG1415 Dim xmlApplicationRequestNode As IXMLDOMNode
    'TK 14/09/2004 BBG1415 Dim xmlCustomerRequestNode As IXMLDOMNode
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
    
    
    'TK 14/09/2004 BBG1415 Set xmlApplicationRequestNode = Nothing
    'TK 14/09/2004 BBG1415 Set xmlCustomerRequestNode = Nothing
    Set xmlCaseActivityRequestNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlImportAccountsResponseNode = Nothing
    Set xmlElem = Nothing
    Set xmlSrceNode = Nothing
    Set xmlNode = Nothing
    Set objBO = Nothing
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
    'TK 14/09/2004 BBG1415 Dim xmlNode As IXMLDOMNode
        
    Dim lngSequenceNumber As Long
    Dim lngTMOfferStage As Long
    Dim strTMIssueOfferTaskName As String
    Dim strTMReIssueOfferTaskName As String
    Dim strPattern As String
'    Dim strCaseID As String
'    Dim strSourceApp As String
'    Dim strActivityId As String
    Dim lngTaskStatus As Long
            
    Dim strFreezeDataIndicator As String
    Dim xmlNodeList As IXMLDOMNodeList
    Dim iIndex As Integer
    'TK 11/11/2004 E2EM00002757 Start
    Dim strCaseStageId As String
    Dim xmlAttrib As IXMLDOMAttribute
    'TK 11/11/2004 E2EM00002757 End
    'TK 15/11/2004 E2EM00002757 Start
    Dim xmlStageResponseNode As IXMLDOMNode
    Dim xmlCaseStageArchiveResponseNode As IXMLDOMNode
    Dim xmlCaseStageArchiveResponseNodeList As IXMLDOMNodeList
    'TK 15/11/2004 E2EM00002757 End

    Set xmlCaseStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseStageRequestNode, "ACTIVITYID"
    Set xmlThisResponseDoc = xmlCreateDOMObject
        
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
    'TK 11/11/2004 E2EM00002757 Start
    strCaseStageId = xmlGetAttributeText(xmlCaseStageResponseNode, "STAGEID")
    ' get the global parameter
    lngTMOfferStage = GetGlobalParamAmount("TMOfferStage")
        
    'Get stage id from Stage where STAGESEQUENCENO =lngTMOfferStage
    Set xmlCaseActivityNode = vxmlRequestNode.ownerDocument.createElement("STAGE")
    vxmlRequestNode.appendChild xmlCaseActivityNode
    ' PSC 23/11/2004 E2EM00002747 - Start
    Set xmlAttrib = xmlCaseActivityNode.ownerDocument.createAttribute("STAGEID")
    xmlAttrib.Text = strCaseStageId
    ' PSC 23/11/2004 E2EM00002747 - End
    xmlCaseActivityNode.Attributes.setNamedItem xmlAttrib
    xmlCopyMandatoryAttribute xmlCaseStageRequestNode, xmlCaseActivityNode, "ACTIVITYID"
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "GetStageList"
    ' run Tm GetCaseTasks request and get the stage information for the current stage
    RunTmRequest vxmlRequestNode.ownerDocument, xmlThisResponseDoc
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    Set xmlStageResponseNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/STAGE")
    ' PSC 23/11/2004 E2EM00002747
    lngSequenceNumber = xmlGetAttributeAsLong(xmlStageResponseNode, "STAGESEQUENCENO")
    'TK 11/11/2004 E2EM00002757 End
    'TK 15/11/2004 E2EM00002757 Start
    Set xmlCaseActivityNode = vxmlRequestNode.selectSingleNode(".//CASEACTIVITY")
    Set xmlAttrib = xmlCaseActivityNode.ownerDocument.createAttribute("ACTIVITYINSTANCE")
    xmlAttrib.Text = "1"
    xmlCaseActivityNode.Attributes.setNamedItem xmlAttrib
   
    xmlSetAttributeValue vxmlRequestNode, "OPERATION", "FindArchiveStageList"
    ' run Tm GetCaseTasks request and get the stage information for the current stage
    'PSC 09/12/2004 E2EM00003099 - Start
    On Error Resume Next
    RunTmRequest vxmlRequestNode.ownerDocument, xmlThisResponseDoc
    
    Dim lngErrNo As Long
    Dim strSource As String
    Dim strDescription As String
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
    
    On Error GoTo IsApplicationAtOfferStageExit
   
    If lngErrNo <> 0 And errGetOmigaErrorNumber(lngErrNo) <> oeRecordNotFound Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    'PSC 09/12/2004 E2EM00003099 - End
    
    Set xmlCaseStageArchiveResponseNodeList = xmlThisResponseDoc.selectNodes("//CASESTAGE")
    'TK 15/11/2004 E2EM00002757 End
    'TK 11/11/2004 E2EM00002757 If lngSequenceNumber > lngTMOfferStage Then
    ' PSC 23/11/2004 E2EM00002747
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
    'TK 11/11/2004 E2EM00002757 ElseIf lngSequenceNumber = lngTMOfferStage Then
    ' PSC 23/11/2004 E2EM00002747
    ElseIf lngSequenceNumber = lngTMOfferStage Then
        ' if the issue offer task is complete then is there a reissue offer task?
        strTMIssueOfferTaskName = GetGlobalParamString("TMIssueOffer") 'Get IssueOffer Task Name
        strPattern = "CASETASK[@TASKID='" & strTMIssueOfferTaskName & "']"
        Set xmlNodeList = xmlCaseStageResponseNode.selectNodes(strPattern) 'Get list of all IssueOffers
        'TK 15/11/2004 E2EM00002757 Start
        If (xmlNodeList.length = 0) Then
            If xmlCaseStageArchiveResponseNodeList.length > 0 Then
                For Each xmlCaseStageArchiveResponseNode In xmlCaseStageArchiveResponseNodeList
                    Set xmlNodeList = xmlCaseStageArchiveResponseNode.selectNodes(strPattern) 'Get list of all IssueOffers from CaseStageArchive
                    If xmlNodeList.length > 0 Then
                        Exit For ' Approve Case task has been completed.
                    End If
                Next
            End If
        End If
        'TK 15/11/2004 E2EM00002757 End
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
    'TK 14/09/2004 BBG1415 Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlCaseActivityNode = Nothing
    Set xmlAttrib = Nothing 'TK 11/11/2004 E2EM00002757
    'TK 15/11/2004 E2EM00002757 Start
    Set xmlStageResponseNode = Nothing
    Set xmlCaseStageArchiveResponseNode = Nothing
    Set xmlCaseStageArchiveResponseNodeList = Nothing
    'TK 15/11/2004 E2EM00002757 End
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Private Sub IssueOffer(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo IssueOfferVbErr
    Const cstrFunctionName As String = "IssueOffer"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlTMCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlQuoteNumberNode As IXMLDOMElement
    Dim xmlQuotationNode As IXMLDOMElement
    Dim xmlTempRequestDoc As FreeThreadedDOMDocument40
    Dim xmlRemodelQuotationRequest As IXMLDOMNode
    Dim xmlTMNewCaseTaskNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskDoc As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskElem As IXMLDOMElement
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlCustomerNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlPrintAttributeNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    Dim xmlCaseTaskCopy As IXMLDOMNode
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlTempReq As FreeThreadedDOMDocument40
    Dim xmlTempChildElem As IXMLDOMElement
    Dim xmlCreateResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTMUpdateCaseTaskRequest As IXMLDOMNode
    Dim xmlTMUpdateCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
    Dim objAppBO As Object
    Dim objBO As Object
    Dim objQuotation As Object
    Dim objAppQuoteBO As Object
    Dim gobjEnquiry As Object
    Dim lngErrorNumber As Long
    Dim dteApplicationApprovalDate As Date
    Dim strApplicationApprovalDate As String
    Dim strAcceptedQuoteNumber As String
    Dim strResponse As String
    Dim strOutputDocument As String
    Dim strMortgageSubQuoteNo As String
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strTMRemodelMortgage As String
    Dim strPrinterDestination As String
    Dim strTMReceivedOffer As String
    Dim strValidationType As String
    Dim strPattern As String
    Dim blnPrintingEnabled As Boolean
    Dim blnApplicationUnderReview As Boolean
    Dim blnValidMortgageSubQuote As Boolean
    Dim xmlCaseStageNode As IXMLDOMNode
    Dim xmlNewRequestNode As IXMLDOMNode
    Dim strSpecialGroup As String
    Dim xmlSpecialGroupNode As IXMLDOMNode
    Dim strDirectIndirectBusiness As String
    Dim xmlDirectIndirectBusiness As IXMLDOMNode
    
    blnPrintingEnabled = False
    
    ' check mandatory items
    xmlCheckMandatoryAttribute vxmlRequestNode, "UNITID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlApplicationNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationNode, "APPLICATIONFACTFINDNUMBER"
    
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "TASKID"
        
    ' start of Application Data retrieval
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)
    
    Set objAppBO = gobjContext.CreateInstance("omApp.ApplicationBO")
    strResponse = objAppBO.GetApplicationData(xmlThisRequestDoc.xml)
    
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    strApplicationApprovalDate = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/APPLICATIONAPPROVALDATE")
    
    If Len(strApplicationApprovalDate) = 0 Then
        errThrowError cstrFunctionName, 4814
    Else
        dteApplicationApprovalDate = CSafeDate(strApplicationApprovalDate)
    End If
    
    strAcceptedQuoteNumber = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/ACCEPTEDQUOTENUMBER")
    If Len(strAcceptedQuoteNumber) = 0 Then
        errThrowError cstrFunctionName, 4818
    End If

    strSpecialGroup = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/SPECIALGROUP")

    strDirectIndirectBusiness = xmlGetNodeText(xmlThisResponseDoc.documentElement, "APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND/DIRECTINDIRECTBUSINESS")
    
    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "QUOTATION")
    Set xmlQuotationNode = xmlGetNode(xmlThisRequestDoc, ".//QUOTATION")
    Set xmlNode = xmlThisRequestDoc.createElement("QUOTATIONNUMBER")
    xmlNode.Text = strAcceptedQuoteNumber
    xmlQuotationNode.appendChild xmlNode
    Set objQuotation = gobjContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
    strResponse = objQuotation.GetData(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    strMortgageSubQuoteNo = xmlGetNodeText(xmlThisResponseDoc, ".//MORTGAGESUBQUOTENUMBER")
    Set xmlQuotationNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlNode = Nothing
    Set xmlThisResponseDoc = Nothing
    ' start of checking for a re-modal quotation task
    Set xmlRemodelQuotationRequest = vxmlRequestNode.cloneNode(True)
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.appendChild xmlRemodelQuotationRequest
    xmlSetAttributeValue xmlRemodelQuotationRequest, "OPERATION", "FindCaseTaskList"
    strTMRemodelMortgage = GetGlobalParamString("TMRemodelMortgage")
    Set xmlTMCaseTaskNode = xmlGetMandatoryNode(xmlThisRequestDoc, "REQUEST/CASETASK")
    Set xmlTMNewCaseTaskNode = xmlThisRequestDoc.createElement("CASETASK")
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "CASEID"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYID"
    xmlCopyAttribute xmlTMCaseTaskNode, xmlTMNewCaseTaskNode, "ACTIVITYINSTANCE"
    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKID", strTMRemodelMortgage
    xmlSetAttributeValue xmlTMNewCaseTaskNode, "TASKSTATUS", "10"
    xmlRemodelQuotationRequest.removeChild xmlTMCaseTaskNode
    xmlRemodelQuotationRequest.appendChild xmlTMNewCaseTaskNode
    Set xmlThisResponseDoc = xmlCreateDOMObject
    On Error Resume Next
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
    ' PSC 07/02/02 SYS3995
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
    Set xmlThisResponseDoc = Nothing
    ' Check if the Application is under
    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True, "APPLICATIONREVIEWHISTORY")
    strResponse = objAppBO.IsAppUnderReview(xmlThisRequestDoc.xml)
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    blnApplicationUnderReview = xmlGetNodeAsBoolean(xmlThisResponseDoc.documentElement, "UNDERREVIEWINDICATOR")
    If blnApplicationUnderReview = True Then
        errThrowError cstrFunctionName, 4819
    End If
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
    Set objAppQuoteBO = gobjContext.CreateInstance(gstrAPPLICATIONQUOTE & ".ApplicationQuoteBO")
    strResponse = objAppQuoteBO.ValidateMortgageSubQuote(xmlThisRequestDoc.xml)
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    blnValidMortgageSubQuote = xmlGetNodeAsBoolean(xmlThisResponseDoc.documentElement, "VALIDMORTGAGESUBQUOTE")
    If blnValidMortgageSubQuote = False Then
        errThrowError cstrFunctionName, 4815
    End If
    ' PSC 27/03/01 SYS2053 - Start
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
    'TK 14/09/2004 BBG1415 Start
    Set xmlSpecialGroupNode = xmlThisRequestDoc.createElement("SPECIALGROUP")
    xmlApplicationNode.appendChild xmlSpecialGroupNode
    If Len(xmlSpecialGroupNode.Text) = 0 Then
        xmlSpecialGroupNode.Text = strSpecialGroup
    End If
    'TK 14/09/2004 BBG1415 End
    'TK 25/10/2004 BBG1697 Start
    Set xmlDirectIndirectBusiness = xmlThisRequestDoc.createElement("DIRECTINDIRECTBUSINESS")
    xmlApplicationNode.appendChild xmlDirectIndirectBusiness
    If Len(xmlDirectIndirectBusiness.Text) = 0 Then
        xmlDirectIndirectBusiness.Text = strDirectIndirectBusiness
    End If
    'TK 25/10/2004 BBG1697 End
    strResponse = objAppQuoteBO.RemodelQuotation(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    Set xmlQuoteNumberNode = Nothing
    Set objAppQuoteBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = Nothing
    ' Print Offer document
    strOutputDocument = xmlGetAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT")
        
    If Len(strOutputDocument) > 0 Then
        blnPrintingEnabled = True
            ' Do the printing thing
        On Error Resume Next
        Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
        ' PSC 07/02/02 SYS3995
        On Error GoTo IssueOfferVbErr
        If Not gobjEnquiry Is Nothing Then
            Set xmlThisCaseTaskDoc = xmlCreateDOMObject
            Set xmlThisCaseTaskElem = xmlThisCaseTaskDoc.createElement("REQUEST")
            Set xmlThisCaseTaskNode = xmlThisCaseTaskDoc.appendChild(xmlThisCaseTaskElem)
            xmlThisCaseTaskElem.setAttribute "OPERATION", "GetPrintAttributes"
                      
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("FINDATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            xmlTempElem.setAttribute "HOSTTEMPLATEID", strOutputDocument
              
            strResponse = gobjEnquiry.OmRequest(xmlThisCaseTaskElem.xml)
            Set xmlOut = xmlLoad(strResponse, cstrFunctionName)
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
            strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
            If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
                And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0) _
                And (InStr(1, strPrinterDestination, "W", vbTextCompare) > 0)) Then
               errThrowError cstrFunctionName, 4822, strPrinterDestination, "IssueOffer"
            End If
            Set xmlPrintAttributeNode = xmlTempElem
            Set xmlThisCaseTaskElem = Nothing
            Set xmlThisCaseTaskNode = Nothing
            Set xmlThisCaseTaskDoc = Nothing
            Set xmlThisCaseTaskDoc = xmlCreateDOMObject 'TK 16/09/2004 BBG1434
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
            Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("./APPLICATION")
            xmlCopyAttribute xmlApplicationNode, xmlTempNode, "APPLICATIONFACTFINDNUMBER"
            
            ' PSC 19/07/2004 BBG1144 - Start
            xmlSetAttributeValue xmlTempNode, "PRINTKFI", "1"
            xmlSetAttributeValue xmlTempNode, "QUOTATIONNUMBER", strAcceptedQuoteNumber
            xmlSetAttributeValue xmlTempNode, "MORTGAGESUBQUOTENUMBER", strMortgageSubQuoteNo
            ' PSC 19/07/2004 BBG1144 - End
            
            Set xmlCustomerNode = xmlApplicationNode.selectSingleNode("./CUSTOMER")
            
            'Note these Attributes are dummies as the OfferTemplate method in omPDM gets all the customer info
            'again - they are needed here because omTM.PrintDocumentforTask wants them
            xmlCopyAttribute xmlCustomerNode, xmlTempNode, "CUSTOMERVERSIONNUMBER"
            xmlCopyAttributeValue xmlCustomerNode, xmlCaseTaskCopy, "CUSTOMERNUMBER", "CUSTOMERIDENTIFIER"
            xmlThisCaseTaskNode.appendChild xmlTempNode
            Set xmlTempElem = xmlThisCaseTaskDoc.createElement("PRINTATTRIBUTES")
            Set xmlTempNode = xmlThisCaseTaskNode.appendChild(xmlTempElem)
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DEFAULTCOPIES", "COPIES"
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "PDMMETHOD", "METHODNAME"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
            xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE"
            xmlTempElem.setAttribute "DESTINATIONTYPE", Left$(strPrinterDestination, 1)
            Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("PRINTER")
            xmlTempElem.setAttribute "COMPRESSIONMETHOD", xmlGetAttributeText(xmlPrinterNode, "COMPRESSIONMETHOD")
            If strPrinterDestination = "Local" Or strPrinterDestination = "Workstation Printer" Then
                xmlTempElem.setAttribute "PRINTER", xmlGetAttributeText(xmlPrinterNode, "PRINTERNAME")
            Else
                xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "REMOTEPRINTERLOCATION", "PRINTER"
            End If
            
            ' PSC 06/10/2004 BBG1572
            xmlSetAttributeValue xmlTempNode, "WEBDOCUMENTTYPE", "OFFER"
              
            xmlThisCaseTaskNode.appendChild xmlTempNode
            PrintDocumentForTask xmlThisCaseTaskElem, vxmlResponseNode
                
        End If
            
        Set xmlOut = Nothing
        Set xmlThisCaseTaskDoc = Nothing
        Set xmlThisCaseTaskElem = Nothing
        Set xmlThisCaseTaskNode = Nothing
        'TK 14/09/2004 BBG1415 Set xmlCustomerNodeList = Nothing
        Set xmlCustomerNode = Nothing
        Set xmlTempNode = Nothing
        Set xmlTempElem = Nothing
        Set xmlPrintAttributeNode = Nothing
        Set xmlPrinterNode = Nothing
        Set gobjEnquiry = Nothing
        Set xmlCaseTaskCopy = Nothing
    End If
    'Generate audit record of the offer produced on ApplicationOffer table
    'Construct the element based request:
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
       
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
                     
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlTempReq = xmlCreateDOMObject
    Set xmlTempElem = xmlTempReq.createElement("APPLICATIONOFFER")
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONNUMBER")
    xmlTempChildElem.Text = strApplicationNumber
    xmlTempElem.appendChild xmlTempChildElem
    Set xmlTempChildElem = xmlTempReq.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTempChildElem.Text = strApplicationFactFindNumber
    xmlTempElem.appendChild xmlTempChildElem
       
    xmlNewRequest.appendChild xmlTempElem
    'BG call omApp.ApplicationBO_UpdateApplicationOffer method with element based request.
    strResponse = objAppBO.CreateApplicationOffer(xmlNewRequest.xml)
    Set xmlCreateResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlCreateResponseDoc.documentElement, , True
    ' Update task as complete
    ' PrintDocumentForTask updates the task as complete so only do this if printing
    ' is not available
    'If blnPrintingEnabled = False Then
        
        'Set xmlThisRequestDoc = xmlCreateDOMObject
        'Set xmlTMUpdateCaseTaskRequest = vxmlRequestNode.cloneNode(True)
        'xmlThisRequestDoc.appendChild xmlTMUpdateCaseTaskRequest
        'xmlSetAttributeValue xmlTMUpdateCaseTaskRequest, "OPERATION", "UpdateCaseTask"
        '
        'Set xmlTMUpdateCaseTaskNode = xmlGetMandatoryNode(xmlTMUpdateCaseTaskRequest, "CASETASK")
        'xmlSetAttributeValue xmlTMUpdateCaseTaskNode, "TASKSTATUS", "40"
        'Set xmlThisResponseDoc = xmlCreateDOMObject
        'RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        'Set xmlThisRequestDoc = Nothing
        'errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        'Set xmlThisResponseDoc = Nothing
        'Set xmlTMUpdateCaseTaskRequest = Nothing
        'Set xmlTMUpdateCaseTaskNode = Nothing
    'End If
    'BG Generate the received offer task or update an existing one with the correct applicationsequencenumber
    'get current CASESTAGE detail
    'BBG609 - Set xmlCaseTaskElem = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlThisRequestDoc = xmlCreateDOMObject
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlElem = xmlThisRequestDoc.createElement("CASESTAGE")
    'BBG609 - xmlCopyAttribute xmlCaseTaskElem, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
    xmlCopyAttributeValue xmlApplicationElem, xmlElem, "APPLICATIONNUMBER", "CASEID"
    'BBG609 - xmlCopyAttribute xmlCaseTaskElem, xmlElem, "ACTIVITYID"
    xmlCopyAttribute xmlCaseTaskNode, xmlElem, "ACTIVITYID"
    xmlThisRequestNode.appendChild xmlElem
    Set xmlElem = xmlThisRequestDoc.createElement("APPLICATION")
    xmlCopyAttribute xmlApplicationElem, xmlElem, "APPLICATIONPRIORITY"
    xmlThisRequestNode.appendChild xmlElem
    Set xmlCurrentStageDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlCurrentStageDoc.async = False
    GetCurrentStage xmlThisRequestNode, xmlCurrentStageDoc
    strTMReceivedOffer = GetGlobalParamString("TMReceivedOffer")
    strValidationType = GetFirstComboValueId("TaskStatus", "I")
    strPattern = "RESPONSE/CASESTAGE/CASETASK[@TASKID='" & strTMReceivedOffer & "' and @TASKSTATUS='" & strValidationType & "']"
    'BBG609 - Set xmlCaseTaskElem = xmlGetNode(xmlCurrentStageDoc, strPattern)
    Set xmlCaseTaskNode = xmlGetNode(xmlCurrentStageDoc, strPattern)
    'BBG609 - Set xmlCaseStageElem = xmlGetNode(xmlCurrentStageDoc, "RESPONSE/CASESTAGE")
    Set xmlCaseStageNode = xmlGetNode(xmlCurrentStageDoc, "RESPONSE/CASESTAGE")

    If Not xmlCaseStageNode Is Nothing And Not xmlCaseTaskNode Is Nothing Then
        'Call TaskManagementBO.UpdateCaseTask
        'Build the Request and update the case task
                   
        'BBG609 - Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        Set xmlNewRequestNode = xmlGetRequestNode(vxmlRequestNode)
        'BBG609 - xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
        xmlSetAttributeValue xmlNewRequestNode, "OPERATION", "UPDATECASETASK"
        'BBG609 - xmlSetAttributeValue xmlCaseTaskElem, "CONTEXT", xmlCreateResponseDoc.selectSingleNode("RESPONSE/OFFER/OFFERSEQUENCENO").Text
        xmlSetAttributeValue xmlCaseTaskNode, "CONTEXT", xmlCreateResponseDoc.selectSingleNode("RESPONSE/OFFER/OFFERSEQUENCENO").Text
        'BBG609 - xmlNewRequestNode.appendChild xmlCaseTaskElem
        xmlNewRequestNode.appendChild xmlCaseTaskNode
        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlNewRequestNode.xml)
        errCheckXMLResponse strResponse, True
    Else
        Set xmlTempRequestDoc = xmlCreateDOMObject
        Set xmlThisRequestNode = xmlTempRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "CreateAdhocCaseTask"
        Set xmlElem = xmlTempRequestDoc.createElement("APPLICATION")
        xmlCopyAttribute xmlApplicationElem, xmlElem, "APPLICATIONPRIORITY"
        xmlThisRequestNode.appendChild xmlElem
        'BBG609 - Set xmlCaseTaskElem = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
        Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
        Set xmlElem = xmlTempRequestDoc.createElement("CASETASK")
        'BBG609 - xmlCopyAttribute xmlCaseTaskElem, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlCaseTaskNode, xmlElem, "SOURCEAPPLICATION"
        xmlCopyAttributeValue xmlApplicationElem, xmlElem, "APPLICATIONNUMBER", "CASEID"
        'BBG609 - xmlCopyAttribute xmlCaseTaskElem, xmlElem, "ACTIVITYID"
        xmlCopyAttribute xmlCaseTaskNode, xmlElem, "ACTIVITYID"
        'BBG609 - xmlCopyAttribute xmlCaseStageElem, xmlElem, "STAGEID"
        xmlCopyAttribute xmlCaseStageNode, xmlElem, "STAGEID"
        'BBG609 - xmlCopyAttribute xmlCaseStageElem, xmlElem, "CASESTAGESEQUENCENO"
        xmlCopyAttribute xmlCaseStageNode, xmlElem, "CASESTAGESEQUENCENO"
        'BBG609 - xmlCopyAttribute xmlCaseStageElem, xmlElem, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlCaseStageNode, xmlElem, "ACTIVITYINSTANCE"
        xmlSetAttributeValue xmlElem, "TASKID", strTMReceivedOffer
        xmlSetAttributeValue xmlElem, "CONTEXT", xmlCreateResponseDoc.selectSingleNode("RESPONSE/OFFER/OFFERSEQUENCENO").Text
        xmlThisRequestNode.appendChild xmlElem
                 
        CreateAdhocCaseTask xmlThisRequestNode
    End If
    If Not gobjContext Is Nothing Then
        gobjContext.SetComplete
    End If
                   
IssueOfferExit:
    
    Set objAppBO = Nothing
    Set objBO = Nothing
    Set objQuotation = Nothing
    Set objAppQuoteBO = Nothing
    Set gobjEnquiry = Nothing
    Set xmlSpecialGroupNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTMCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlNode = Nothing
    Set xmlQuoteNumberNode = Nothing
    Set xmlQuotationNode = Nothing
    Set xmlTempRequestDoc = Nothing
    Set xmlRemodelQuotationRequest = Nothing
    Set xmlTMNewCaseTaskNode = Nothing
    Set xmlOut = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlPrintAttributeNode = Nothing
    Set xmlPrinterNode = Nothing
    Set xmlCaseTaskCopy = Nothing
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTempReq = Nothing
    Set xmlTempChildElem = Nothing
    Set xmlCreateResponseDoc = Nothing
    Set xmlTMUpdateCaseTaskRequest = Nothing
    Set xmlTMUpdateCaseTaskNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlCurrentStageDoc = Nothing
    Set xmlCaseStageNode = Nothing
    Set xmlNewRequestNode = Nothing
    Set xmlSpecialGroupNode = Nothing
    'TK 25/10/2004 BBG1697
    Set xmlDirectIndirectBusiness = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
IssueOfferVbErr:
    If Not gobjContext Is Nothing Then
        gobjContext.SetAbort
    End If
    GoTo IssueOfferExit
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
    Set xmlAppRec = Nothing
ApproveRecommendApplicationExit:
    
    
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub RunCreditCheck(ByVal vxmlRequestNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' procedure:    RunCreditCheck
' author:       Mark Coleman
' date:         22/08/2002
' description:  Runs a credit score, full bureau download and imports the bureau data into Omiga
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
' return:       N/A
' notes:        The Credit Check and Bureau Download are run in a new transaction. This is so
'               that the Credit_Check task can be committed independently from the Import_Bureau_Data.
'               If the subsequent ImportBureauData process fails, an adhoc case task can be created
'               so that it can be re-run later.
'------------------------------------------------------------------------------------------
On Error GoTo RunCreditCheckExit
    
Const cstrFunctionName As String = "RunCreditCheck"
    
Dim objNTxBO As IOmTmNTxBO
Dim xmlTempResponse As IXMLDOMNode
Dim xmlTempDoc As FreeThreadedDOMDocument40
Dim xmlTempRequest As IXMLDOMElement
    
    'Initialise ======================================================================
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.async = False
    Set xmlTempResponse = xmlTempDoc.createElement("RESPONSE")
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".OmTmNTxBO")
    'Validate ========================================================================
    xmlCheckMandatoryNode vxmlRequestNode, "APPLICATION"
    xmlCheckMandatoryNode vxmlRequestNode, "CASETASK"
    'Do Credit Score and Bureau Download =============================================
    objNTxBO.RunExperianCreditCheck vxmlRequestNode
    'Import Bureau data ==============================================================
    Set xmlTempRequest = vxmlRequestNode.cloneNode(True)
    xmlTempRequest.setAttribute "CREDITCHECK", "1"
    Call ImportBureauData(xmlTempRequest)
RunCreditCheckExit:
   
    Set objNTxBO = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTempDoc = Nothing
    Set xmlTempRequest = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub RunRiskAssessment(ByVal vxmlRequestNode As IXMLDOMNode)
    On Error GoTo RunRiskAssessmentExit
    Const cstrFunctionName As String = "RunRiskAssessment"
    Dim xmlApplicationElem As IXMLDOMElement, xmlNewRequest As IXMLDOMElement
    Dim xmlCaseTakElem As IXMLDOMElement
    Dim objBO As Object
    Dim strApplicationNumber As String, strAFFNumber As String
    Dim strResponse As String
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
    'Create Request and run RiskAssessment
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    Set xmlApplicationElem = vxmlRequestNode.selectSingleNode(".//APPLICATION")
    ' PSC 20/11/01 SYS2867/SYS2988
    xmlNewRequest.appendChild xmlMakeNodeElementBased(xmlApplicationElem, False, "RISKASSESSMENT")
    Set objBO = gobjContext.CreateInstance("omRA.RiskAssessmentBO")
    strResponse = objBO.RunRiskAssessment(xmlNewRequest.xml)
    Dim lngErrNo As Long
    lngErrNo = errCheckXMLResponse(strResponse, False)
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, 4813
    End If
    ' Now update Case Task
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
RunRiskAssessmentExit:
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTakElem = Nothing
    Set objBO = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub ValidateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo ValidateValuationReportExit
    
    Const cstrFunctionName As String = "ValidateValuationReport"
    
'    Dim strValResponse As String
    Dim xmlCaseNode  As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlCaseTaskReq As IXMLDOMElement
    Dim xmlValRequest As IXMLDOMElement
    Dim xmlValuation As IXMLDOMElement
    Dim xmlReqValuation As IXMLDOMElement
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlDOMCaseResp As FreeThreadedDOMDocument40
    
    Dim objAppProcBO As Object
    
    ' Setup default response block
    Set xmlThisResponseDoc = xmlCreateDOMObject
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
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
    If objAppProcBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    'add response from valuation call to this response node
    Dim xmlAppStatusNode As IXMLDOMNode
    Dim xmlLTVNode As IXMLDOMNode
    xmlThisResponseDoc.loadXML objAppProcBO.omAppProcRequest(xmlValRequest.xml)
    Set objAppProcBO = Nothing
    
    'parse response
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    Set xmlAppStatusNode = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatusNode Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatusNode
    End If
    Set xmlLTVNode = xmlThisResponseDoc.selectSingleNode(".//LTV")
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
    Set xmlDOMCaseResp = xmlCreateDOMObject
    xmlDOMCaseResp.loadXML gobjTmBO.TmRequest(xmlCaseTaskReq.xml)
    errCheckXMLResponseNode xmlDOMCaseResp.selectSingleNode("RESPONSE"), , True
    ' Append to the response block
    AttachResponseData xmlDOMCaseResp.documentElement, xmlResponseNode
ValidateValuationReportExit:
    Set xmlCaseTaskReq = Nothing
    Set xmlResponseNode = Nothing
    Set xmlAppStatusNode = Nothing
    Set xmlLTVNode = Nothing
    Set xmlCaseNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlValRequest = Nothing
    Set xmlValuation = Nothing
    Set xmlReqValuation = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlDOMCaseResp = Nothing
    Set objAppProcBO = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
End Sub
Private Sub CreateValuationReport(ByVal vxmlRequestNode As IXMLDOMNode)
    On Error GoTo CreateValuationReportExit
    Dim strValResponse As String
    'TK 14/09/2004 BBG1415 Dim xmlOperationNode As IXMLDOMElement
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
    strValResponse = objAppProcBO.omAppProcRequest(xmlValRequest.xml)
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
    Set xmlValElem = Nothing
    Set xmlValRequest = Nothing
    Set objAppProcBO = Nothing
    Set xmlCaseReq = Nothing
    Set xmlTableNode = Nothing
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
    
    Const cstrFunctionName As String = "CreateValuationReportNoInst"
    
    'Declare Variables
    Dim xmlExtractDOM As FreeThreadedDOMDocument40
    Dim xmlValuerResponse As FreeThreadedDOMDocument40
    Dim xmlValuerRequest As FreeThreadedDOMDocument40
    Dim xmlComboRequest As FreeThreadedDOMDocument40
    Dim xmlComboResponse As FreeThreadedDOMDocument40
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
    
    'amend request node
    xmlChangeNodeName vxmlRequestNode, "APPLICATION", "VALUATION"
        
    'set up connection to Application Processing & Third Party Business Objects
    Set objThirdPartyBO = gobjContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO")
    Set objAppProcBO = gobjContext.CreateInstance(gstrAPPLICATIONPROCESSING & ".omAppProcBO")
    Set objComboDO = gobjContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        
    If objThirdPartyBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrTHIRDPARTY_COMPONENT & ".ThirdPartyBO"
    End If
    
    If objAppProcBO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrAPPLICATIONPROCESSING & ".omAppProcBO"
    End If
    
    If objComboDO Is Nothing Then
        errThrowError cstrFunctionName, oeObjectNotCreatable, gstrBASE_COMPONENT & ".ComboDO"
    End If
    
    Set xmlExtractDOM = xmlCreateDOMObject
    Set xmlValuerRequest = xmlCreateDOMObject
    Set xmlValuerResponse = xmlCreateDOMObject
    Set xmlComboRequest = xmlCreateDOMObject
    Set xmlComboResponse = xmlCreateDOMObject
    
    'DPF 5/11/2002 - BMIDS00852 - use Combo Assist to pull back valuer id for Validation type 'L'
    'TK 30/06/2004 BBG895 Standard valuation type changes to H
    strStandardValuer = "H"
    Set xmlList = xmlComboRequest.createElement("LIST")
    xmlComboRequest.appendChild xmlList
    Set xmlNode = xmlComboRequest.createElement("LISTNAME")
    xmlNode.Text = "ValuationType"
    xmlList.appendChild xmlNode
    strComboList = objComboDO.GetComboList(xmlList.xml)
    Set xmlComboResponse = xmlLoad(strComboList, "CreateValuationReportNoInst")
        
    Set xmlValuationType = xmlComboResponse.selectSingleNode("//LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE = '" & strStandardValuer & "']")
    'DPF - END OF BMIDS00852
      
    'Retrieve the Directory GUID of a default Valuer set up in supervisor
    Set xmlValuerReqElem = xmlGetRequestNode(vxmlRequestNode)
    xmlValuerReqElem.setAttribute "OPERATION", "FindPanelValuerList"
    Set xmlValuerList = xmlValuerRequest.createElement("VA_PANELVALUERLIST")
    Set xmlValuerValType = xmlValuerRequest.createElement("VALUATIONTYPE")
    Set xmlValuerPanelID = xmlValuerRequest.createElement("PANELID")
    xmlValuerPanelID.Text = "Default"
    xmlValuerValType.Text = xmlValuationType.selectSingleNode("VALUEID").Text 'DPF 05/11/2002 - BMIDS00852
    xmlValuerList.appendChild xmlValuerValType
    xmlValuerList.appendChild xmlValuerPanelID
    xmlValuerReqElem.appendChild xmlValuerList
    xmlValuerRequest.appendChild xmlValuerReqElem
        
    strValuerResponse = objThirdPartyBO.FindPanelValuerList(xmlValuerRequest.xml)
    'Check Response
    errCheckXMLResponse strValuerResponse, True
        
    'Assign directory guid retrieved into Variable
    'TK 14/09/2004 BBG1415 xmlValuerResponse.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlValuerResponse.validateOnParse = False
    xmlValuerResponse.loadXML (strValuerResponse)
    'AW     20/12/2002  BM0190
    'strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "VA_PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
    strDirectoryGUID = xmlGetMandatoryNode(xmlValuerResponse.documentElement, "PANELVALUERLISTLIST/VA_PANELVALUERLIST/DIRECTORYGUID").Text
        
    'Create 'Dummy' Valuer Instruction record
    Set xmlInstElem = xmlGetMandatoryNode(vxmlRequestNode, "./VALUATION").cloneNode(True)
    xmlChangeNodeName xmlInstElem, "VALUATION", "VALUERINSTRUCTION"
    xmlInstElem.setAttribute "VALUERPANELNO", "Default"
    xmlInstElem.setAttribute "DIRECTORYGUID", strDirectoryGUID
    Set xmlInstRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlInstRequest.setAttribute "OPERATION", "CreateValuerInstructions"
    xmlInstRequest.appendChild xmlInstElem.cloneNode(True)
    strInstResponse = objAppProcBO.omAppProcRequest(xmlInstRequest.xml)
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
            
    strReportResponse = objAppProcBO.omAppProcRequest(xmlValRequest.xml)
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
    
    '<VSA> VisualStudio Analyser Support
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
    Set objAppProcBO = Nothing
    Set objThirdPartyBO = Nothing
    Set objComboDO = Nothing
    Set xmlTableNode = Nothing
    Set xmlValuationType = Nothing
            
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
CreateValuationReportNoInstErr:
    gobjContext.SetAbort
    'TK 30/06/2004 BBG895 - MsgBox "Source:  " & Err.Source & " Description:  " & Err.Description & " Number:  " & Err.Number
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
    Dim intCaseStageSeqNo As Integer ', lngErrNo As Long
    Dim strCondition As String, strTaskStatus As String
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
    Set xmlResponseDoc = xmlCreateDOMObject
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
    strResponse = objAppProcBO.omAppProcRequest(xmlTempRequest.xml)
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
    strReportResponse = objAppProcBO.omAppProcRequest(xmlCreateRepElem.xml)
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
    Set xmlCurrentStageDoc = Nothing
    Set objComboDO = Nothing
    Set xmlNode = Nothing
    Set xmlCreateRepElem = Nothing
    Set xmlApplicationElem = Nothing
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    ' Call referencesBO.CreateReportOnTitle
    Set objBO = gobjContext.CreateInstance("omROT.omRotBO")
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.setProperty "NewParser", True
    'TK 14/09/2004 BBG1415 xmlThisResponseDoc.validateOnParse = False
    Set xmlThisResponseDoc = xmlCreateDOMObject
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
    'TK 14/09/2004 BBG1415 Set xmlThisResponseDoc = Nothing
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
      
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    Dim xmlChildNode As IXMLDOMNode
        
    Set xmlElement = vxmlRequestNode
    Set xmlIn = xmlCreateDOMObject
    Set xmlOut = xmlCreateDOMObject
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
    Select Case Left$(Trim$(strContactTypeName), 3)
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
Private Sub ProcessAutomaticTasks(ByVal vxmlRequest As IXMLDOMNode)
    On Error GoTo ProcessAutomaticTasksExit
    Const strFunctionName As String = "ProcessAutomaticTasks"
    Dim blnAutomaticTaskInd As Boolean
    Dim strInterface As String
    Dim strOutputDocument As String
'    Dim strTaskId As String
    Dim strTaskName As String
    Dim strPrinterDestination As String
    Dim strApplicant As String, strCustomerIdentifier As String
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskDoc As New FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskElem As IXMLDOMElement
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlCustomerNodeList As IXMLDOMNodeList
    Dim xmlCustomerNode As IXMLDOMNode
    Dim strTmpResponse As String
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlPrintAttributeNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    Dim gobjEnquiry As Object
    xmlThisCaseTaskDoc.async = False
    Set xmlCaseTaskNode = vxmlRequest.selectSingleNode("./CASETASK")
    blnAutomaticTaskInd = xmlGetAttributeAsBoolean(xmlCaseTaskNode, "AUTOMATICTASKIND")
       
    If blnAutomaticTaskInd Then
             
        strTaskName = xmlGetAttributeText(xmlCaseTaskNode, "TASKNAME")
        strInterface = xmlGetAttributeText(xmlCaseTaskNode, "INTERFACE")
        strOutputDocument = xmlGetAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT")
           
        If strInterface <> "" Then
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
            xmlThisCaseTaskNode.appendChild xmlCaseTaskNode
            Set xmlApplicationNode = vxmlRequest.selectSingleNode("./APPLICATION")
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
        If strOutputDocument <> "" And strInterface = "" Then
            '
            ' Call OmPrintBO.GetPrintAtributes
            '
            ' PSC 26/11/01 SYS3195 - Start
            On Error Resume Next
            Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
            On Error GoTo ProcessAutomaticTasksExit
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
                    
                strTmpResponse = gobjEnquiry.OmRequest(xmlThisCaseTaskElem.xml)
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
                If xmlTempElem.getAttribute("DEFAULTCOPIES") = "" Then
                  errThrowError strFunctionName, 4823, "", strTaskName
                End If
                strPrinterDestination = xmlGetAttributeText(xmlTempElem, "PRINTERDESTINATIONTYPE")
                strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
                If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
                    And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0) _
                    And (InStr(1, strPrinterDestination, "W", vbTextCompare) > 0)) Then
                        errThrowError strFunctionName, 4822, strPrinterDestination, strTaskName
                End If
                Set xmlPrintAttributeNode = xmlTempElem
    '
    '           Call OmTmBo.PrintDocumentForTask
    '
                Set xmlThisCaseTaskElem = Nothing
                Set xmlThisCaseTaskNode = Nothing
                Set xmlThisCaseTaskDoc = Nothing
    '           Set xmlThisCaseTaskDoc = New xmlDocument
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
                strApplicant = xmlGetAttributeText(xmlCaseTaskNode, "APPLICANT")
                '
                'AQR 2266 DRC - changed to check for applicant and then look up customer node
                'to get the Customerversion for the customer whose number matches the
                'identifier
                strCustomerIdentifier = xmlGetAttributeText(xmlCaseTaskNode, "CUSTOMERIDENTIFIER")
                If (strApplicant = "1") And (strCustomerIdentifier <> "") Then
                   Set xmlCustomerNodeList = xmlApplicationNode.selectNodes("CUSTOMER")
                   For Each xmlCustomerNode In xmlCustomerNodeList
                     If xmlGetAttributeText(xmlCustomerNode, "CUSTOMERNUMBER") = strCustomerIdentifier Then
                        xmlCopyAttribute xmlCustomerNode, xmlTempNode, "CUSTOMERVERSIONNUMBER"
                     End If
                   Next
                End If
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
                xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE"
                'AQR SYS2266 - DRC
                xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
                xmlTempElem.setAttribute "DESTINATIONTYPE", Left$(strPrinterDestination, 1)
                If strPrinterDestination = "Local" Or strPrinterDestination = "Workstation Printer" Then
                    Set xmlPrinterNode = vxmlRequest.selectSingleNode("PRINTER")
                    xmlTempElem.setAttribute "PRINTER", xmlGetAttributeText(xmlPrinterNode, "PRINTERNAME")
                Else
    '
    '               strPrinterDestination = "Remote" Or "File"
    '
                     xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, _
                                          "REMOTEPRINTERLOCATION", "PRINTER"
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
        If strOutputDocument = "" And strInterface = "" Then
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
    End If
ProcessAutomaticTasksExit:
    Set xmlOut = Nothing
    Set xmlThisCaseTaskDoc = Nothing
    Set xmlThisCaseTaskElem = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerNodeList = Nothing
    Set xmlCustomerNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempElem = Nothing
    Set xmlPrintAttributeNode = Nothing
    Set xmlPrinterNode = Nothing
    Set gobjEnquiry = Nothing
    errCheckError strFunctionName
End Sub
Private Sub ValidateReportOnTitle(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo ValidateReportOnTitleErr
    Const cstrFunctionName As String = "ValidateReportOnTitle"
                               
    Dim objBO As Object
    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlRequestCaseTaskNode As IXMLDOMNode
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlROTNode As IXMLDOMNode
    Dim xmlElem As IXMLDOMElement
    Dim xmlThisTaskNode  As IXMLDOMNode
    Dim xmlAppStatus As IXMLDOMNode
    
    Dim blnReissueOfferInd As Boolean
    
    Dim strTMReissueOffer As String
    Dim strTaskDueDateAndTime As String
    Dim strResponse As String

    blnReissueOfferInd = False
    
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    
'Stage1:
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CASETASK"
    xmlGetMandatoryNode vxmlRequestNode, "APPLICATION"
    
    ' Call omRotBO.ValidateReportOnTitle
    Set objBO = gobjContext.CreateInstance("omROT.omRotBO")
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.loadXML objBO.OmRotRequest(vxmlRequestNode.xml)
    Set objBO = Nothing
    
    errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    
    Set xmlROTNode = xmlThisResponseDoc.selectSingleNode(".//ROT")
    If Not xmlROTNode Is Nothing Then
        blnReissueOfferInd = xmlGetAttributeAsBoolean(xmlROTNode, "REISSUEOFFERIND")
        vxmlResponseNode.appendChild xmlROTNode.cloneNode(True)
    End If
    
    Set xmlAppStatus = xmlThisResponseDoc.selectSingleNode(".//APPSTATUS")
    If Not xmlAppStatus Is Nothing Then
        vxmlResponseNode.appendChild xmlAppStatus
    End If
    
    Set xmlRequestCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    
'Stage2:
    If blnReissueOfferInd = True Then
        
        strTMReissueOffer = GetGlobalParamString("TMReissueOffer")
        'Get TaskDueDateAndTime Used when Creating CaseTask
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
        Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetTaskList"
        Set xmlElem = xmlThisRequestDoc.createElement("TASK")
        xmlElem.setAttribute "TASKID", strTMReissueOffer
        xmlThisRequestNode.appendChild xmlElem
        Set objBO = gobjContext.CreateInstance("MsgTm.MsgTmBO")
        strResponse = objBO.TmRequest(xmlThisRequestNode.xml)
        xmlThisResponseDoc.setProperty "NewParser", True
        xmlThisResponseDoc.validateOnParse = False
        xmlThisResponseDoc.loadXML strResponse
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
        Set xmlThisTaskNode = xmlThisResponseDoc.selectSingleNode(".//TASK")
        If Not xmlThisTaskNode Is Nothing Then ' error ???
            SetCaseTaskDueDateAndTime xmlThisTaskNode, xmlThisTaskNode
        End If
           
        strTaskDueDateAndTime = xmlGetAttributeText(xmlThisTaskNode, "TASKDUEDATEANDTIME", "")
        ' create CreateCaseTask request
        Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
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
        xmlSetAttributeValue xmlElem, "TASKID", strTMReissueOffer
        xmlSetAttributeValue xmlElem, "TASKDUEDATEANDTIME", strTaskDueDateAndTime
        xmlThisRequestNode.appendChild xmlElem
        ' run Tm CreateCaseTask request
        xmlThisResponseDoc.async = False
        RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
        ' parse response
        errCheckXMLResponseNode xmlThisResponseDoc.selectSingleNode("RESPONSE"), , True
    End If
'Stage3:
    
    'MSla 30/06/2004 BBG882
    'If blnSolBankAccMatch = True Then
       
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
    
    'End If
    'MSla 30/06/2004 BBG882 - End
    gobjContext.SetComplete
        
ValidateReportOnTitleExit:
    
    Set objBO = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlRequestCaseTaskNode = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlElem = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlROTNode = Nothing
    Set xmlThisTaskNode = Nothing
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
    Set xmlTempReq = xmlCreateDOMObject
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
    Set objBO = Nothing
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTmRequest = Nothing
    Set xmlTempReq = Nothing
    Set xmlTempElem = Nothing
    Set xmlTempChildElem = Nothing
    Set xmlCaseTaskElem = Nothing
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
Private Sub MuraliTemp_RunHunterInterface(ByVal vxmlRequestNode As IXMLDOMNode)
    
    'For future enhancements - Please Do not Delete - waiting for reply from BBG
    
    On Error GoTo MuraliTemp_RunHunterInterfaceExit
    
    Const cstrFunctionName As String = "MuraliTemp_RunHunterInterface"
    
    Dim xmlApplicationElem As IXMLDOMElement
    Dim xmlNewRequest As IXMLDOMElement
    Dim xmlCaseTaskElem As IXMLDOMElement
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisResponseDoc  As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlFindCaseTaskList As IXMLDOMNode
    Dim xmlTMNewCaseTaskNode  As IXMLDOMElement
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlCaseTaskNodeList As IXMLDOMNodeList
    Dim xmlCaseTask As IXMLDOMNode
    
    Dim objBO As Object
    Dim objAppBO  As Object
    
    Dim strApplicationNumber As String
    Dim strAFFNumber As String
    Dim strResponse As String
    Dim strValidationType As String
    Dim strErrDescription As String
    Dim strCaseID As String
    
    Dim lngErrNo As Long
    
    Dim intTaskStatus As Integer
    
    Dim blnHunterInterfaceTaskCompleted As Boolean
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    
    Set xmlApplicationElem = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    strApplicationNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetMandatoryAttributeText(xmlApplicationElem, "APPLICATIONFACTFINDNUMBER")
        
    If Not gobjContext Is Nothing Then
        Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
        Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    Else
        errThrowError cstrFunctionName, oeObjectNotCreatable, " MsgTm and omApp "
    End If
    
    'Check whether the Hunter Task has been completed in the previous stages if yes then return the taskstatus
    'completed else returns record not found which is not an error
    'MV Start
    strCaseID = xmlGetAttributeText(vxmlRequestNode.selectSingleNode(".//CASETASK"), "CASEID")
    
    Set xmlThisRequestNode = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlThisRequestNode, "OPERATION", "FindCaseTaskList"
    
    Set xmlFindCaseTaskList = xmlOut.createElement("CASETASK")
    xmlThisRequestNode.appendChild xmlFindCaseTaskList
    xmlSetAttributeValue xmlFindCaseTaskList, "SOURCEAPPLICATION", "OMIGA"
    xmlSetAttributeValue xmlFindCaseTaskList, "CASEID", strCaseID
    xmlSetAttributeValue xmlFindCaseTaskList, "TASKID", "Hunter_Interface"
    xmlSetAttributeValue xmlFindCaseTaskList, "TASKDUEDATEANDTIME", Format(Now(), "DD/MM/YYYY")
    
    xmlThisRequestDoc.appendChild xmlThisRequestNode
    
    Set xmlThisResponseDoc = xmlCreateDOMObject
    
    RunTmRequest xmlThisRequestDoc, xmlThisResponseDoc
    errCheckXMLResponse xmlThisResponseDoc.xml, False
    
    blnHunterInterfaceTaskCompleted = False
    
    Set xmlCaseTaskNodeList = xmlThisResponseDoc.documentElement.selectNodes("CASETASK")
    For Each xmlCaseTask In xmlCaseTaskNodeList
        intTaskStatus = CSafeInt(xmlGetAttributeText(xmlCaseTask, "TASKSTATUS"))
        If intTaskStatus = 40 Then
            blnHunterInterfaceTaskCompleted = True
            Exit For
        End If
    Next
    
    'If the task is not completed then Generate Hunter File else update the task with status as completed
    If Not blnHunterInterfaceTaskCompleted Then
        
        'Create Request and run Hunter Interface
        Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
        Set xmlApplicationElem = vxmlRequestNode.selectSingleNode(".//APPLICATION")
        xmlNewRequest.appendChild xmlApplicationElem
        
        Set objBO = gobjContext.CreateInstance(gstrHUNTERINTERFACE_COMPONENT & ".omHIBO")
        
        strResponse = objBO.GenerateRecords(xmlNewRequest.xml)
        
        On Error Resume Next
        errCheckXMLResponse strResponse, True
        lngErrNo = Err.Number
        strErrDescription = Err.Description
        On Error GoTo MuraliTemp_RunHunterInterfaceExit
        If lngErrNo <> 0 Then
            errThrowError cstrFunctionName, oeTmHunterInterfaceFailed, strErrDescription
        End If
        
    End If
    'End
    
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
    
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
MuraliTemp_RunHunterInterfaceExit:
    
    Set xmlApplicationElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing
    Set xmlOut = Nothing
    Set xmlFindCaseTaskList = Nothing
    Set xmlTMNewCaseTaskNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlCaseTaskNodeList = Nothing
    Set objBO = Nothing
    Set objAppBO = Nothing
        
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
    Dim lngErrNo As Long
    Dim strErrDescription As String
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

    Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")

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
Dim xmlResponseNode As IXMLDOMNode
Dim xmlTempRequest As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlStageTask As IXMLDOMNode
Dim xmlCaseTask As IXMLDOMElement
Dim xmlCaseActivity As IXMLDOMElement
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlApplication As IXMLDOMNode
Dim xmlNode As IXMLDOMNode
Dim xmlTaskElem As IXMLDOMElement
Dim strImportBureauDataTaskId As String
Dim strResponse As String
Dim lngErr As Long
Dim intCaseStageSeqNo As Integer
    'Initialise
    Set xmlResponseDoc = xmlCreateDOMObject
    'TK 14/09/2004 BBG1415 xmlResponseDoc.async = False
    'Validate
    Set xmlCaseTask = xmlGetMandatoryNode(vxmlRequestNode, ".//CASETASK")
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, ".//APPLICATION")
    'Call NTxBO.ImportBureauData
    Set objNTxBO = gobjContext.CreateInstance(App.Title & ".omTmNTxBO")
    objNTxBO.ImportBureauData vxmlRequestNode
    'BM0376 Call NTxBO.RunIncomeCalcs
    objNTxBO.RunIncomeCalcs vxmlRequestNode
    'BM0376 End
ImportBureauDataExit:
    Set objNTxBO = Nothing
    Set xmlResponseNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlStageTask = Nothing
    Set xmlCaseTask = Nothing
    Set xmlAttribute = Nothing
    Set xmlApplication = Nothing
    Set xmlCaseActivity = Nothing
    Set xmlNode = Nothing
    Set xmlTaskElem = Nothing
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
        xmlTempRequest.appendChild xmlApplication.cloneNode(True)
        CreateAdhocCaseTask xmlTempRequest
        lngErr = oeTmCCOKBureauImportFailed
    Else
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
    '<VSA> Visual Studio Analyser Support
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
''    Dim objMsgTmBO As Object
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
    Dim strCurrentStageID As String
    Dim xmlNodeList As IXMLDOMNodeList  'BM0177 MDC 20/12/2002
            
    'extract values from the XML
    strNewUserId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@NEWUSERID")
    strNewUnitId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@NEWUNITID")
    strOldUserId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@OLDUSERID")
    strOldUnitId = xmlGetMandatoryNodeText(vxmlRequestNode, "USERHISTORY/@OLDUNITID")
            
    'BUILD THE CREATE USER HISTORY REQUEST
    Set xmlCreateUserHistoryDOM = New FreeThreadedDOMDocument40
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
    GetCurrentStage vxmlRequestNode, xmlCurrentStageResponse
        
    strCurrentStageID = xmlGetMandatoryNodeText(xmlCurrentStageResponse, "RESPONSE/CASESTAGE/@STAGEID")
        
    'GET ALL THE STAGE TASKS FOR THE CURRENT STAGE
    'BUILD GETSTAGETASK REQUEST
    Set xmlGetStageTaskDOM = New FreeThreadedDOMDocument40
    'close the request node
    xmlGetStageTaskDOM.appendChild vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlGetStageTaskDOM.documentElement, "OPERATION", "GETSTAGETASKLIST"
    Set xmlStageTaskNode = xmlGetStageTaskDOM.createElement("STAGETASK")
    xmlGetStageTaskDOM.documentElement.appendChild xmlStageTaskNode
    xmlSetAttributeValue xmlStageTaskNode, "STAGEID", strCurrentStageID
    'Get the stagetask for this case task
    Set xmlGetStageTaskResponse = New FreeThreadedDOMDocument40
    RunTmRequest xmlGetStageTaskDOM, xmlGetStageTaskResponse
    'Create a DOM to hold the case tasks that are going to be transferred
    Set xmlTransferTaskListDOM = New FreeThreadedDOMDocument40
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
    Set xmlCurrentStageResponse = Nothing
    Set xmlStageTaskForCaseTaskId = Nothing
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
'SDS  BBG1741  10/12/2004__START
'PSC 06/01/2003 BM0089 - Start
' Functionality moved from MoveToNextStage
'Private Sub MoveToNextStageNewTx(ByVal vxmlRequestNode As IXMLDOMNode, _
'                                 ByVal vxmlResponseNode As IXMLDOMNode)
'    On Error GoTo MoveToNextStageNewTxExit
'    Const cstrFunctionName As String = "MoveToNextStageNewTx"
'    Dim xmlCurrentStageRequestNode As IXMLDOMNode
'    Dim xmlApplicationRequestNode As IXMLDOMNode
'    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
'    Dim xmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40
'    Dim xmlNextStageDoc As FreeThreadedDOMDocument40
'    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
'    Dim xmlNextCaseStageDoc As FreeThreadedDOMDocument40
'    Dim xmlCreateCaseStageDoc As FreeThreadedDOMDocument40
'    Dim xmlCurrentStageNode As IXMLDOMNode
'    Dim xmlNextStageNode As IXMLDOMNode
'    Dim xmlRbResponseNode As IXMLDOMNode
'    Dim xmlNextCaseStageNode As IXMLDOMNode
'    Dim objNTxBO As IOmTmNTxBO
'    'BS BM0136 07/02/2003
'    Dim xmlExceptionStageDoc As New FreeThreadedDOMDocument40
'    Dim xmlTaskNode As IXMLDOMNode
'    'BS BM0136 End 07/02/2003
'    ' validate request
'    Set xmlCurrentStageRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTSTAGE")
'    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "SOURCEAPPLICATION"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "CASEID"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "ACTIVITYID"
'    xmlCheckMandatoryAttribute xmlCurrentStageRequestNode, "STAGEID"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
'    ' get current CASESTAGE detail =====================================================
'    Set xmlCurrentStageDoc = xmlCreateDOMObject
'    'TK 14/09/2004 BBG1415 xmlCurrentStageDoc.async = False
'    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
'        errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
'    End If
'
'    ' get next CASESTAGE detail ========================================================
'    Set xmlNextStageDoc = xmlCreateDOMObject
'    'TK 14/09/2004 BBG1415 xmlNextStageDoc.async = False
'    Set xmlNextCaseStageNode = Nothing
'    Do While xmlNextCaseStageNode Is Nothing
'
'        ' get next STAGE in sequence ===================================================
'        GetNextStageInSequence _
'            vxmlRequestNode, _
'            xmlCurrentStageNode, _
'            xmlNextStageNode, _
'            xmlNextStageDoc
'        Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
'        If AnyRulesForStage(xmlNextStageNode) = True Then
'            If xmlApplicationDataDoc Is Nothing Then
'                ' get APPLICATION data =============================================================
'                Set xmlApplicationDataDoc = xmlCreateDOMObject
'                'TK 14/09/2004 BBG1415 xmlApplicationDataDoc.async = False
'                Set objNTxBO = gobjContext.CreateInstance("OmTm.OmTmNTxBO")
'                objNTxBO.GetApplicationData xmlApplicationRequestNode, xmlApplicationDataDoc, "TMRBTemplate"
'                Set objNTxBO = Nothing
'                Set xmlRbResponseNode = xmlApplicationDataDoc.selectSingleNode("RESPONSE")
'            End If
'            ' run rules ========================================================================
'            If xmlNextCaseStageDoc Is Nothing Then
'                Set xmlNextCaseStageDoc = xmlCreateDOMObject
'                'TK 14/09/2004 BBG1415 xmlNextCaseStageDoc.async = False
'            Else
'                xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
'            End If
'            RunUserRules _
'                vxmlRequestNode, _
'                xmlNextStageNode, _
'                xmlRbResponseNode, _
'                xmlNextCaseStageDoc
'            Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'            'BS BM0136 07/02/2003
'            'If the case has been auto-declined/cancelled then the exception stageid and exception reason
'            'attributes will be returned in the CASESTAGE node of xmlNextCaseStageDoc
'            If Not xmlNextCaseStageNode Is Nothing Then
'                'Next stage is set. If it is an exception stage then change the CURRENTSTAGE node to CASESTAGE
'                'and copy the STAGEID and EXCEPTIONREASON attributes on the requestNode, re-populate
'                'xmlNextStageNode and xmlNextCaseStageNode with the exception stage/task details
'                'Is it an exception stage
'                If xmlAttributeValueExists(xmlNextCaseStageNode, "EXCEPTIONREASON") Then
'                    'Rename the CURRENTSTAGE node on the RequestNode to CASESTAGE and copy the
'                    'STAGEID and EXCEPTIONREASON attributes
'                    xmlSetAttributeValue vxmlRequestNode.selectSingleNode("CURRENTSTAGE"), "STAGEID", xmlGetAttributeText(xmlNextCaseStageNode, "STAGEID")
'                    xmlSetAttributeValue vxmlRequestNode.selectSingleNode("CURRENTSTAGE"), "EXCEPTIONREASON", xmlGetAttributeText(xmlNextCaseStageNode, "EXCEPTIONREASON")
'                    xmlChangeNodeName vxmlRequestNode, "CURRENTSTAGE", "CASESTAGE"
'                    'Get exception stage/task details
'                    Set xmlExceptionStageDoc = xmlCreateDOMObject
'                    GetRequestedStage vxmlRequestNode, xmlExceptionStageDoc
'                    Set xmlNextStageNode = Nothing
'                    Set xmlNextStageNode = xmlExceptionStageDoc.selectSingleNode("RESPONSE/STAGE")
'                    If xmlNextStageNode Is Nothing Then
'                         errThrowError cstrFunctionName, oeTmNoNextStage
'                    End If
'                    'Double-Check it is an exception stage, if not then error
'                    If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
'                        errThrowError cstrFunctionName, oeTmNotExceptionStage
'                    End If
'                    ' remove non DEFAULT tasks
'                    For Each xmlTaskNode In xmlNextStageNode.childNodes
'                        If xmlGetAttributeAsBoolean(xmlTaskNode, "DEFAULTFLAG") = False Then
'                            xmlNextStageNode.removeChild xmlTaskNode
'                        End If
'                    Next
'                    'Create CaseStage/CaseTask details
'                    If xmlNextCaseStageDoc Is Nothing Then
'                        Set xmlNextCaseStageDoc = xmlCreateDOMObject
'                        'TK 14/09/2004 BBG1415 xmlNextCaseStageDoc.async = False
'                    Else
'                        xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
'                    End If
'                    CreateCaseStageFromStage _
'                        xmlNextStageNode, _
'                        xmlNextCaseStageDoc
'                    Set xmlNextCaseStageNode = Nothing
'                    Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
'                End If
'            End If
'            'BS BM0136 End 07/02/2003
'        Else
'            If xmlNextCaseStageDoc Is Nothing Then
'                Set xmlNextCaseStageDoc = xmlCreateDOMObject
'                'TK 14/09/2004 BBG1415 xmlNextCaseStageDoc.async = False
'            Else
'                xmlNextCaseStageDoc.removeChild xmlNextCaseStageDoc.firstChild
'            End If
'            CreateCaseStageFromStage _
'                xmlNextStageNode, _
'                xmlNextCaseStageDoc
'            Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
'        End If
'    Loop
'    ' find next CASESTAGE in rules response
'    If xmlNextCaseStageNode Is Nothing Then
'        errThrowError cstrFunctionName, oeTmNoNextStage
'    End If
'    ' prepare CreateCaseStage request ==================================================
'    Set xmlCreateCaseStageDoc = xmlCreateDOMObject
'    PrepareCreateCaseStageRequest _
'        vxmlRequestNode, _
'        xmlCurrentStageNode, _
'        xmlNextCaseStageNode, _
'        xmlCreateCaseStageDoc
'    AddDefaultValuesToCaseStage _
'        vxmlRequestNode, _
'        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
'        xmlNextStageNode
'    ' prepare for update to current stage ==============================================
'    ' add incomplete tasks from current stage to next stage ============================
'    If Not xmlCurrentStageNode Is Nothing Then
'
'        'BS BM0136 07/02/2003
'        'Do not check authority level if case auto-declined/cancelled
'        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
'            ' first check we are auth'd to create the next stage   SYS1788 JLD
'            Dim xmlValidateDoc As FreeThreadedDOMDocument40
'            Set xmlValidateDoc = New FreeThreadedDOMDocument40
'            Dim xmlElem As IXMLDOMElement
'            Dim xmlNode As IXMLDOMNode
'            ' clone initial request node
'            Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
'            ' set operation name
'            xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
'            Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
'            Set xmlNode = xmlNode.appendChild(xmlElem)
'            ' copy STAGEID from next case stage
'            xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
'            RunTmRequest xmlValidateDoc
'        End If
'        'BS BM0136 End 07/02/2003
'
'        Set xmlUpdateCurrentStageDoc = xmlCreateDOMObject
'        PrepareCurrentStageUpdateRequest _
'            vxmlRequestNode, _
'            xmlCurrentStageNode, _
'            xmlUpdateCurrentStageDoc
'        'BS BM0136 07/02/2003
'        'It is now possible for the next stage to be an exception stage (if it was auto-declined/cancelled).
'        'If it isn't an exception stage then carry forward tasks as before, else cancel tasks
''
''        ProcessCarryForwardTasks _
''            xmlCurrentStageNode, _
''            xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
''            xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False _
'        Then
'            ProcessCarryForwardTasks _
'                xmlCurrentStageNode, _
'                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
'                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'        Else
'            ProcessCancelTasks _
'                xmlCurrentStageNode, _
'                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'        End If
'        'BS BM0136 End 07/02/2003
'
'        ' run both requests as a composite transaction
'        Dim xmlUpdateNodes As Collection
'        Set xmlUpdateNodes = New Collection
'        xmlUpdateNodes.Add xmlUpdateCurrentStageDoc.firstChild
'        xmlUpdateNodes.Add xmlCreateCaseStageDoc.firstChild
'        ' PSC 27/11/01 SYS3221 - Start
'        RunCompositeRequest xmlUpdateNodes
'   'AQR SYS2266 DRC
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
'    End If
'    ' PSC 27/11/01 SYS3221 - End
'    Set objNTxBO = Nothing
'
'    ' call omApp.ApplicationBO to create application stage on omiga ====================
'    CreateApplicationStage _
'        vxmlRequestNode, _
'        xmlApplicationRequestNode, _
'        xmlNextStageNode, _
'        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'    'BM0340 MDC 22/04/2003
'    ' get New current (i.e. updated) CASESTAGE detail =====================================================
''    Set xmlCurrentStageDoc = New FreeThreadedDOMDocument40
''    xmlCurrentStageDoc.async = False
''    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
''
''    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
''    vxmlResponseNode.appendChild xmlCurrentStageNode
'    'BM0340 MDC 22/04/2003 - End
'
'MoveToNextStageNewTxExit:
'    Set xmlCurrentStageRequestNode = Nothing
'    Set xmlApplicationRequestNode = Nothing
'    Set xmlCurrentStageNode = Nothing
'    Set xmlUpdateCurrentStageDoc = Nothing
'    Set xmlNextStageNode = Nothing
'    Set xmlRbResponseNode = Nothing
'    Set xmlNextCaseStageNode = Nothing
'    Set xmlCurrentStageDoc = Nothing
'    'TK 14/09/2004 BBG1415 Set xmlUpdateCurrentStageDoc = Nothing
'    Set xmlNextStageDoc = Nothing
'    Set xmlApplicationDataDoc = Nothing
'    Set xmlNextCaseStageDoc = Nothing
'    Set xmlCreateCaseStageDoc = Nothing
'    Set xmlValidateDoc = Nothing
'    Set xmlElem = Nothing
'    Set xmlNode = Nothing
'    'BS BM0136 07/02/2003
'    Set xmlExceptionStageDoc = Nothing
'    Set xmlTaskNode = Nothing
'    Set objNTxBO = Nothing
'    'BS BM0136 End 07/02/2003
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub
'Private Sub MoveToStageNewTx(ByVal vxmlRequestNode As IXMLDOMNode, _
'                             ByVal vxmlResponseNode As IXMLDOMNode)
'' Functionality moved from MoveToStage
'
'    On Error GoTo MoveToStageNewTxExit
'    Const cstrFunctionName As String = "MoveToStageNewTx"
'    Dim xmlCurrentStageDoc As FreeThreadedDOMDocument40
'    Dim xmlNextStageDoc As FreeThreadedDOMDocument40
'    Dim xmlNextCaseStageDoc As FreeThreadedDOMDocument40
'    Dim xmlCreateCaseStageDoc As FreeThreadedDOMDocument40
'    Dim xmlUpdateCurrentStageDoc As FreeThreadedDOMDocument40
'    Dim xmlApplicationDataDoc As FreeThreadedDOMDocument40
'    Dim xmlRbResponseNode As IXMLDOMNode
'    Dim xmlNode As IXMLDOMNode
'    Dim xmlApplicationRequestNode As IXMLDOMNode
'    Dim xmlCurrentStageNode As IXMLDOMNode
'    Dim xmlNextStageNode As IXMLDOMNode
'    Dim xmlNextCaseStageNode As IXMLDOMNode
'    Dim xmlRequestCaseStageNode As IXMLDOMNode
'    Dim objNTxBO As IOmTmNTxBO
'    Dim objCommonBO As IomTmCommonBO
'    ' validate request
'    Set xmlRequestCaseStageNode = xmlGetMandatoryNode(vxmlRequestNode, "CASESTAGE")
'    Set xmlApplicationRequestNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "SOURCEAPPLICATION"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "CASEID"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "ACTIVITYID"
'    xmlCheckMandatoryAttribute xmlRequestCaseStageNode, "STAGEID"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONNUMBER"
'    xmlCheckMandatoryAttribute xmlApplicationRequestNode, "APPLICATIONFACTFINDNUMBER"
'    ' get requested CASESTAGE detail =====================================================
'    Set xmlNextStageDoc = New FreeThreadedDOMDocument40
'    GetRequestedStage vxmlRequestNode, xmlNextStageDoc
'    Set xmlNextStageNode = xmlNextStageDoc.selectSingleNode("RESPONSE/STAGE")
'    If xmlNextStageNode Is Nothing Then
'        errThrowError cstrFunctionName, oeTmNoNextStage
'    End If
'    ' remove non DEFAULT tasks
'    For Each xmlNode In xmlNextStageNode.childNodes
'        If xmlGetAttributeAsBoolean(xmlNode, "DEFAULTFLAG") = False Then
'            xmlNextStageNode.removeChild xmlNode
'        End If
'    Next
'    ' get current CASESTAGE detail =====================================================
'    Set xmlCurrentStageDoc = xmlCreateDOMObject
'    'TK 14/09/2004 BBG1415 xmlCurrentStageDoc.async = False
'    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False Then
'        If IncompleteMandatoryStageTasks(xmlCurrentStageNode) Then
'            errThrowError cstrFunctionName, oeTmMandatoryTasksOutstanding
'        End If
'    End If
'
'    If AnyRulesForStage(xmlNextStageNode) = True Then
'        'get APPLICATION data
'        Set xmlApplicationDataDoc = New FreeThreadedDOMDocument40
'        xmlApplicationDataDoc.async = False
'        Set objCommonBO = gobjContext.CreateInstance(App.Title & ".OmTmCommonBO")
'        objCommonBO.GetApplicationData xmlApplicationRequestNode, xmlApplicationDataDoc, "TMRBTemplate"
'        Set objCommonBO = Nothing
'        Set xmlRbResponseNode = xmlApplicationDataDoc.selectSingleNode("RESPONSE")
'        ' Run rules
'        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
'        xmlNextCaseStageDoc.async = False
'        RunUserRules _
'            vxmlRequestNode, _
'            xmlNextStageNode, _
'            xmlRbResponseNode, _
'            xmlNextCaseStageDoc
'        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'
'    Else
'
'        ' create CASESTAGE from STAGE
'        Set xmlNextCaseStageDoc = New FreeThreadedDOMDocument40
'        xmlNextCaseStageDoc.async = False
'        CreateCaseStageFromStage _
'            xmlNextStageNode, _
'            xmlNextCaseStageDoc
'        Set xmlNextCaseStageNode = xmlNextCaseStageDoc.selectSingleNode("CASESTAGE")
'    End If
'    ' stage not applicable
'    If xmlNextCaseStageNode Is Nothing Then
'        errThrowError cstrFunctionName, oeTmStageNotApplicable
'    End If
'    ' prepare CreateCaseStage request ==================================================
'    Set xmlCreateCaseStageDoc = New FreeThreadedDOMDocument40
'    xmlCreateCaseStageDoc.async = False
'    PrepareCreateCaseStageRequest _
'        vxmlRequestNode, _
'        xmlCurrentStageNode, _
'        xmlNextCaseStageNode, _
'        xmlCreateCaseStageDoc
'    AddDefaultValuesToCaseStage _
'        vxmlRequestNode, _
'        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
'        xmlNextStageNode
'    ' prepare for update to current stage ==============================================
'    If Not xmlCurrentStageNode Is Nothing Then
'
'        ' first check we are auth'd to create the next stage SYS1788 JLD
'        Dim xmlValidateDoc As FreeThreadedDOMDocument40
'        Set xmlValidateDoc = New FreeThreadedDOMDocument40
'        Dim xmlElem As IXMLDOMElement
'        ' clone initial request node
'        Set xmlNode = xmlValidateDoc.appendChild(vxmlRequestNode.cloneNode(False))
'        ' set operation name
'        xmlNode.Attributes.getNamedItem("OPERATION").Text = "ValidateStageAuthority"
'        Set xmlElem = xmlValidateDoc.createElement("CASESTAGE")
'        Set xmlNode = xmlNode.appendChild(xmlElem)
'        ' copy STAGEID from next case stage
'        xmlCopyAttribute xmlNextCaseStageNode, xmlNode, "STAGEID"
'        RunTmRequest xmlValidateDoc
'        Set xmlUpdateCurrentStageDoc = New FreeThreadedDOMDocument40
'        xmlUpdateCurrentStageDoc.async = False
'        PrepareCurrentStageUpdateRequest _
'            vxmlRequestNode, _
'            xmlCurrentStageNode, _
'            xmlUpdateCurrentStageDoc
'        If xmlGetAttributeAsBoolean(xmlNextStageNode, "EXCEPTIONSTAGEINDICATOR") = False _
'        Then
'            ProcessCarryForwardTasks _
'                xmlCurrentStageNode, _
'                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE"), _
'                xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'        Else
'            ProcessCancelTasks _
'                xmlCurrentStageNode, _
'                xmlUpdateCurrentStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'        End If
'        ' run both requests as a composite transaction
'        Dim xmlUpdateNodes As Collection
'        Set xmlUpdateNodes = New Collection
'        xmlUpdateNodes.Add xmlUpdateCurrentStageDoc.firstChild
'        xmlUpdateNodes.Add xmlCreateCaseStageDoc.firstChild
'        ' PSC 27/11/01 SYS3221 - Start
'        RunCompositeRequest xmlUpdateNodes
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
'    End If
'    ' PSC 27/11/01 SYS3221 - End
'    Set objNTxBO = Nothing
'    ' call omApp.ApplicationBO to create application stage on omiga ====================
'    CreateApplicationStage _
'        vxmlRequestNode, _
'        xmlApplicationRequestNode, _
'        xmlNextStageNode, _
'        xmlCreateCaseStageDoc.selectSingleNode("REQUEST/CASESTAGE")
'    ' get New current (i.e. updated) CASESTAGE detail =====================================================
'    Set xmlCurrentStageDoc = xmlCreateDOMObject
'    'TK 14/09/2004 BBG1415 xmlCurrentStageDoc.async = False
'    GetCurrentStage vxmlRequestNode, xmlCurrentStageDoc
'    Set xmlCurrentStageNode = xmlCurrentStageDoc.selectSingleNode("RESPONSE/CASESTAGE")
'    vxmlResponseNode.appendChild xmlCurrentStageNode
'MoveToStageNewTxExit:
'    Set xmlNode = Nothing
'    Set xmlRequestCaseStageNode = Nothing
'    Set xmlApplicationRequestNode = Nothing
'    Set xmlCurrentStageNode = Nothing
'    Set xmlNextStageNode = Nothing
'    Set xmlNextCaseStageNode = Nothing
'    Set xmlCurrentStageDoc = Nothing
'    Set xmlCreateCaseStageDoc = Nothing
'    Set xmlNextStageDoc = Nothing
'    Set xmlNextCaseStageDoc = Nothing
'    Set xmlUpdateCurrentStageDoc = Nothing
'    Set xmlApplicationDataDoc = Nothing
'    Set xmlValidateDoc = Nothing
'    Set xmlElem = Nothing
'    Set xmlRbResponseNode = Nothing
'    errCheckError cstrFunctionName, TypeName(Me)
'End Sub
'PSC 06/01/2003 BM0089 - End
'SDS  BBG1741  10/12/2004__END
'BM0340 MDC 17/04/2003
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
'BM0340 MDC 17/04/2003 - End
'TK 01/06/2004 BBG406 - Start
Private Sub RunTransactCreditSearch(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo RunTransactCreditSearchExit
    Const cstrFunctionName As String = "RunTransactCreditSearch"
    
    Dim xmlNewRequest       As IXMLDOMElement
    Dim xmlCaseTaskElem     As IXMLDOMElement
    Dim objBO               As Object
    Dim xmlElem             As IXMLDOMElement
    Dim strResponse         As String
    Dim lngErrNo            As Long
'    Dim strErrDescription   As String
    Dim strValidationType   As String
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATECREDITSEARCH"
    Set xmlElem = xmlNewRequest.ownerDocument.createElement("REQUESTSOURCE")
    xmlSetAttributeValue xmlElem, "SOURCE", "OMIGA"
    xmlNewRequest.appendChild xmlElem

    Set xmlElem = xmlNewRequest.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlElem, "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode(".//APPLICATION/@APPLICATIONNUMBER").Text
    xmlSetAttributeValue xmlElem, "APPLICATIONFACTFINDNUMBER", vxmlRequestNode.selectSingleNode(".//APPLICATION/@APPLICATIONFACTFINDNUMBER").Text
    xmlSetAttributeValue xmlElem, "CUSTOMERNUMBER", vxmlRequestNode.selectSingleNode(".//CASETASK/@CUSTOMERIDENTIFIER").Text
    xmlSetAttributeValue xmlElem, "ADDRESSGUID", vxmlRequestNode.selectSingleNode(".//CASETASK/@CONTEXT").Text
    xmlNewRequest.appendChild xmlElem
    
    Set objBO = gobjContext.CreateInstance("omTransact.omTransactBO")
    strResponse = objBO.RunTransactDecision(xmlNewRequest.xml)
    
    On Error Resume Next

    lngErrNo = errCheckXMLResponse(strResponse, False)
    'TK 13/08/2004 BBG406
    On Error GoTo RunTransactCreditSearchExit
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, 4813
    End If
    
    ' Get the task status for validation type "CP"
    If lngErrNo = 0 Then 'TK 13/08/2004 BBG406
        strValidationType = GetFirstComboValueId("TaskStatus", "CP")
    End If
    
    ' Now update Case Task
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "UPDATECASETASK"
    Set xmlCaseTaskElem = vxmlRequestNode.selectSingleNode(".//CASETASK").cloneNode(True)
    xmlSetAttributeValue xmlCaseTaskElem, "USERID", xmlGetAttributeText(vxmlRequestNode, "USERID")
    xmlSetAttributeValue xmlCaseTaskElem, "UNITID", xmlGetAttributeText(vxmlRequestNode, "UNITID")
    xmlSetAttributeValue xmlCaseTaskElem, "USERAUTHORITYLEVEL", xmlGetAttributeText(vxmlRequestNode, "USERAUTHORITYLEVEL")
    xmlSetAttributeValue xmlCaseTaskElem, "TASKSTATUS", strValidationType
    xmlNewRequest.appendChild xmlCaseTaskElem
    
    Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    Exit Sub
    
RunTransactCreditSearchExit:
    Set xmlElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    Set objBO = Nothing
    errCheckError cstrFunctionName

End Sub
'TK 01/06/2004 BBG406 - End

'TK 27/07/2004 E2EM00000113 - Start
Private Sub RunTransactFullDecision(ByVal vxmlRequestNode As IXMLDOMNode)

    On Error GoTo RunTransactFullDecisionExit
    Const cstrFunctionName As String = "RunTransactFullDecision"
    
    Dim xmlNewRequest       As IXMLDOMElement
    Dim xmlCaseTaskElem     As IXMLDOMElement
    Dim objBO               As Object
    Dim xmlElem             As IXMLDOMElement
    Dim strResponse         As String
    Dim lngErrNo            As Long
    Dim strErrDescription   As String
    Dim strValidationType   As String
    
    Set xmlNewRequest = xmlGetRequestNode(vxmlRequestNode)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEFULLDECISION"
    Set xmlElem = xmlNewRequest.ownerDocument.createElement("REQUESTSOURCE")
    xmlSetAttributeValue xmlElem, "SOURCE", "OMIGA"
    xmlNewRequest.appendChild xmlElem

    Set xmlElem = xmlNewRequest.ownerDocument.createElement("APPLICATION")
    xmlSetAttributeValue xmlElem, "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode(".//APPLICATION/@APPLICATIONNUMBER").Text
    xmlSetAttributeValue xmlElem, "APPLICATIONFACTFINDNUMBER", vxmlRequestNode.selectSingleNode(".//APPLICATION/@APPLICATIONFACTFINDNUMBER").Text
    xmlNewRequest.appendChild xmlElem
    
    Set objBO = gobjContext.CreateInstance("omTransact.omTransactBO")
    strResponse = objBO.RunTransactDecision(xmlNewRequest.xml)
    
    On Error Resume Next
    errCheckXMLResponse strResponse, True
    lngErrNo = Err.Number
    strErrDescription = Err.Description
    On Error GoTo RunTransactFullDecisionExit
    If lngErrNo <> 0 Then
        errThrowError cstrFunctionName, oeTmTransactFullDecisionFailed, strErrDescription
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
    
    Set objBO = gobjContext.CreateInstance(gstrMsgTm_COMPONENT & ".MsgTmBO")
    strResponse = objBO.TmRequest(xmlNewRequest.xml)
    errCheckXMLResponse strResponse, True
    
    Exit Sub
    
RunTransactFullDecisionExit:
    Set xmlElem = Nothing
    Set xmlNewRequest = Nothing
    Set xmlCaseTaskElem = Nothing
    Set objBO = Nothing
    errCheckError cstrFunctionName

End Sub
'TK 27/07/2004 E2EM00000113 - End

'PSC 20/08/2004 BBG1233 - Start

Private Sub OverturnKFI(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo OverturnKFIExit
    Const cstrFunctionName As String = "OverturnKFI"

    Dim xmlThisResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlThisCaseTaskNode As IXMLDOMNode
    Dim xmlPrintAttributeNode As IXMLDOMNode
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlPrinterNode As IXMLDOMNode
    Dim xmlAppFactFindNode As IXMLDOMNode
    
    Dim objAppBO As Object
    Dim gobjEnquiry As Object
    
    Dim strResponse As String
    Dim strAcceptedQuoteNumber As String
    Dim strActiveQuoteNumber As String
    Dim strQuoteNumber As String
    Dim strPrinterDestination As String
    Dim strOutputDocument As String
    
        ' check mandatory items
    xmlCheckMandatoryAttribute vxmlRequestNode, "UNITID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERID"
    xmlCheckMandatoryAttribute vxmlRequestNode, "USERAUTHORITYLEVEL"
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    xmlCheckMandatoryAttribute xmlApplicationNode, "APPLICATIONNUMBER"
    xmlCheckMandatoryAttribute xmlApplicationNode, "APPLICATIONFACTFINDNUMBER"
    Set xmlCaseTaskNode = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "CASEID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlCaseTaskNode, "TASKID"
    strOutputDocument = xmlGetMandatoryAttributeText(xmlCaseTaskNode, "OUTPUTDOCUMENT", 10011)

    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")

    Set xmlThisRequestDoc = xmlCreateElementRequestFromNode(vxmlRequestNode, "APPLICATION", True)
    Set objAppBO = gobjContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    strResponse = objAppBO.GetApplicationData(xmlThisRequestDoc.xml)
    Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
    
    Set xmlAppFactFindNode = xmlGetMandatoryNode(xmlThisResponseDoc, "RESPONSE/APPLICATIONLATESTDETAILS/APPLICATION/APPLICATIONFACTFIND")
    
    strAcceptedQuoteNumber = xmlGetNodeText(xmlAppFactFindNode, "ACCEPTEDQUOTENUMBER")
    strActiveQuoteNumber = xmlGetNodeText(xmlAppFactFindNode, "ACTIVEQUOTENUMBER")
    
    If Len(Trim$(strAcceptedQuoteNumber)) <> 0 Then
        strQuoteNumber = strAcceptedQuoteNumber
    ElseIf Len(Trim$(strActiveQuoteNumber)) <> 0 Then
        strQuoteNumber = strActiveQuoteNumber
    Else
        errThrowError cstrFunctionName, 10012
    End If
    
    On Error Resume Next
    Set gobjEnquiry = gobjContext.CreateInstance(gstrPRINT_COMPONENT & ".omPrintBO")
    ' PSC 07/02/02 SYS3995
    On Error GoTo OverturnKFIExit
    If Not gobjEnquiry Is Nothing Then
        
        Set xmlThisCaseTaskNode = xmlThisRequestDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlThisCaseTaskNode, "OPERATION", "GetPrintAttributes"
                  
        Set xmlTempNode = xmlThisRequestDoc.createElement("FINDATTRIBUTES")
        xmlThisCaseTaskNode.appendChild xmlTempNode
        xmlSetAttributeValue xmlTempNode, "HOSTTEMPLATEID", strOutputDocument
          
        strResponse = gobjEnquiry.OmRequest(xmlThisCaseTaskNode.xml)
        Set xmlThisResponseDoc = xmlLoad(strResponse, cstrFunctionName)
        errCheckXMLResponseNode xmlThisResponseDoc.documentElement, , True
        Set xmlPrintAttributeNode = xmlThisResponseDoc.selectSingleNode("RESPONSE/ATTRIBUTES")
        
        If xmlGetAttributeText(xmlPrintAttributeNode, "INACTIVEINDICATOR", "") = "1" Then
          errThrowError cstrFunctionName, 4820, "", "OverturnKFI"
        End If
        
        If Len(Trim$(xmlGetAttributeText(xmlPrintAttributeNode, "DEFAULTCOPIES", ""))) = 0 Then
          errThrowError cstrFunctionName, 4823, "", "OverturnKFI"
        End If
        
        strPrinterDestination = xmlGetAttributeText(xmlPrintAttributeNode, "PRINTERDESTINATIONTYPE")
        'strPrinterDestination = GetValidationTypeForValueID("PrinterDestination", CInt(strPrinterDestination))
        strPrinterDestination = GetComboText("PrinterDestination", CInt(strPrinterDestination))
        
         If ((InStr(1, strPrinterDestination, "L", vbTextCompare) > 0) _
            And (InStr(1, strPrinterDestination, "R", vbTextCompare) > 0) _
            And (InStr(1, strPrinterDestination, "W", vbTextCompare) > 0)) Then
        'If (strPrinterDestination <> "L" And _
        '    strPrinterDestination <> "R" And _
        '    strPrinterDestination <> "W") Then
           errThrowError cstrFunctionName, 4822, strPrinterDestination, "OverTurnKFI"
        End If
        
        Set xmlThisCaseTaskNode = xmlThisRequestDoc.createElement("REQUEST")
        xmlSetAttributeValue xmlThisCaseTaskNode, "OPERATION", "PrintDocumentForTask"
        
        xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERID"
        xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "UNITID"
        xmlCopyAttribute vxmlRequestNode, xmlThisCaseTaskNode, "USERAUTHORITYLEVEL"
        xmlThisCaseTaskNode.appendChild xmlCaseTaskNode.cloneNode(True)
        
        Set xmlTempNode = xmlThisRequestDoc.createElement("APPLICATION")
        xmlThisCaseTaskNode.appendChild xmlTempNode
        xmlCopyAttribute xmlApplicationNode, xmlTempNode, "APPLICATIONFACTFINDNUMBER"
        xmlSetAttributeValue xmlTempNode, "PRINTKFI", "1"
        xmlSetAttributeValue xmlTempNode, "QUOTATIONNUMBER", strQuoteNumber
        
        Set xmlTempNode = xmlThisRequestDoc.createElement("PRINTATTRIBUTES")
        xmlThisCaseTaskNode.appendChild xmlTempNode
        xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DPSTEMPLATEID", "DPSDOCUMENTID"
        xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "DEFAULTCOPIES", "COPIES"
        xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "PDMMETHOD", "METHODNAME"
        xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "RECIPIENTTYPE"
        xmlCopyAttribute xmlPrintAttributeNode, xmlTempNode, "DELIVERYTYPE"
        xmlSetAttributeValue xmlTempNode, "DESTINATIONTYPE", Left$(strPrinterDestination, 1)
        Set xmlPrinterNode = vxmlRequestNode.selectSingleNode("PRINTER")
        xmlCopyAttributeValue xmlPrinterNode, xmlTempNode, "COMPRESSIONMETHOD", "COMPRESSIONMETHOD"
        If strPrinterDestination = "Local" Or strPrinterDestination = "Workstation Printer" Then
            xmlCopyAttributeValue xmlPrinterNode, xmlTempNode, "PRINTERNAME", "PRINTER"
        Else
            xmlCopyAttributeValue xmlPrintAttributeNode, xmlTempNode, "REMOTEPRINTERLOCATION", "PRINTER"
        End If
        
        ' PSC 06/10/2004 BBG1572
        xmlSetAttributeValue xmlTempNode, "WEBDOCUMENTTYPE", "KFI"

        xmlThisCaseTaskNode.appendChild xmlTempNode
        PrintDocumentForTask xmlThisCaseTaskNode, vxmlResponseNode
            
    End If

OverturnKFIExit:

    Set xmlThisResponseDoc = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisCaseTaskNode = Nothing
    Set xmlPrintAttributeNode = Nothing
    Set xmlTempNode = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlAppFactFindNode = Nothing
    Set xmlPrinterNode = Nothing
    Set objAppBO = Nothing
    Set gobjEnquiry = Nothing
 
    errCheckError cstrFunctionName, TypeName(Me)

End Sub
'PSC 20/08/2004 BBG1233 - End

' SR 19/01/2006 : BBGRb74 - new method 'UpdateCaseTaskReferenceNumber'
Private Sub UpdateCaseTaskReferenceNumber( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateCaseTaskReferenceNumberErr
Const cstrFunctionName As String = "UpdateCaseTaskReferenceNumber"
    
    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskDataDoc As FreeThreadedDOMDocument40, xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlThisRequestNode As IXMLDOMNode, xmlCaseTask As IXMLDOMNode, xmlTemp As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim strCaseTaskName As String 'SR BBGRb604
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.async = False
    ' clone request node and set opetation to get the CaseTask records
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetCaseTasks"
    
    Set xmlTemp = vxmlRequestNode.selectSingleNode("CASETASK")
    Set xmlCaseTask = xmlThisRequestDoc.createElement("CASETASK")
    xmlCopyAttribute xmlTemp, xmlCaseTask, "SOURCEAPPLICATION"
    xmlCopyAttribute xmlTemp, xmlCaseTask, "CASEID"
    xmlCopyAttribute xmlTemp, xmlCaseTask, "ACTIVITYID"
    xmlCopyAttribute xmlTemp, xmlCaseTask, "ACTIVITYINSTANCE"
    xmlCopyAttribute xmlTemp, xmlCaseTask, "PRODUCTSEQUENCENUMBER"
    xmlCopyAttribute xmlTemp, xmlCaseTask, "BUSINESSAREATYPE"
     
    xmlThisRequestNode.appendChild xmlCaseTask
    Set xmlCaseTaskDataDoc = New FreeThreadedDOMDocument40
    xmlCaseTaskDataDoc.async = False
    
    RunTmRequest xmlThisRequestDoc, xmlCaseTaskDataDoc
    
    'Update CaseTask table now
    xmlThisRequestDoc.loadXML ("")
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "UpdateCaseTask"
    
    For Each xmlNode In xmlCaseTaskDataDoc.selectNodes("//CASETASK")
        Set xmlCaseTask = xmlThisRequestDoc.createElement("CASETASK")
        xmlCopyAttribute xmlNode, xmlCaseTask, "CASEACTIVITYGUID"
        xmlCopyAttribute xmlNode, xmlCaseTask, "SOURCEAPPLICATION"
        xmlCopyAttribute xmlNode, xmlCaseTask, "CASEID"
        xmlCopyAttribute xmlNode, xmlCaseTask, "ACTIVITYID"
        xmlCopyAttribute xmlNode, xmlCaseTask, "ACTIVITYINSTANCE"
        xmlCopyAttribute xmlNode, xmlCaseTask, "STAGEID"
        xmlCopyAttribute xmlNode, xmlCaseTask, "CASESTAGESEQUENCENO"
        xmlCopyAttribute xmlNode, xmlCaseTask, "TASKID"
        xmlCopyAttribute xmlNode, xmlCaseTask, "TASKINSTANCE"
        xmlCopyAttribute xmlTemp, xmlCaseTask, "REFERENCENUMBER"

        'SR BBGRb604 - build the task name and set it
        strCaseTaskName = xmlGetAttributeText(xmlTemp, "REFERENCENUMBER") & "_" & _
                          xmlGetAttributeText(xmlNode, "OMIGAPRODUCTTYPE") & "_" & _
                          xmlGetAttributeText(xmlNode, "TASKID") & "_" & _
                          xmlGetAttributeText(xmlNode, "OWNINGCUSTOMERS")
        strCaseTaskName = IIf(Len(strCaseTaskName) > 100, Left(strCaseTaskName, 100), strCaseTaskName)
        xmlSetAttributeValue xmlCaseTask, "CASETASKNAME", strCaseTaskName
                
        xmlThisRequestNode.appendChild xmlCaseTask.cloneNode(True)
    Next xmlNode

    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    xmlResponseDoc.async = False
    RunTmRequest xmlThisRequestDoc, xmlResponseDoc
    errCheckXMLResponseNode xmlResponseDoc.documentElement, vxmlResponseNode, True
    
    gobjContext.SetComplete
UpdateCaseTaskReferenceNumberExit:
    Set xmlThisRequestDoc = Nothing
    Set xmlCaseTaskDataDoc = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlCaseTask = Nothing
    Set xmlTemp = Nothing
    Set xmlNode = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
UpdateCaseTaskReferenceNumberErr:
    gobjContext.SetAbort
    GoTo UpdateCaseTaskReferenceNumberExit
End Sub

' SR 19/01/2006 : BBGRb74 - new method 'AddDefaultProductTasks'
Private Sub AddDefaultProductTasks( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
On Error GoTo AddDefaultProductTasksErr
Const cstrFunctionName As String = "AddDefaultProductTasks"

    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskIn As IXMLDOMNode, xmlProtProdIn As IXMLDOMNode
    Dim xmlTempResponse As IXMLDOMNode, xmlTasks As IXMLDOMNode
    Dim xmlNewRequest As IXMLDOMNode, xmlNode As IXMLDOMNode, xmlCaseTask As IXMLDOMNode
    
    Dim strBusinessAreaType As String, strInTrustInd As String, strTaskId As String
    'BBGRb1298
    Dim xmlStageTaskRequestDoc As FreeThreadedDOMDocument40
    Dim xmlStageTaskDetailDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskRequestNode As IXMLDOMNode
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlStageTaskDetailNode As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim strTaskPattern As String
    Dim strTask As String
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlCaseTaskIn = vxmlRequestNode.selectSingleNode(".//CASETASK")
    strBusinessAreaType = xmlGetMandatoryAttributeText(xmlCaseTaskIn, "BUSINESSAREATYPE")
    Dim strCaseTakName As String    'SR BBGRb604
    
    Set xmlTempResponse = xmlDoc.createElement("RESPONSE")
    If strBusinessAreaType = "M" Then
        GetAllGlobalBandedParamValuesAsXml "POSTMMtgeOnSubmitTasks", "STRING", xmlTempResponse
        Set xmlTasks = xmlTempResponse
    Else
        If strBusinessAreaType = "G" Or strBusinessAreaType = "F" Or strBusinessAreaType = "I" Then
            Set xmlProtProdIn = xmlGetMandatoryNode(vxmlRequestNode, "//PROTECTIONPRODUCTDETAIS")
            strInTrustInd = xmlGetMandatoryAttributeText(xmlProtProdIn, "INTRUSTINDICATOR")
            GetAllGlobalBandedParamValuesAsXml "POSTMLifeIncOnSubmitTasks", "STRING", xmlTempResponse
            
            If strInTrustInd = "1" Then
                strTaskId = GetGlobalParamString("POSTMLifeIncWITSubmitTask")
                If Len(strTaskId) > 0 Then
                    Set xmlNode = xmlDoc.createElement("GLOBALPARAM")
                    xmlSetAttributeValue xmlNode, "NAME", "POSTMLifeIncWITSubmitTasks"
                    xmlSetAttributeValue xmlNode, "STRING", strTaskId
                    xmlTempResponse.appendChild xmlNode
                End If
            End If
            Set xmlTasks = xmlTempResponse
        End If
    End If
    
    'Now create casetask records for all those tasks retrieved above
    Set xmlNewRequest = vxmlRequestNode.cloneNode(False)
    xmlSetAttributeValue xmlNewRequest, "OPERATION", "CREATEADHOCCASETASK"
    xmlNewRequest.appendChild xmlGetMandatoryNode(vxmlRequestNode, "//APPLICATION")
    Set xmlCaseTask = xmlCaseTaskIn.cloneNode(False)
    
    Set xmlCaseTaskRequestNode = vxmlRequestNode.selectSingleNode("CASETASK")

    'BBGRb1298 create GetStageTaskDetail request
    Set xmlStageTaskRequestDoc = xmlCreateDOMObject

    ' clone request node
    Set xmlThisRequestNode = xmlStageTaskRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set OPERATION attribute
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageTaskDetail"
    ' create STAGETASK node
    Set xmlElem = xmlStageTaskRequestDoc.createElement("STAGETASK")
    xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "STAGEID"
    xmlThisRequestNode.appendChild xmlElem
    ' run TmRequest
    Set xmlStageTaskDetailDoc = xmlCreateDOMObject
    RunTmRequest xmlStageTaskRequestDoc, xmlStageTaskDetailDoc
    
    For Each xmlNode In xmlTasks.childNodes
        strTaskId = xmlGetAttributeText(xmlNode, "STRING")
        xmlSetAttributeValue xmlCaseTask, "TASKID", strTaskId
        If (strBusinessAreaType <> "M") Then
            'SR BBGRb604
            If (Len(xmlGetAttributeText(xmlCaseTaskIn, "PROVIDER")) > 0) Then
                If (Len(xmlGetAttributeText(xmlCaseTaskIn, "REFERENCENUMBER")) > 0) Then
                    strCaseTakName = xmlGetAttributeText(xmlCaseTaskIn, "REFERENCENUMBER")
                Else
                    strCaseTakName = xmlGetAttributeText(xmlCaseTaskIn, "PRODUCTSEQUENCENUMBER")
                End If
            End If
            strCaseTakName = strCaseTakName & "_" & xmlGetAttributeText(xmlCaseTaskIn, "OMIGAPRODUCTTYPE") & "_" & _
                             xmlGetAttributeText(xmlNode, "STRING") & "_" & _
                             xmlGetAttributeText(xmlCaseTaskIn, "OWNINGCUSTOMERS")
            If (Len(strCaseTakName) > 100) Then
               strCaseTakName = Left(strCaseTakName, 100)
            End If
            'SR BBGRb604 - End
            xmlSetAttributeValue xmlCaseTask, "CASETASKNAME", strCaseTakName
            
        End If
            
        'BBGRb1298 Add the stage task detail
        strTaskPattern = "RESPONSE/STAGETASK[@TASKID='" & strTaskId & "']"
        Set xmlStageTaskDetailNode = xmlStageTaskDetailDoc.selectSingleNode(strTaskPattern)
        If Not xmlStageTaskDetailNode Is Nothing Then
            If xmlGetAttributeText(xmlStageTaskDetailNode, "MANDATORYFLAG") <> "" Then
               xmlSetAttributeValue xmlCaseTask, "MANDATORYFLAG", xmlGetAttributeText(xmlStageTaskDetailNode, "MANDATORYFLAG")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "CARRYFORWARDINDICATOR") <> "" Then
                xmlSetAttributeValue xmlCaseTask, "CARRYFORWARDINDICATOR", xmlGetAttributeText(xmlStageTaskDetailNode, "CARRYFORWARDINDICATOR")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "TRIGGERTASKIND") <> "" Then
                xmlSetAttributeValue xmlCaseTask, "TRIGGERTASKIND", xmlGetAttributeText(xmlStageTaskDetailNode, "TRIGGERTASKIND")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "USERAUTHORITYLEVEL") <> "" Then
                xmlSetAttributeValue xmlCaseTask, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlStageTaskDetailNode, "USERAUTHORITYLEVEL")
            End If
        End If

        xmlNewRequest.appendChild xmlCaseTask.cloneNode(False)
    Next xmlNode
   
    
    CreateAdhocCaseTask xmlNewRequest
    gobjContext.SetComplete
AddDefaultProductTasksExit:
    Set xmlDoc = Nothing
    Set xmlCaseTaskIn = Nothing
    Set xmlProtProdIn = Nothing
    Set xmlTempResponse = Nothing
    Set xmlTasks = Nothing
    'BBGRb1298
    Set xmlElem = Nothing
    Set xmlCaseTaskRequestNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlStageTaskDetailNode = Nothing
    Set xmlStageTaskRequestDoc = Nothing
    Set xmlStageTaskDetailDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
AddDefaultProductTasksErr:
    gobjContext.SetAbort
    GoTo AddDefaultProductTasksExit
End Sub

' SR 23/01/2006 : BBGRb74 - new method 'TransferApplicationOwnershipOnSubmit'
Private Sub TransferApplicationOwnershipOnSubmit( _
                    ByVal vxmlRequestNode As IXMLDOMNode, _
                    ByVal vxmlResponseNode As IXMLDOMNode)
                    
On Error GoTo TransferApplicationOwnershipOnSubmitErr
Const cstrFunctionName As String = "TransferApplicationOwnershipOnSubmit"
                    
    Dim Cn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim adoStr As ADODB.Stream
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskIn As IXMLDOMNode
    Dim xmlNewReq As IXMLDOMElement, xmlUserHistory As IXMLDOMNode
    
    Dim objCRUDTxBO As Object
    
    Dim strNewUnitId As String, strNewUserId As String
    Dim strUnitId As String, strBusinessType As String

    Set xmlCaseTaskIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
    strUnitId = xmlGetAttributeText(vxmlRequestNode, "UNITID")
    strBusinessType = xmlGetAttributeText(xmlCaseTaskIn, "BUSINESSAREATYPE")
    
    Set Cn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Cn.ConnectionString = adoGetDbConnectString()
    Cn.Open
    cmd.ActiveConnection = Cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_GetActionOwnerForLinkedBrokerSupportUnit"

    cmd.Parameters.Append cmd.CreateParameter("@p_UnitId", adBSTR, adParamInput, 10, strUnitId)
    cmd.Parameters.Append cmd.CreateParameter("@p_BusinessAreaType", adBSTR, adParamInput, 1, strBusinessType)
    
    Set adoStr = New ADODB.Stream
    adoStr.Open
    cmd.Properties("Output Stream") = adoStr
    cmd.Properties("XML Root") = "RESPONSE"

    cmd.Execute , , adExecuteStream
    
    Dim strTemp As String
    strTemp = adoStr.ReadText
    If strTemp <> "" Then
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.async = False
        xmlDoc.loadXML strTemp
                
        Set xmlNode = xmlDoc.selectSingleNode("RESPONSE/ACTIONOWNER")
        strNewUnitId = xmlGetAttributeText(xmlNode, "UNITID")
        strNewUserId = xmlGetAttributeText(xmlNode, "USERID")
    End If
    
    If Len(strNewUnitId) <> 0 And Len(strNewUserId) <> 0 Then
        'Now transfer the applicationOwnership to new user id/ unit id
        Set xmlNewReq = vxmlRequestNode.cloneNode(False)
        xmlNewReq.setAttribute "CRUD_OP", "CREATE"
        xmlNewReq.setAttribute "ENTITY_REF", "USERHISTORY"
        
        Set xmlUserHistory = xmlDoc.createElement("USERHISTORY")
        xmlSetAttributeValue xmlUserHistory, "APPLICATIONNUMBER", xmlGetAttributeText(xmlCaseTaskIn, "CASEID")
        xmlSetAttributeValue xmlUserHistory, "UNITID", strNewUnitId
        xmlSetAttributeValue xmlUserHistory, "USERID", strNewUserId
        xmlSetSysDateToNodeAttrib xmlUserHistory, "USERHISTORYDATE"
        xmlNewReq.appendChild xmlUserHistory
        
        Set objCRUDTxBO = gobjContext.CreateInstance("omCRUD.omCRUDTxBO")
        strTemp = objCRUDTxBO.OmRequest(xmlNewReq.xml)
        Set xmlDoc = xmlLoad(strTemp, cstrFunctionName)
        errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    End If

    gobjContext.SetComplete
    
TransferApplicationOwnershipOnSubmitExit:

    Set Cn = Nothing
    Set cmd = Nothing
    Set adoStr = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlCaseTaskIn = Nothing
    Set objCRUDTxBO = Nothing
    Set xmlNewReq = Nothing
    Set xmlUserHistory = Nothing
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
TransferApplicationOwnershipOnSubmitErr:
    gobjContext.SetAbort
    GoTo TransferApplicationOwnershipOnSubmitExit
End Sub

' SR 25/01/2006 : BBGRb74 - new method 'SubmitProduct'
Private Sub SubmitProduct(ByVal vxmlRequestNode As IXMLDOMNode, _
                            ByVal vxmlResponseNode As IXMLDOMNode)
                            
On Error GoTo SubmitProductErr
Const cstrFunctionName As String = "SubmitProduct"

    Dim xmlNewReq As IXMLDOMElement, xmlAFF As IXMLDOMElement, xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40

    Dim objCRUDBO As Object, objCRUDTxBO As Object
    
    Dim strTempResp As String, strTaskId As String, strBusinessAreaType As String
    Dim intSubmitFlag As Integer
    'BBGRb1298
    Dim xmlStageTaskRequestDoc As FreeThreadedDOMDocument40
    Dim xmlStageTaskDetailDoc As FreeThreadedDOMDocument40
    Dim xmlCaseTaskRequestNode As IXMLDOMNode
    Dim xmlThisRequestNode As IXMLDOMNode
    Dim xmlStageTaskDetailNode As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim strTaskPattern As String
    Dim strTask As String
    
    'Determine whether this is the first product to be submitted (for this application)
    Set xmlDoc = New FreeThreadedDOMDocument40
    Set xmlNode = vxmlRequestNode.selectSingleNode("//APPLICATION")
    
    Set xmlNewReq = vxmlRequestNode.cloneNode(False)
    xmlNewReq.setAttribute "CRUD_OP", "READ"
    xmlNewReq.setAttribute "ENTITY_REF", "APPLICATIONFACTFIND"
    
    
    Set xmlAFF = xmlDoc.createElement("APPLICATIONFACTFIND")
    xmlCopyAttribute xmlNode, xmlAFF, "APPLICATIONNUMBER"
    xmlCopyAttribute xmlNode, xmlAFF, "APPLICATIONFACTFINDNUMBER"
    xmlNewReq.appendChild xmlAFF

    Set objCRUDBO = gobjContext.CreateInstance("omCRUD.omCRUDBO")
    Set objCRUDTxBO = gobjContext.CreateInstance("omCRUD.omCRUDTxBO")
    strTempResp = objCRUDBO.OmRequest(xmlNewReq.xml)
    Set xmlDoc = xmlLoad(strTempResp, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    
    Set xmlAFF = xmlDoc.selectSingleNode("//APPLICATIONFACTFIND")
    intSubmitFlag = xmlGetAttributeAsInteger(xmlAFF, "SUBMITFLAG")
    If intSubmitFlag <> 1 Then
    
        'BBGRb1298 Need to get the stage task detail for our adhoc tasks
        Set xmlCaseTaskRequestNode = vxmlRequestNode.selectSingleNode("CASETASK")
    
        ' create GetStageTaskDetail request ==================================================
        Set xmlStageTaskRequestDoc = xmlCreateDOMObject
    
        ' clone request node
        Set xmlThisRequestNode = xmlStageTaskRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
        ' set OPERATION attribute
        xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = "GetStageTaskDetail"
        ' create STAGETASK node
        Set xmlElem = xmlStageTaskRequestDoc.createElement("STAGETASK")
        xmlCopyAttribute xmlCaseTaskRequestNode, xmlElem, "STAGEID"
        xmlThisRequestNode.appendChild xmlElem
        ' run TmRequest
        Set xmlStageTaskDetailDoc = xmlCreateDOMObject
        RunTmRequest xmlStageTaskRequestDoc, xmlStageTaskDetailDoc
    
        Set xmlNewReq = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlNewReq, "OPERATION", "TransferApplicationOwnerOnSubmit"
        xmlNewReq.appendChild vxmlRequestNode.selectSingleNode("//CASETASK").cloneNode(False)
        TransferApplicationOwnershipOnSubmit xmlNewReq, vxmlResponseNode
        
        'Now update the SubmitFlag of ApplicationFactFind
        Set xmlNewReq = vxmlRequestNode.cloneNode(False)
        xmlNewReq.setAttribute "CRUD_OP", "UPDATE"
        xmlNewReq.setAttribute "ENTITY_REF", "APPLICATIONFACTFIND"
        xmlSetAttributeValue xmlAFF, "SUBMITFLAG", "1"
        
        xmlNewReq.appendChild xmlAFF
        strTempResp = objCRUDTxBO.OmRequest(xmlNewReq.xml)
        Set xmlDoc = xmlLoad(strTempResp, cstrFunctionName)
        errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True

        'Create task 'POSTMWelcomeTask'
        strTaskId = GetGlobalParamString("POSTMWelcomeTask")
        Set xmlNewReq = vxmlRequestNode.cloneNode(False)
        xmlSetAttributeValue xmlNewReq, "OPERATION", "CREATEADHOCCASETASK"
        
        xmlNewReq.appendChild xmlGetMandatoryNode(vxmlRequestNode, "//APPLICATION").cloneNode(False)
        
        Set xmlNode = vxmlRequestNode.selectSingleNode("//CASETASK").cloneNode(False)
        xmlSetAttributeValue xmlNode, "TASKID", strTaskId
        
        'BBGRb1298 Add the stage task detail
        strTaskPattern = "RESPONSE/STAGETASK[@TASKID='" & strTaskId & "']"
        Set xmlStageTaskDetailNode = xmlStageTaskDetailDoc.selectSingleNode(strTaskPattern)
        If Not xmlStageTaskDetailNode Is Nothing Then
            If xmlGetAttributeText(xmlStageTaskDetailNode, "MANDATORYFLAG") <> "" Then
               xmlSetAttributeValue xmlNode, "MANDATORYFLAG", xmlGetAttributeText(xmlStageTaskDetailNode, "MANDATORYFLAG")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "CARRYFORWARDINDICATOR") <> "" Then
                xmlSetAttributeValue xmlNode, "CARRYFORWARDINDICATOR", xmlGetAttributeText(xmlStageTaskDetailNode, "CARRYFORWARDINDICATOR")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "TRIGGERTASKIND") <> "" Then
                xmlSetAttributeValue xmlNode, "TRIGGERTASKIND", xmlGetAttributeText(xmlStageTaskDetailNode, "TRIGGERTASKIND")
            End If
            If xmlGetAttributeText(xmlStageTaskDetailNode, "USERAUTHORITYLEVEL") <> "" Then
                xmlSetAttributeValue xmlNode, "USERAUTHORITYLEVEL", xmlGetAttributeText(xmlStageTaskDetailNode, "USERAUTHORITYLEVEL")
            End If
        End If
        
        xmlNewReq.appendChild xmlNode
        
        CreateAdhocCaseTask xmlNewReq
    End If
    
    Set xmlNewReq = vxmlRequestNode.cloneNode(True)
    xmlSetAttributeValue xmlNewReq, "OPERATION", "AddDefaultProductTasks"
    'SR BBGRb604
    Set xmlNode = vxmlRequestNode.selectSingleNode(".//CASETASK")
    strBusinessAreaType = xmlGetMandatoryAttributeText(xmlNode, "BUSINESSAREATYPE")
    If strBusinessAreaType <> "P" Then 'SR BBGRb604 - End
        AddDefaultProductTasks xmlNewReq, vxmlResponseNode
    End If
    
    gobjContext.SetComplete
SubmitProductExit:
    Set xmlNewReq = Nothing
    Set xmlAFF = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing

    Set objCRUDBO = Nothing
    Set objCRUDTxBO = Nothing
    'BBGRb1298
    Set xmlElem = Nothing
    Set xmlCaseTaskRequestNode = Nothing
    Set xmlThisRequestNode = Nothing
    Set xmlStageTaskDetailNode = Nothing
    Set xmlStageTaskRequestDoc = Nothing
    Set xmlStageTaskDetailDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
SubmitProductErr:
    gobjContext.SetAbort
    GoTo SubmitProductExit
End Sub

' SR 26/01/2006 : BBGRb144 - new method 'UpdateAppStatusAndSubmitProduct'
Private Sub UpdateAppStatusAndSubmitProduct(ByVal vxmlRequestNode As IXMLDOMNode, _
                            ByVal vxmlResponseNode As IXMLDOMNode)
                            
On Error GoTo UpdateAppStatusAndSubmitProductErr
Const cstrFunctionName As String = "UpdateAppStatusAndSubmitProduct"
    
    UpdateApplicationStatus vxmlRequestNode
    SubmitProduct vxmlRequestNode, vxmlResponseNode
    
    gobjContext.SetComplete
UpdateAppStatusAndSubmitProductExit:
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
UpdateAppStatusAndSubmitProductErr:
    
    gobjContext.SetAbort
    GoTo UpdateAppStatusAndSubmitProductExit
End Sub

' SR 26/01/2006 : BBGRb144 - new method 'UpdateAppStatus'
Private Sub UpdateApplicationStatus(ByVal vxmlRequestNode As IXMLDOMNode)
                            
On Error GoTo UpdateApplicationStatusErr
Const cstrFunctionName As String = "UpdateApplicationStatus"
    
    Dim xmlNode As IXMLDOMNode
    Dim Cn As ADODB.Connection
    Dim cmd As ADODB.Command
    
    Dim strApplNumber As String, strAFFNumber As String, strBussAreaType As String
    Dim intPackageNumber As Integer, intProdSeqNumber As Integer, intNewStatus As Integer
    Dim intStatusType As Integer, intOldstatus As Integer
    Dim strUserId As String, strUnitId As String
        
    Set xmlNode = vxmlRequestNode.selectSingleNode("//PRODUCT")
    strApplNumber = xmlGetAttributeText(xmlNode, "APPLICATIONNUMBER")
    strAFFNumber = xmlGetAttributeText(xmlNode, "APPLICATIONFACTFINDNUMBER")
    intPackageNumber = xmlGetAttributeAsInteger(xmlNode, "PACKAGENUMBER")
    strBussAreaType = xmlGetAttributeText(xmlNode, "BUSINESSAREATYPE")
    intProdSeqNumber = xmlGetAttributeAsInteger(xmlNode, "PRODUCTSEQUENCENUMBER")
    
    Set xmlNode = vxmlRequestNode.selectSingleNode("//STATUSHISTORY")
    intStatusType = xmlGetAttributeAsInteger(xmlNode, "STATUSTYPE")
    intOldstatus = xmlGetAttributeAsInteger(xmlNode, "OLDSTATUS")
    intNewStatus = xmlGetAttributeAsInteger(xmlNode, "NEWSTATUS")
        
    strUserId = xmlGetAttributeText(vxmlRequestNode, "USERID")
    strUnitId = xmlGetAttributeText(vxmlRequestNode, "UNITID")
        
    Set Cn = New ADODB.Connection
    Set cmd = New ADODB.Command
    Cn.ConnectionString = adoGetDbConnectString()
    Cn.Open
    cmd.ActiveConnection = Cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_UpdateProductAndApplicationStatus"
            
    cmd.Parameters.Append cmd.CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, 12, strApplNumber)
    cmd.Parameters.Append cmd.CreateParameter("@p_AFFNumber", adInteger, adParamInput, 5, strAFFNumber)
    cmd.Parameters.Append cmd.CreateParameter("@p_PackageNumber", adInteger, adParamInput, 5, intPackageNumber)
    cmd.Parameters.Append cmd.CreateParameter("@p_BusinessAreaType", adBSTR, adParamInput, 1, strBussAreaType)
    cmd.Parameters.Append cmd.CreateParameter("@p_ProductSequenceNumber", adInteger, adParamInput, 5, intProdSeqNumber)
    cmd.Parameters.Append cmd.CreateParameter("@p_StatusType", adInteger, adParamInput, 5, intStatusType)
    cmd.Parameters.Append cmd.CreateParameter("@p_NewStatus", adInteger, adParamInput, 5, intNewStatus)
    cmd.Parameters.Append cmd.CreateParameter("@p_OldStatus", adInteger, adParamInput, 5, intOldstatus)
    cmd.Parameters.Append cmd.CreateParameter("@p_userId", adBSTR, adParamInput, 10, strUserId)
    cmd.Parameters.Append cmd.CreateParameter("@p_UnitId", adBSTR, adParamInput, 10, strUnitId)
    
    cmd.Execute

UpdateApplicationStatusExit:
    Set xmlNode = Nothing
    Set Cn = Nothing
    Set cmd = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    Exit Sub
UpdateApplicationStatusErr:
    gobjContext.SetAbort
    GoTo UpdateApplicationStatusExit
                            
End Sub


